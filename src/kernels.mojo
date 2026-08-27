"""Batched shared-string kernels for XLSX serialization."""

from max.algorithm import parallelize
from std.sys.info import simd_width_of

comptime BPtr = UnsafePointer[UInt8, AnyOrigin[mut=True]]
comptime IPtr = UnsafePointer[Int64, AnyOrigin[mut=True]]
comptime W = simd_width_of[DType.float64]()
comptime PARALLEL_THRESHOLD = 16384
comptime ESCAPE_CHUNK = 4096
comptime MAX_WORKERS = 8


def strings_equal(data: BPtr, offsets: IPtr, first: Int, second: Int) -> Bool:
    var first_start = Int(offsets[first])
    var second_start = Int(offsets[second])
    var length = Int(offsets[first + 1] - offsets[first])
    if length != Int(offsets[second + 1] - offsets[second]):
        return False
    var j = 0
    while j + W <= length:
        var first_values = data.load[width=W, alignment=1](first_start + j)
        var second_values = data.load[width=W, alignment=1](second_start + j)
        if first_values != second_values:
            return False
        j += W
    while j < length:
        if data[first_start + j] != data[second_start + j]:
            return False
        j += 1
    return True


def string_hash(data: BPtr, offsets: IPtr, index: Int) -> UInt64:
    var value = UInt64(1469598103934665603)
    for j in range(Int(offsets[index]), Int(offsets[index + 1])):
        value = (value ^ UInt64(data[j])) * UInt64(1099511628211)
    return value


def is_hex(value: UInt8) -> Bool:
    return (
        (value >= UInt8(48) and value <= UInt8(57))
        or (value >= UInt8(65) and value <= UInt8(70))
        or (value >= UInt8(97) and value <= UInt8(102))
    )


def is_excel_escape(data: BPtr, position: Int, end: Int) -> Bool:
    return (
        position + 7 <= end
        and data[position] == UInt8(95)
        and data[position + 1] == UInt8(120)
        and is_hex(data[position + 2])
        and is_hex(data[position + 3])
        and is_hex(data[position + 4])
        and is_hex(data[position + 5])
        and data[position + 6] == UInt8(95)
    )


@always_inline
def ordinary_block(data: BPtr, position: Int) -> Bool:
    var values = data.load[width=W, alignment=1](position)
    var special = (
        values.le(UInt8(31))
        | values.eq(UInt8(38))
        | values.eq(UInt8(60))
        | values.eq(UInt8(62))
        | values.eq(UInt8(95))
        | values.eq(UInt8(239))
    )
    return special.cast[DType.uint8]().reduce_add() == 0


def put_literal(dst: BPtr, position: Int, literal: StringSlice) -> Int:
    var result = position
    for i in range(literal.byte_length()):
        dst[result] = literal.as_bytes()[i]
        result += 1
    return result


def put_control_escape(dst: BPtr, position: Int, value: UInt8) -> Int:
    var result = put_literal(dst, position, "_x00")
    var high = Int(value) // 16
    var low = Int(value) % 16
    dst[result] = UInt8(48 + high) if high < 10 else UInt8(55 + high)
    dst[result + 1] = UInt8(48 + low) if low < 10 else UInt8(55 + low)
    dst[result + 2] = UInt8(95)
    return result + 3


def escaped_length(data: BPtr, start: Int, end: Int) -> Int:
    var position = start
    var result = 0
    while position < end:
        if position + W <= end and ordinary_block(data, position):
            result += W
            position += W
            continue
        var value = data[position]
        if is_excel_escape(data, position, end):
            result += 13
            position += 7
        elif (
            value <= UInt8(8)
            or value == UInt8(11)
            or value == UInt8(12)
            or value == UInt8(13)
            or (value >= UInt8(14) and value <= UInt8(31))
        ):
            result += 7
            position += 1
        elif value == UInt8(38):
            result += 5
            position += 1
        elif value == UInt8(60) or value == UInt8(62):
            result += 4
            position += 1
        elif (
            position + 3 <= end
            and value == UInt8(239)
            and data[position + 1] == UInt8(191)
            and (
                data[position + 2] == UInt8(190)
                or data[position + 2] == UInt8(191)
            )
        ):
            result += 7
            position += 3
        else:
            result += 1
            position += 1
    return result


def escape_one(data: BPtr, start: Int, end: Int, dst: BPtr, output_start: Int):
    var position = start
    var result = output_start
    while position < end:
        if position + W <= end and ordinary_block(data, position):
            dst.store[alignment=1](
                result, data.load[width=W, alignment=1](position)
            )
            result += W
            position += W
            continue
        var value = data[position]
        if is_excel_escape(data, position, end):
            result = put_literal(dst, result, "_x005F")
            for j in range(7):
                dst[result] = data[position + j]
                result += 1
            position += 7
        elif (
            value <= UInt8(8)
            or value == UInt8(11)
            or value == UInt8(12)
            or value == UInt8(13)
            or (value >= UInt8(14) and value <= UInt8(31))
        ):
            result = put_control_escape(dst, result, value)
            position += 1
        elif value == UInt8(38):
            result = put_literal(dst, result, "&amp;")
            position += 1
        elif value == UInt8(60):
            result = put_literal(dst, result, "&lt;")
            position += 1
        elif value == UInt8(62):
            result = put_literal(dst, result, "&gt;")
            position += 1
        elif (
            position + 3 <= end
            and value == UInt8(239)
            and data[position + 1] == UInt8(191)
            and (
                data[position + 2] == UInt8(190)
                or data[position + 2] == UInt8(191)
            )
        ):
            if data[position + 2] == UInt8(190):
                result = put_literal(dst, result, "_xFFFE_")
            else:
                result = put_literal(dst, result, "_xFFFF_")
            position += 3
        else:
            dst[result] = value
            result += 1
            position += 1


@export("mxw_dedup_utf8")
def mxw_dedup_utf8(
    data_addr: Int,
    data_length: Int,
    offsets_addr: Int,
    offsets_length: Int,
    count: Int,
    slots_addr: Int,
    capacity: Int,
    slots_length: Int,
    ids_addr: Int,
    ids_length: Int,
) abi("C") -> Int:
    if (
        data_addr == 0
        or offsets_addr == 0
        or slots_addr == 0
        or ids_addr == 0
        or data_length < 0
        or count < 0
        or offsets_length != count + 1
        or ids_length < count
        or capacity < 2
        or slots_length < capacity
        or (capacity & (capacity - 1)) != 0
        or capacity < count
    ):
        return -1
    var data = BPtr(unsafe_from_address=data_addr)
    var offsets = IPtr(unsafe_from_address=offsets_addr)
    var slots = IPtr(unsafe_from_address=slots_addr)
    var ids = IPtr(unsafe_from_address=ids_addr)
    if offsets[0] != 0 or offsets[count] != Int64(data_length):
        return -2
    for i in range(count):
        if offsets[i] < 0 or offsets[i + 1] < offsets[i]:
            return -2
    var unique_count = 0
    for i in range(count):
        var slot = Int(string_hash(data, offsets, i) & UInt64(capacity - 1))
        while slots[slot] != -1:
            var previous = Int(slots[slot])
            if strings_equal(data, offsets, previous, i):
                ids[i] = ids[previous]
                break
            slot = (slot + 1) & (capacity - 1)
        if slots[slot] == -1:
            slots[slot] = Int64(i)
            ids[i] = Int64(unique_count)
            unique_count += 1
    return unique_count


@export("mxw_escape_xml")
def mxw_escape_xml(
    data_addr: Int,
    data_length: Int,
    offsets_addr: Int,
    offsets_length: Int,
    count: Int,
    escaped_offsets_addr: Int,
    escaped_offsets_length: Int,
    dst_addr: Int,
    dst_capacity: Int,
) abi("C") -> Int:
    if (
        data_addr == 0
        or offsets_addr == 0
        or escaped_offsets_addr == 0
        or dst_addr == 0
        or data_length < 0
        or count < 0
        or offsets_length != count + 1
        or escaped_offsets_length < count + 1
        or dst_capacity < 0
    ):
        return -1
    var data = BPtr(unsafe_from_address=data_addr)
    var offsets = IPtr(unsafe_from_address=offsets_addr)
    var escaped_offsets = IPtr(unsafe_from_address=escaped_offsets_addr)
    var dst = BPtr(unsafe_from_address=dst_addr)
    if offsets[0] != 0 or offsets[count] != Int64(data_length):
        return -2
    for index in range(count):
        if offsets[index] < 0 or offsets[index + 1] < offsets[index]:
            return -2
    var result = 0
    escaped_offsets[0] = 0

    @parameter
    @__copy_capture(data, offsets, escaped_offsets, count)
    def measure_chunk(chunk: Int):
        var begin = chunk * ESCAPE_CHUNK
        var end = min(count, begin + ESCAPE_CHUNK)
        for index in range(begin, end):
            escaped_offsets[index + 1] = Int64(
                escaped_length(data, Int(offsets[index]), Int(offsets[index + 1]))
            )

    var chunks = (count + ESCAPE_CHUNK - 1) // ESCAPE_CHUNK
    if count >= PARALLEL_THRESHOLD:
        parallelize[measure_chunk](chunks, min(chunks, MAX_WORKERS))
    else:
        for chunk in range(chunks):
            measure_chunk(chunk)
    for index in range(count):
        result += Int(escaped_offsets[index + 1])
        escaped_offsets[index + 1] = Int64(result)
    if result > dst_capacity:
        return -3

    @parameter
    @__copy_capture(data, offsets, escaped_offsets, dst, count)
    def escape_chunk(chunk: Int):
        var begin = chunk * ESCAPE_CHUNK
        var end = min(count, begin + ESCAPE_CHUNK)
        for index in range(begin, end):
            escape_one(
                data,
                Int(offsets[index]),
                Int(offsets[index + 1]),
                dst,
                Int(escaped_offsets[index]),
            )

    if count >= PARALLEL_THRESHOLD:
        parallelize[escape_chunk](chunks, min(chunks, MAX_WORKERS))
    else:
        for chunk in range(chunks):
            escape_chunk(chunk)
    return result
