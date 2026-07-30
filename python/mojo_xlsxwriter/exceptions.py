class XlsxWriterException(Exception):
    pass


class InvalidWorksheetName(XlsxWriterException):
    pass


class DuplicateWorksheetName(XlsxWriterException):
    pass


class OverlappingRange(XlsxWriterException):
    pass


class FileCreateError(XlsxWriterException):
    pass
