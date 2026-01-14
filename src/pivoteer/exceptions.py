"""Custom exception hierarchy for pivoteer."""


class PivoteerError(Exception):
    """Base exception for pivoteer errors."""


class TemplateNotFoundError(PivoteerError):
    """Raised when the input template cannot be located or opened."""


class TableNotFoundError(PivoteerError):
    """Raised when the target Excel table cannot be resolved."""


class XmlStructureError(PivoteerError):
    """Raised when expected XML structures are missing or malformed."""


class InvalidDataError(PivoteerError):
    """Raised when the supplied data cannot be injected."""


class WriteError(PivoteerError):
    """Raised when writing the output file fails."""


class PivotCacheError(PivoteerError):
    """Raised when pivot cache metadata cannot be updated."""
