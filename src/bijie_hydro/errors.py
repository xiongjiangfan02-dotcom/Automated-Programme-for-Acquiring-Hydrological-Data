class CollectorError(Exception):
    """采集模块可预期错误基类。"""


class ConfigurationError(CollectorError):
    pass


class SourceError(CollectorError):
    pass


class AccessRestrictedError(SourceError):
    """Public endpoint rejected access or requested a rate limit pause."""


class DecodeError(CollectorError):
    pass


class QualityError(CollectorError):
    pass


class TemplateError(CollectorError):
    pass

