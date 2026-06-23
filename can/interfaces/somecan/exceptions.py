# coding: utf-8

from can import CanError


class SmError(CanError):
    def __init__(self, sm_error_code, error_string, function):
        self.sm_error_code = sm_error_code
        text = "%s failed (%s)" % (function, error_string)
        super(SmError, self).__init__(text)
