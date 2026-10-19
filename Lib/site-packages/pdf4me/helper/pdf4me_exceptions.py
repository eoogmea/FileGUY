class Pdf4meClientException(Exception):
    """Raised when wrong arguments have been provided."""

    def __init__(self, msg):
        self.msg = msg
        Exception.__init__(self, msg)


class Pdf4meBackendException(Exception):
    """Raised when an error on the back-end occurred."""

    def __init__(self, msg):
        self.msg = msg
        Exception.__init__(self, msg)
