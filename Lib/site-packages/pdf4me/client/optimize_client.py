from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException
from pdf4me.model import Optimize


class OptimizeClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def optimize(self, optimize):
        """The predefined optimization is carried out.

        :param optimize: optimization configuration
        :type optimize: Optimize
        :return: OptimizeRes, contains optimized PDF
        """

        # check optimize validity
        self.__check_optimize_object_validity(optimize)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=optimize,
                                                                   controller='Optimize/Optimize')

        return res

    def optimize_by_profile(self, profile, file):
        """The provided file will be optimized w.r.t the chosen profile. For instance if the profile 'max' has been
        selected, the resulting PDF will be optimized for maximal memory size reduction.

        :param profile: optimization profile (e.g. 'max' for maximal compression)
        :type profile: str
        :param file: to be optimized
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('profile', profile)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='Optimize/OptimizeByProfile')

    def __check_optimize_object_validity(self, optimize):
        """Checks whether the optimize object contains the essential information to be
                processed by the server."""

        if optimize is None:
            raise Pdf4meClientException('The optimize parameter cannot be None.')
        elif optimize.document is None or optimize.document.doc_data is None:
            raise Pdf4meClientException('The optimize document cannot be None.')
        elif optimize.optimize_action is None:
            raise Pdf4meClientException('The optimize_action cannot be None.')
        elif optimize.optimize_action.use_profile != True:
            raise Pdf4meClientException('The use_profile parameter of optimize_action has to be set to true.')
