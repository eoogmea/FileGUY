from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException
from pdf4me.model import Extract, ExtractResources


class ExtractClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def extract(self, extract):
        """The predefined extraction is carried out.

        :param extract: extraction configuration
        :type extract: Extract
        :return: ExtractRes, contains extracted pages in form of a PDF
        """

        # check extract validity
        self.__check_extract_object_validity(extract)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=extract,
                                                                   controller='Extract/Extract')

        return res

    def extract_pages(self, page_nrs, file):
        """The chosen pages will be extracted from the given PDF and form a new PDF.

        :param page_nrs: the page numbers which will be extracted, number 1 corresponds to the first page.
        :type page_nrs: str: comma-separated page numbers (no spaces)
        :param file: to extract the pages from
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('pageNrs', page_nrs)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='Extract/ExtractPages')

    def extract_resources(self, extract_resources):
        """The predefined extraction is carried out.

        :param extract_resources: extraction configuration
        :type extract_resources: ExtractResources
        :return: ExtractResourcesRes, contains extracted resource info
        """

        # check extract_resources validity
        self.__check_extract_resources_object_validity(extract_resources)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=extract_resources,
                                                                   controller='Extract/ExtractResources')

        return res

    def __check_extract_object_validity(self, extract):
        """Checks whether the extract object contains the essential information to be
        processed by the server."""

        if extract is None:
            raise Pdf4meClientException('The extract parameter cannot be None.')
        elif extract.document is None or extract.document.doc_data is None:
            raise Pdf4meClientException('The extract document cannot be None.')
        elif extract.extract_action is None:
            raise Pdf4meClientException('The extract_action cannot be None.')
        elif extract.extract_action.extract_pages is None or len(extract.extract_action.extract_pages) == 0:
            raise Pdf4meClientException('The extract_pages of extract_action cannot be None or empty.')

    def __check_extract_resources_object_validity(self, extract_resources):
        """Checks whether the extract_resources object contains the essential information to be
        processed by the server."""

        if extract_resources is None:
            raise Pdf4meClientException('The extract_resources parameter cannot be None.')
        elif extract_resources.document is None or extract_resources.document.doc_data is None:
            raise Pdf4meClientException('The extract_resources document cannot be None.')
        elif extract_resources.extract_resources_action is None:
            raise Pdf4meClientException('The extract_resources_action cannot be None.')
