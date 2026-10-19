import base64

from pdf4me.helper.json_converter import JsonConverter
from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException
from pdf4me.helper.response_checker import ResponseChecker
from pdf4me.model import Split


class SplitClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def split(self, split):
        """The predefined split is carried out.

        :param split: split configuration
        :type split: Split
        :return: ExtractRes, contains the document splinters of the original PDF document
        """

        # check split validity
        self.__check_split_object_validity(split)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=split,
                                                                   controller='Split/Split')

        return res

    def split_by_page_nr(self, page_nr, file):
        """Splits the PDF after the pageNr, this results in two smaller PDF documents.

        :param page_nr: determines after which page the split takes place
        :type page_nr: str:
        :param file: to split into two
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('pageNr', page_nr)]

        res = self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                          controller='Split/SplitByPageNr')

        # get json
        res = JsonConverter().load(res)

        # extract the two documents
        pdf_1 = base64.b64decode(res[0])
        pdf_2 = base64.b64decode(res[1])

        return pdf_1, pdf_2

    def split_recurring(self, page_nr, file):
        """Splits the PDF after the pageNr, this results in two smaller PDF documents.

        :param page_nr: determines after which page the split takes place
        :type page_nr: str:
        :param file: to split into two
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('pageNr', page_nr)]

        res = self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                          controller='Split/SplitRecurring')

        # get json
        res = JsonConverter().load(res)

        # extract the two documents
        pdfs = []
        for pdf in res:
            pdfs.append(base64.b64decode(pdf))
        # pdf_1 = base64.b64decode(res[0])
        # pdf_2 = base64.b64decode(res[1])

        return pdfs

    def __check_split_object_validity(self, split):
        """Checks whether the split object contains the essential information to be
        processed by the server."""

        if split is None:
            raise Pdf4meClientException('The split parameter cannot be None.')
        elif split.document is None or split.document.doc_data is None:
            raise Pdf4meClientException('The split document cannot be None.')
        elif split.split_action is None:
            raise Pdf4meClientException('The split_action cannot be None.')
        elif split.split_action.split_after_page is None or split.split_action.split_after_page == 0:
            raise Pdf4meClientException('The split_after_page of split_action cannot be None or zero.'
                                        'The first page of a PDF corresponds to page number one.')
