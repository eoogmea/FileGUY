from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException
from pdf4me.model import Merge


class MergeClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def merge(self, merge):
        """The predefined merge is carried out.

        :param merge: merge configuration
        :type merge: Merge
        :return: MergeRes, contains merged PDF
        """

        # check merge validity
        self.__check_merge_object_validity(merge)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=merge,
                                                                   controller='Merge/Merge')

        return res

    def merge_2_pdfs(self, file1, file2):
        """Merges the two provided PDF files.

        :param file1: first PDF
        :type file1: file handler, use the method get_file_handler from FileReader to obtain it
        :param file2: second PDF
        :type file2: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file1', file1), ('file2', file2)]
        params = []

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='Merge/Merge2Pdfs')

    def __check_merge_object_validity(self, merge):
        """Checks whether the merge object contains the essential information to be
        processed by the server."""

        if merge is None:
            raise Pdf4meClientException('The merge parameter cannot be None.')
        elif merge.documents is None:
            raise Pdf4meClientException('The merge documents cannot be None.')
        elif merge.merge_action is None:
            raise Pdf4meClientException('The merge_action cannot be None.')

        numDocs = len(merge.documents)
        # check whether there are at least two documents
        if numDocs < 2:
            raise Pdf4meClientException('The merge documents must contain at least two documents.')
        # check whether all documents are not undefined neither is their docData
        docs = merge.documents
        for i in range(numDocs):
            currentDoc = docs[i]
            if currentDoc is None or currentDoc.doc_data is None:
                raise Pdf4meClientException('The merge documents cannot be None nor can the document.docData.')
