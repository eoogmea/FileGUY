from pdf4me.model import DropDocumentReq, ProduceDocuments


class DocumentApi(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def drop_document(self, drop_document_req):
        """drop_document

        :param drop_document_req:
        :type drop_document_req: DropDocumentReq
        :return: DropDocumentRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=drop_document_req,
                                                                   controller='Document/DropDocument')

        return res

    def get_document(self, job_id, document_id):
        """get_document

        :param job_id:
        :type job_id:str
        :param document_id:
        :type document_id:str
        :return: file
        """
        params = [('jobId', job_id), ('documentId', document_id)]

        return self.pdf4me_client.custom_http.get_wrapper(query_strings=params, controller='Document/GetDocument')

    def get_documents(self, job_id):
        """get_document_url

        :param job_id:
        :type job_id:str
        :return: GetDocumentListRes
        """
        params = [('jobId', job_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Document/GetDocuments')

    def get_document_url(self,  job_id, document_id):
        """get_document_url

        :param job_id:
        :type job_id:str
        :param document_id:
        :type document_id:str
        :return: GetDocumentUrlRes
        """
        params = [('jobId', job_id), ('documentId', document_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Document/GetDocumentUrl')

    def get_document_url_with_ext(self, job_id, document_id, file_name):
        """get_document_url_with_ext

        :param str job_id:
        :param str document_id:
        :param str file_name:
        :return: GetDocumentUrlRes
        """
        params = [('jobId', job_id), ('documentId', document_id), ('fileName', file_name)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params,
                                                         controller='Document/GetDocumentUrlWithExt')

    def get_image_url(self,  job_id, image_id):
        """get_document_url

        :param job_id:
        :type job_id:str
        :param image_id:
        :type image_id:str
        :return: GetDocumentUrlRes
        """
        params = [('jobId', job_id), ('imageId', image_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Document/GetImageUrl')

    def get_image_url_with_ext(self, job_id, image_id, file_name):
        """get_document_url_with_ext

        :param str job_id:
        :param str image_id:
        :param str file_name:
        :return: GetDocumentUrlRes
        """
        params = [('jobId', job_id), ('imageId', image_id), ('fileName', file_name)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params,
                                                         controller='Document/GetImageUrlWithExt')

    def produce_documents(self, produce_documents):
        """produce_documents

        :param ProduceDocuments produce_documents:
        :return: ProduceDocumentsRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=produce_documents,
                                                                   controller='Document/ProduceDocuments')

        return res
