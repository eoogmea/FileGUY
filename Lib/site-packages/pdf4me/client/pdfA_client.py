from pdf4me.helper.json_converter import JsonConverter
from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException
from pdf4me.model import CreatePdfA, Rotate, Protect, Validate, Repair, SignPdf, DocMetadata


class PdfAClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def pdfA(self, create_pdfA):
        """The predefined PDF/A creation is carried out.

        :param create_pdfA: PDF/A configuration
        :type create_pdfA: CreatePdfA
        :return: CreatePdfARes, contains generated PDF/A.
        """

        # check create_pdfA validity
        self.__check_pdfA_object_validity(create_pdfA)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=create_pdfA,
                                                                   controller='PdfA/PdfA')

        return res

    def create_pdfA(self, file, pdf_compliance=None):
        """Converts PDF documents to the PDF/A format for long-term archiving.

        :param pdf_compliance: PDF/A compliance level
        :type pdf_compliance: str
        :param file: to be converted to PDF/A
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]

        if (pdf_compliance is None):
            params = []
        else:
            params = [('pdfCompliance', pdf_compliance)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='PdfA/CreatePdfA')

    def __check_pdfA_object_validity(self, create_pdfA):
        """checks whether the create_pdfA object contains the essential information to be
        processed by the server."""

        if create_pdfA is None:
            raise Pdf4meClientException('The create_pdfA parameter cannot be None.')
        elif create_pdfA.document is None or create_pdfA.document.doc_data is None:
            raise Pdf4meClientException('The create_pdfA document cannot be None.')
        elif create_pdfA.pdf_a_action is None:
            raise Pdf4meClientException('The pdf_a_action cannot be None.')

    # Rotate
    def rotate(self, rotate):
        """The predefined rotation is carried out.

        :param rotate: Rotate configuration
        :type rotate: Rotate
        :return: RotateRes, contains rotated pages.
        """

        # check rotate validity
        self.__check_rotate_object_validity(rotate)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=rotate,
                                                                   controller='PdfA/Rotate')

        return res

    def rotate_page(self, file, page_nr, rotate):
        """Rotates selected pages of the document in given direction.

        :param page_nr: determines after which page the rotate takes place
        :type page_nr: str
        :param rotate: determines rotate direction
        :type rotate: str
        :param file: file to be rotated
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('pageNr', page_nr), ('rotate', rotate)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='PdfA/RotatePage')

    def rotate_document(self, file, rotate):
        """Rotates all pages of the document in given direction.

        :param rotate: determines rotate direction
        :type rotate: str
        :param file: file to be rotated
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('rotate', rotate)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='PdfA/RotateDocument')

    def __check_rotate_object_validity(self, rotate):
        """checks whether the rotate object contains the essential information to be
        processed by the server."""

        if rotate is None:
            raise Pdf4meClientException('The rotate parameter cannot be None.')
        elif rotate.document is None or rotate.document.doc_data is None:
            raise Pdf4meClientException('The rotate document cannot be None.')
        elif rotate.rotate_action is None:
            raise Pdf4meClientException('The rotate_action cannot be None.')

    # Protect
    def protect(self, protect):
        """The predefined rotation is carried out.

        :param protect: protect configuration
        :type protect: Protect
        :return: ProtectRes, contains protected document.
        """

        # check protect validity
        self.__check_protect_object_validity(protect)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=protect,
                                                                   controller='PdfA/Protect')

        return res

    def protect_document(self, file, password, permissions):
        """protects selected pages of the document in given direction.

        :param password: password to be set for document
        :type password: str
        :param permissions: permissions to be set for document
        :type permissions: str
        :param file: to be protected
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('password', password), ('permissions', permissions)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='PdfA/ProtectDocument')

    def unlock_document(self, file, password):
        """protects all pages of the document in given direction.

        :param password: password of the document
        :type password: str
        :param file: to be unlocked
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('password', password)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='PdfA/UnlockDocument')

    def __check_protect_object_validity(self, protect):
        """checks whether the protect object contains the essential information to be
        processed by the server."""

        if protect is None:
            raise Pdf4meClientException('The protect parameter cannot be None.')
        elif protect.document is None or protect.document.doc_data is None:
            raise Pdf4meClientException('The protect document cannot be None.')
        elif protect.protect_action is None:
            raise Pdf4meClientException('The protect_action cannot be None.')

    # validate
    def validate(self, validate):
        """The predefined rotation is carried out.

        :param validate: validate configuration
        :type validate: Validate
        :return: validateRes, contains validation info.
        """

        # check validate validity
        self.__check_validate_object_validity(validate)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=validate,
                                                                   controller="PdfA/Validate")

        return res

    def validate_document(self, file, pdf_compliance):
        """validates selected pages of the document in given direction.

        :param pdf_compliance: PDF/A compliance to check
        :type pdf_compliance: str
        :param file: to be validated
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('pdfCompliance', pdf_compliance)]

        res = self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='PdfA/ValidateDocument')
        # get json
        res = JsonConverter().load(res)

        return res

    def __check_validate_object_validity(self, validate):
        """checks whether the validate object contains the essential information to be
        processed by the server."""

        if validate is None:
            raise Pdf4meClientException('The validate parameter cannot be None.')
        elif validate.document is None or validate.document.doc_data is None:
            raise Pdf4meClientException('The validate document cannot be None.')
        elif validate.validate_action is None:
            raise Pdf4meClientException('The validate_action cannot be None.')

    # Repair
    def repair(self, repair):
        """The predefined rotation is carried out.

        :param repair: repair configuration
        :type repair: Repair
        :return: repairRes, contains repaired document with validation info.
        """

        # check repair validity
        self.__check_repair_object_validity(repair)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=repair,
                                                                   controller="PdfA/Repair")

        return res

    def repair_document(self, file):
        """repairs selected pages of the document in given direction.

        :param file: to be repaired
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = []

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='PdfA/RepairDocument')

    def __check_repair_object_validity(self, repair):
        """checks whether the repair object contains the essential information to be
        processed by the server."""

        if repair is None:
            raise Pdf4meClientException('The repair parameter cannot be None.')
        elif repair.document is None or repair.document.doc_data is None:
            raise Pdf4meClientException('The repair document cannot be None.')
        # elif repair.repair_action is None:
        #    raise Pdf4meClientException('The repair_action cannot be None.')

    def sign_pdf(self, sign_pdf):
        """The predefined rotation is carried out.

        :param sign_pdf: SignPdf configuration
        :type sign_pdf: SignPdf
        :return: SignPdfRes, contains signed document.
        """

        # check sign_pdf validity
        self.__check_sign_pdf_object_validity(sign_pdf)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=sign_pdf,
                                                                   controller="PdfA/SignPdf")

        return res

    def __check_sign_pdf_object_validity(self, sign_pdf):
        """checks whether the sign_pdf object contains the essential information to be
        processed by the server."""

        if sign_pdf is None:
            raise Pdf4meClientException('The sign_pdf parameter cannot be None.')
        elif sign_pdf.document is None or sign_pdf.document.doc_data is None:
            raise Pdf4meClientException('The sign_pdf document cannot be None.')
        elif sign_pdf.sign_action is None:
            raise Pdf4meClientException('The sign_action cannot be None.')

    def metadata(self, file):
        """repairs selected pages of the document in given direction.

        :param file: to be repaired
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: DocMetadata, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = []

        res = self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                          controller='PdfA/Metadata')

        # get json
        return JsonConverter().load(res)

