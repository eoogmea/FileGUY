import base64

from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException
from pdf4me.model import CreateImages
from pdf4me.helper.json_converter import JsonConverter


class ImageClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def create_images(self, create_images):
        """The predefined image creation is carried out.

        :param create_images: image creation configuration
        :type create_images: CreateImages
        :return: CreateImagesRes, contains images
        """

        # check create_images validity
        self.__check_create_images_object_validity(create_images)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=create_images,
                                                                   controller='Image/CreateImages')

        return res

    def create_thumbnail(self, width, page_nr, image_format, file):
        """Produces a thumbnail of the page referenced by the pageNr. Be aware of the one-indexing of the page numbers.

        :param width: size of the produced thumbnail
        :type width: int
        :param page_nr: number of the page which should be captured by the thumbnail
        :type page_nr: str
        :param image_format: picture format of thumbnail, e.g. 'jpg'
        :type image_format: str
        :param file: file to capture thumbnails from
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of produced thumbnail, can be directly written to image file on disk
        """

        streams = [('file', file)]
        params = [('width', width), ('pageNr', page_nr), ('imageFormat', image_format)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='Image/CreateThumbnail')

    def create_thumbnails(self, width, page_nrs, image_format, file):
        """Produces a thumbnail of the page referenced by the pageNr. Be aware of the one-indexing of the page numbers.

        :param width: size of the produced thumbnail
        :type width: int
        :param page_nrs: number of the page which should be captured by the thumbnail
        :type page_nrs: str
        :param image_format: picture format of thumbnail, e.g. 'jpg'
        :type image_format: str
        :param file: file to capture thumbnails from
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of produced thumbnail, can be directly written to image file on disk
        """

        streams = [('file', file)]
        params = [('width', width), ('pageNrs', page_nrs), ('imageFormat', image_format)]

        res = self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                          controller='Image/CreateThumbnails')

        # get json
        res = JsonConverter().load(res)

        # extract the two documents
        pdfs = []
        for pdf in res:
            pdfs.append(base64.b64decode(pdf))
        # pdf_1 = base64.b64decode(res[0])
        # pdf_2 = base64.b64decode(res[1])

        return pdfs

    def __check_create_images_object_validity(self, create_images):
        """Checks whether the create_images object contains the essential information to be
                processed by the server."""

        if create_images is None:
            raise Pdf4meClientException('The create_images parameter cannot be None.')
        elif create_images.document is None or create_images.document.doc_data is None:
            raise Pdf4meClientException('The create_images document cannot be None.')
        elif create_images.image_action is None:
            raise Pdf4meClientException('The image_action cannot be None.')
        elif create_images.image_action.page_selection is None:
            raise Pdf4meClientException('The page_selection of the image_action cannot be None.')
