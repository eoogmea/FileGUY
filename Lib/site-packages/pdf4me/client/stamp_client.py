from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException
from pdf4me.model import Stamp


class StampClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def stamp(self, stamp):
        """The predefined stamp is carried out.

        :param stamp: stamp configuration
        :type stamp: Stamp
        :return: StampRes, contains the stamped PDF
        """

        # check stamp validity
        self.__check_stamp_object_validity(stamp)

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=stamp, controller='Stamp/Stamp')

        return res

    def text_stamp(self, text, pages, align_x, align_y, file):
        """Places a custom text stamp on the pages of your choice. The position of the stamp is specified by alignX and alignY.

        :param text: custom text on stamp
        :type text: str
        :param pages: pages to be stamped
        :type pages: str: {'all', 'first', 'last', 'odd', 'even', 'notFirst', 'otLast'} or a list of comma separated page numbers e.g. '1,4', if the first and fourth page need to be stamped
        :param align_x: horizontal position
        :type align_x: str: {'left', 'center', 'right'}
        :param align_y: vertical position
        :type align_y: str: {'top', 'middle', 'bottom'}
        :param file: to be stamped
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('text', text), ('pages', pages), ('alignX', align_x), ('alignY', align_y)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='Stamp/TextStamp')

    def __check_stamp_object_validity(self, stamp):
        """checks whether the stamp object contains the essential information to be
        processed by the server."""

        if stamp is None:
            raise Pdf4meClientException('The stamp parameter cannot be None.')
        elif stamp.document is None or stamp.document.doc_data is None:
            raise Pdf4meClientException('The stamp document cannot be None.')
        elif stamp.stamp_action is None:
            raise Pdf4meClientException('The stamp_action cannot be None.')
        elif stamp.stamp_action.alpha is None:
            raise Pdf4meClientException('The alpha parameter of stamp_action cannot be None.')
        elif stamp.stamp_action.text is None and stamp.stamp_action.image is None:
            raise Pdf4meClientException('The image and text parameter of stampAction cannot both be None.')
