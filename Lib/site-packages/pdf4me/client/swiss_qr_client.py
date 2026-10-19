from pdf4me.model import SwissQrCreatorReq, ReadBarcodes, SwissQrCreatorListReq


class SwissQRApi(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def create_swiss_qr_bill(self, swiss_qr_creator_req):  
        """create_swiss_qr_bill  

        :param SwissQrCreatorReq swiss_qr_creator_req:
        :return: SwissQrCreatorRes
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object=swiss_qr_creator_req,
                                                                    controller='SwissQR/CreateSwissQrBill')

    def create_swiss_qr_bill_list(self, swiss_qr_creator_list_req):  
        """create_swiss_qr_bill_list  

        :param SwissQrCreatorListReq swiss_qr_creator_list_req:
        :return: SwissQrCreatorListRes
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object=swiss_qr_creator_list_req,
                                                                    controller='SwissQR/CreateSwissQrBillList')

    def read_swiss_qr_code(self, read_barcodes):  
        """read_swiss_qr_code  

        :param ReadBarcodes read_barcodes:
        :return: ReadBarcodesRes
                 If the method is called asynchronously,
                 returns the request thread.
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object=read_barcodes,
                                                                    controller='SwissQR/ReadSwissQrCode')
