import base64


class FileReader(object):

    def get_file_data(self, path):
        """Returns the content of the file in base64.

        :param path: path to file
        :type path: str
        :return: base64
        """

        with open(path, 'rb') as file:
            pdf_base64 = file.read()

        data = base64.b64encode(pdf_base64)
        data = data.decode('utf-8')

        return data

    def get_file_handler(self, path):
        """Returns the file handler of the provided file.

        :param path: path to file
        :type path: str
        :return: the file handler of the provided file
        """
        return open(path, 'rb')
