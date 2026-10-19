import requests

from pdf4me.helper.json_converter import JsonConverter
from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException, Pdf4meBackendException
from pdf4me.helper.response_checker import ResponseChecker


# from pdf4me.helper.token_generator import TokenGenerator

class CustomHttp(object):

    def __init__(self, token, apiurl):

        self.token = token
        self.json_converter = JsonConverter()
        self.url = "https://api.pdf4me.com/"
        if apiurl is not None and len(apiurl) != 0:
            self.url = apiurl
        self.userAgent = "pdf4me-python/0.8.23"

    def post_universal_object(self, universal_object, controller):
        """Sends a post request to the specified controller with the given
        universal_object as a body.

        :param universal_object: object to be sent
        :type universal_object: object
        :param controller: swagger controller
        :type controller: str
        :return: post response
        """

        # prepare post request
        request_url = self.url + controller
        headers = {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'Authorization': 'Basic ' + self.token,
            'User-Agent': self.userAgent
        }

        # convert body to json
        body = self.json_converter.dump(element=universal_object)

        # send request
        res = requests.post(request_url, data=body, headers=headers)

        # check status code
        self.__check_status_code(res)

        # check docLogs for error messages
        self.__check_docLogs_for_error_messages(res)

        # read content from response
        json_response = self.json_converter.load(res.text)

        return json_response

    def post_wrapper(self, octet_streams, values, controller):
        """Builds a post requests from the given parameters.

        :param octet_streams: (key: file identifier, value: open(fileName, 'rb'))) pairs
        :type octet_streams: list
        :param values: (key: identifier of value, value: content of value) pairs
        :type values: list
        :param controller: swagger controller
        :type controller: str
        :return: post response
        """

        # prepare post request
        request_url = self.url + controller
        header = {'Authorization': 'Basic ' + self.token, 'User-Agent': self.userAgent}

        # build files
        if octet_streams is not None and len(octet_streams) != 0:
            files = {key: value for (key, value) in octet_streams}
        else:
            files = None

        # build values
        if len(values) != 0:
            data = {key: value for (key, value) in values}
        else:
            data = None

        # send request
        if files is None:
            if data is None:
                raise Pdf4meClientException("Please provide at least one value or an octet-stream.")
            else:
                res = requests.post(request_url, data=data, headers=header)
        else:
            if data is None:
                res = requests.post(request_url, files=files, headers=header)
            else:
                res = requests.post(request_url, files=files, data=data, headers=header)

        # check status code
        self.__check_status_code(res)

        # check docLogs for error messages
        self.__check_docLogs_for_error_messages(res)

        return res.content

    def get_object(self, query_strings, controller):
        """Sends a get request to the specified controller with the given query strings.

        :param query_strings: params to be sent
        :type query_strings: str
        :param controller: swagger controller
        :type controller: str
        :return: post response
        """

        # prepare post request
        request_url = self.url + controller
        headers = {
            'Authorization': 'Basic ' + self.token,
            'User-Agent': self.userAgent
        }

        # send request
        res = requests.get(request_url, data=query_strings, headers=headers)

        # check status code
        self.__check_status_code(res)

        # check docLogs for error messages
        self.__check_docLogs_for_error_messages(res)

        # read content from response
        json_response = self.json_converter.load(res.text)
        
        return json_response

    def get_wrapper(self, query_strings, controller):
        """Sends a get request to the specified controller with the given
        query string and returns a file

        :param query_strings: params to be sent
        :type query_strings: str
        :param controller: swagger controller
        :type controller: str
        :return: file
        """

        # prepare post request
        request_url = self.url + controller
        headers = {
            'Authorization': 'Basic ' + self.token,
            'User-Agent': self.userAgent
        }

        # send request
        res = requests.get(request_url, data=query_strings, headers=headers)

        # check status code
        self.__check_status_code(res)

        # check docLogs for error messages
        self.__check_docLogs_for_error_messages(res)

        return res.content

    def __check_status_code(self, response):
        '''
        Checks whether the status code is either 200 or 204, otw. throws a Pdf4meBackendException.
        :param response: post response
        :type response: requests.Response
        :return: None
        '''

        status_code = response.status_code
        status_reason = response.reason

        if status_code == 500:
            server_error = self.json_converter.load(response.text)['error_message']
            trace_id = self.json_converter.load(response.text)['trace_id']
            raise Pdf4meBackendException('HTTP 500 ' + status_reason + " : trace_id " + trace_id + " : " + server_error)
        elif status_code != 200 and status_code != 204:
            error = response.text
            raise Pdf4meBackendException('HTTP ' + str(status_code) + ': ' + status_reason + " : " + error)

    def __check_docLogs_for_error_messages(self, response):
        '''
        Checks whether the HTTP response's docLogs contain any error message, in case of an error
         a Pdf4meBackendException is thrown.
        :param response: post response
        :type response: requests.Response
        :return: None
        '''

        ResponseChecker().check_response_for_errors(response.text)
