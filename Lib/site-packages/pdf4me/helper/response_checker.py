from pdf4me.helper.json_converter import JsonConverter
from pdf4me.helper.pdf4me_exceptions import Pdf4meBackendException

# python version dependent import
try:
    from json import JSONDecoder
except ImportError:
    pass


class ResponseChecker(object):

    def check_response_for_errors(selfself, res):
        '''
        Checks document logs for error reports, in case an error is found a Pdf4meBackendException is thrown.
        :param res: content of response
        :type res: string
        :return: None
        '''

        jsonConverter = JsonConverter()

        if res is None:
            raise Pdf4meBackendException('Server Error')

        # check whether the res contains none, one or multiple documents
        one_doc = None
        multiple_docs = None
        docs = []

        try:
            one_doc = jsonConverter.load(res)['document']
        except KeyError:
            pass
        except ValueError:
            pass
        except TypeError:
            pass
        except JSONDecoder:
            pass

        try:
            multiple_docs = jsonConverter.load(res)['documents']
        except KeyError:
            pass
        except ValueError:
            pass
        except TypeError:
            pass
        except JSONDecoder:
            pass

        if one_doc is not None:
            docs = [one_doc]
        elif multiple_docs is not None:
            docs = multiple_docs
        # otw docs is empty

        # extract the logs
        doc_logs = []
        for i in range(len(docs)):
            doc_logs = doc_logs + docs[i]['doc_logs']

        # check all logs for errors
        for i in range(len(doc_logs)):
            current_doc_log = doc_logs[i]
            if (current_doc_log['doc_log_level'] == 3):
                raise Pdf4meBackendException(current_doc_log['message'])
