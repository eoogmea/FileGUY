import jprops

from pdf4me.helper.custom_http import CustomHttp
from pdf4me.helper.pdf4me_exceptions import Pdf4meClientException

DEFAULT_CONFIG_PROPERTIES = '../config.properties'


class Pdf4meClient(object):

    def __init__(self, token=None, apiurl=None, path_to_config_file=DEFAULT_CONFIG_PROPERTIES):

        # in case either token is not provided, load them from the property file
        if token is None:

            # load token from property file
            try:
                with open(path_to_config_file) as fp:
                    props = list(jprops.iter_properties(fp))

                    prop_token = None
                    prop_apiurl = None

                    # check whether token and apiurl exist
                    for key, value in props:
                        if key == 'token':
                            prop_token = value
                        if key == 'apiurl':
                            prop_apiurl = value

                    if prop_token is None:
                        raise Pdf4meClientException(
                            'Please store the token in the  ' + path_to_config_file + 'file or provide your token in the Pdf4meClient constructor.')
            except IOError:
                raise Pdf4meClientException(
                    'The config.properties file is not stored at \'' + path_to_config_file + '\'. Please initialize the Pdf4meClient object with the correct path to your conifg.properties file. Or else provide your token in the Pdf4meClient constructor.')

            self.custom_http = CustomHttp(prop_token, prop_apiurl)
        else:
            self.custom_http = CustomHttp(token, apiurl)
