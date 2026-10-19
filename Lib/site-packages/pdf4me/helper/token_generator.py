# import base64
# import json
# import time

# import adal
# from msrestazure.azure_active_directory import AADTokenCredentials

# TENANT = "devynooxlive.onmicrosoft.com"
# AUTHORITY_URI = "https://login.microsoftonline.com/" + TENANT

# # slack: time before expiration, when already a new token is acquired
# SLACK = 60  # 1min


# class TokenGenerator(object):

#     def __init__(self, client_id, secret):

#         self.client_id = client_id
#         self.resource_uri = self.client_id
#         self.secret = secret
#         self.token = None

#     def get_token(self):
#         '''
#         Loads or aquires the authorization token.
#         :return: token string
#         '''
#         if self.__valid_token() is True:
#             return self.token
#         else:
#             return self.__get_new_token()

#     def __valid_token(self):
#         '''
#         Checks whether the stored token has not expired yet.
#         :return: whether the stored token is still valid
#         '''

#         if self.token is None:
#             return False

#         timestamp = self.token.split('.')[1]
#         # correct padding
#         timestamp = timestamp + '=' * (-len(timestamp) % 4)
#         # json format
#         json_bytes = base64.b64decode(timestamp)
#         # extract expiration time [sec]
#         json_string = json_bytes.decode("utf-8")
#         json_object = json.loads(json_string)
#         expiration = json_object["exp"]

#         # current time
#         now = time.time()

#         # check
#         if now < expiration - SLACK:
#             return True
#         else:
#             return False

#     def __get_new_token(self):
#         '''
#         Acquires a new token using the clientID and secret.
#         :return: newly generated token
#         '''

#         # generate token
#         context = adal.AuthenticationContext(AUTHORITY_URI, api_version=None)
#         mgmt_token = context.acquire_token_with_client_credentials(self.resource_uri, self.client_id, self.secret)
#         credentials = AADTokenCredentials(mgmt_token, self.client_id)
#         token = credentials.token["access_token"]

#         # store new token
#         self.token = token

#         return self.token
