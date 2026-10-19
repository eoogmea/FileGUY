from pdf4me.model import UserPreference


class UserClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def get_customer(self):  
        """get_customer  

        :return: CustomerBO
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='User/GetCustomer')

    def get_user_applications(self):  
        """get_user_applications  

        :return: GetUserApplicationRes
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='User/GetUserApplications')

    def get_user_applications_by_login_id(self, login_id):  
        """get_user_applications_by_login_id  

        :param str login_id:
        :return: GetUserApplicationRes
        """
        params = [('loginId', login_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params,
                                                         controller='User/GetUserApplicationsByLoginId')

    def get_user_preference(self):  
        """get_user_preference  

        :return: GetUserPrefRes
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='User/GetUserPreference')

    def get_user_token(self, login_id):  
        """get_user_token  

        :param str login_id:
        :return: GetUserTokenRes
        """
        params = [('loginId', login_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='User/GetUserToken')

    def save_user_preference(self, user_preference_req):  
        """save_user_preference  

        :param UserPreference user_preference_req:
        :return: SaveUserPrefRes
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object=user_preference_req,
                                                                    controller='User/SaveUserPreference')
