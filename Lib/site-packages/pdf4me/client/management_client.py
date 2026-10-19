from pdf4me.model import UsageSubscriptionReq, UsageApplicationReq


class ManagementClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def check_availability(self):  
        """check_availability  

        :return: CheckAvailabilityRes
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params,
                                                         controller='Management/CheckAvailability')

    def get_api_call_detail(self, api_call_id):  
        """get_api_call_detail  

        :param str api_call_id:
        :return: ApiCallDetail
        """
        params = [('apiCallId', api_call_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Management/GetApiCallDetail')

    def get_latest_package(self):  
        """get_latest_package  

        :return: Package
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Management/GetLatestPackage')

    def get_subscription_usage(self):  
        """get_subscription_usage  

        :return: SubscriptionUsage
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object={},
                                                                    controller='Management/GetSubscriptionUsage')

    def get_usage_application(self, usage_application_req):  
        """get_usage_application  

        :param UsageApplicationReq usage_application_req:
        :return: UsageApplicationRes
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object=usage_application_req,
                                                                    controller='Management/GetUsageApplication')

    def get_usage_subscription(self, usage_subscription_req):  
        """get_usage_subscription  

        :param UsageSubscriptionReq usage_subscription_req:
        :return: UsageSubscriptionRes
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object=usage_subscription_req,
                                                                    controller='Management/GetUsageSubscription')

    def version(self):  
        """Version  

        A simple method to get the current Version  
        :return: VersionRes
        """
        return self.pdf4me_client.custom_http.post_universal_object(universal_object={}, controller='Management/Version')
