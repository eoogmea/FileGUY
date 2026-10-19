from pdf4me.model import JobFlow, JobFlowPlan, RunJobFlow, JobConfig, Machine


class JobClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    # def create_archive_job_config(self, archive_job_req):  
    #     """CreateArchiveJobConfig  
    #
    #     Create Archive Job which continuously converts your PDF to Pdf-A.  
    #
    #     :param ArchiveJobReq archive_job_req:
    #     :return: ArchiveJobRes
    #     """
    #     res = self.pdf4me_client.custom_http.post_universal_object(universal_object=archive_job_req,
    #                                                                controller='Job/CreateArchiveJobConfig')
    #
    #     return res

    def delete_job_flow(self, job_flow_req): 
        """delete_job_flow  

        :param JobFlow job_flow_req:
        :return: DeleteJobFlowRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=job_flow_req,
                                                                   controller='Job/DeleteJobFlow')

        return res

    def delete_job_flow_plan(self, job_flow_plan_req):  
        """delete_job_flow_plan  

        :param JobFlowPlan job_flow_plan_req:
        :return: DeleteJobFlowPlanRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=job_flow_plan_req,
                                                                   controller='Job/DeleteJobFlowPlan')

        return res

    def get_all_job_flow_plans(self):  
        """get_all_job_flow_plans  

        :return: GetJobFlowPlanRes
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Job/GetAllJobFlowPlans')

    def get_all_job_flows(self):  
        """get_all_job_flows  

        :return: GetJobFlowRes
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Job/GetAllJobFlows')

    def get_job_flows(self, job_flow_id):  
        """get_job_flows  

        :param str job_flow_id:
        :return: list[JobFlow]
        """
        params = [('jobFlowId', job_flow_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Job/GetJobFlows')

    def get_job_flows_by_machine_id(self, machine_id):  
        """get_job_flows_by_machine_id  

        :param str machine_id:
        :return: list[JobFlow]
        """
        params = [('machineId', machine_id)]

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Job/GetJobFlows')

    def get_machines(self):  
        """get_machines  

        :return: GetMachineRes
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Job/GetMachines')

    def job_configs(self):  
        """JobConfigs  

        :return: list[JobConfig]
        """
        params = []

        return self.pdf4me_client.custom_http.get_object(query_strings=params, controller='Job/JobConfigs')

    def run_job_flow(self, run_job_flow):  
        """run_job_flow  

        :param RunJobFlow run_job_flow:
        :return: RunJobRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=run_job_flow,
                                                                   controller='Job/RunJobFlow')

        return res

    def save_job_config(self, job_config):  
        """SaveJobConfig  

        Create a Job Configuration for recurring executions or pipelining of pdf4me functionality.  With pipelining it is possible to run pdf4me base functionality in a specific order as a single execution.  

        :param JobConfig job_config:
        :return: JobConfigRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=job_config,
                                                                   controller='Job/SaveJobConfig')

        return res

    def save_job_flow(self, job_flow_req):  
        """save_job_flow  

        :param JobFlow job_flow_req:
        :return: SaveJobFlowRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=job_flow_req,
                                                                   controller='Job/SaveJobFlow')

        return res

    def save_job_flow_plan(self, job_flow_plan_req):  
        """save_job_flow_plan  

        :param JobFlowPlan job_flow_plan_req:
        :return: SaveJobFlowPlanRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=job_flow_plan_req,
                                                                   controller='Job/SaveJobFlowPlan')

        return res

    def save_job_flows(self, job_flow):  
        """save_job_flows  

        :param JobFlow job_flow:
        :return: JobFlowPlan
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=job_flow,
                                                                   controller='Job/SaveJobFlows')

        return res

    def save_machine(self, machine_req):  
        """save_machine  

        :param Machine machine_req:
        :return: SaveMachineRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=machine_req,
                                                                   controller='Job/SaveMachine')

        return res
