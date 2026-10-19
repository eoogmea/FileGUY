from pdf4me.model import CreateMailing, MailMergeDoc, MailMergeJob


class DocGenerationClient(object):

    def __init__(self, pdf4me_client):
        self.pdf4me_client = pdf4me_client

    def create_mailing(self, create_mailing):
        """create_mailing

        :param create_mailing:
        :type create_mailing: CreateMailing
        :return: CreateMailingRes
        """
        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=create_mailing,
                                                                   controller='DocGeneration/CreateMailing')

        return res

    def mail_merge(self, file, jobid_ext):
        """The provided Non-PDF file gets converted to a PDF.

        :param jobid_ext:
        :type jobid_ext:
        :param file: file to be converted
        :type file: file handler, use the method get_file_handler from FileReader to obtain it
        :return: bytes of resulting PDF file, can be directly written to file on disk
        """

        streams = [('file', file)]
        params = [('jobIdExt', jobid_ext)]

        return self.pdf4me_client.custom_http.post_wrapper(octet_streams=streams, values=params,
                                                           controller='DocGeneration/MailMerge')

    def mail_merge_doc(self, mail_merge_doc):
        """mail_merge_doc

        :param mail_merge_doc:
        :type mail_merge_doc: MailMergeDoc
        :return: MailMergeDocRes
        """

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=mail_merge_doc,
                                                                   controller='DocGeneration/MailMergeDoc')

        return res

    def mail_merge_job(self, mail_merge_job):
        """mail_merge_job

        :param mail_merge_job:
        :type mail_merge_job: MailMergeJob
        :return: MailMergeJobRes
        """

        res = self.pdf4me_client.custom_http.post_universal_object(universal_object=mail_merge_job,
                                                                   controller='DocGeneration/MailMergeJob')

        return res
