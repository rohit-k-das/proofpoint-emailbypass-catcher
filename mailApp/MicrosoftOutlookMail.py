import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import datetime
import urllib
import json
import logging
import time
import concurrent.futures
import ConfigParser
import os

logger = logging.getLogger(__name__)

MAX_THREADS = 14  # Get max number of threads for multi-threading

graph_api = "https://graph.microsoft.com/v1.0/{0}"

Config = ConfigParser.ConfigParser()
Config.read(os.path.join(os.path.abspath(os.path.dirname(__file__)),'Mail_creds'))
graph_application_id = Config.get('Settings', 'Microsoft_Application_ID')
graph_secret = Config.get('Settings', 'Microsoft_Application_Secret')
graph_tenant_id = Config.get('Settings', 'Microsoft_Tenant_ID')
company_domain = Config.get('Settings', 'Company_Domain')

emails = []  # All mails that match the search criteria
filtered_emails = []  # All mails from emails that are not in the deleted folder
filtered_deleted_emails = []  # All mails from emails that are in the deleted folder
active_employee_usernames = []  # A list of username of all active employees


# Generate session to be used in get/post request with max of 3 retries and interval of 1 second
def session_generator():
    session = requests.Session()
    retry = Retry(connect=3, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session


# Make OAuth Call to get Access token
def generate_access_token():
    logger.info('Generating access token to access the Outlook Monitoring app in Azure')
    access_token = ""
    expiry_time = datetime.datetime.now()
    token_url = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token".format(graph_tenant_id)
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    # Current default scope defines all permissions associated with Mail Monitoring App
    data = "client_id=%s&scope=https://graph.microsoft.com/.default&client_secret=%s&grant_type=client_credentials" % (graph_application_id, graph_secret)
    session = session_generator()
    resp = session.post(token_url, headers=headers, data=data)
    if resp.ok:
        response = resp.json()
        access_token = response['access_token']
        expiry_time = datetime.datetime.now() + datetime.timedelta(seconds=response['expires_in'])
        logger.info('Successfully generated access token for Outlook Monitoring app')
    # Handle Rate Limiting
    elif resp.status_code == 429:
        seconds_to_sleep = resp.headers['Retry-After']
        logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
        time.sleep(seconds_to_sleep)
        access_token, expiry_time = generate_access_token()
    # Handle broken api call
    elif resp.status_code == 503 or resp.status_code == 504:
        seconds_to_sleep = 1
        logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
        time.sleep(seconds_to_sleep)
        access_token, expiry_time = generate_access_token()
    else:
        logger.error("Unable to create access token")
        logger.error("%d:%s" % (resp.status_code, resp.text))
    return access_token, expiry_time


def check_and_renew_access_token(access_token):
    global oauth_access_token
    global expiry_time
    global oauth_refresh_token
    if access_token == oauth_access_token:
        logger.warning("Current access token is about to expire. Generating new access token")
        oauth_access_token, expiry_time = generate_access_token()
        if not oauth_access_token and expiry_time > datetime.datetime.now():
            logger.critical('Unable to generate new access token. Exiting..')
            exit(-1)


# Class to handle mails
class OutlookEmail:
    def __init__(self):
        self.sender = None
        self.recipient = None  # Actual recipient from the header
        self.envelope_recipient = None
        self.id = None
        self.received_date = None
        self.subject = None
        self.email_read = False
        self.in_deleteditems = False
        self.body = None
        self.header = None
        self.ccrecipients = None
        self.bccrecipients = None
        self.message_id = None
        self.has_attachments = False  # Doesn't include inline attachment
        self.requested_recipient = None  # As mentioned by the user

    # Moves the mail to the required folder based on the folder id
    def move_email_to_folder(self, mail_folder_id):
        status = False
        access_token = oauth_access_token
        headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
        user_url = graph_api.format('users/{0}/messages/{1}/move')
        data = {"destinationId": mail_folder_id}
        session = session_generator()
        query_start_time = datetime.datetime.now()

        # Make the API call if token expiry time is greater than 1 minute
        if int((expiry_time - query_start_time).seconds) > 60:
            resp = session.post(user_url.format(self.requested_recipient, self.id), headers=headers, json=data)
            if resp.status_code == 201:
                logger.info('Mail moved from user {0}'.format(self.requested_recipient))
                status = True
            elif resp.status_code == 429:
                seconds_to_sleep = resp.headers['Retry-After']
                logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
                time.sleep(seconds_to_sleep)
                status = self.move_email_to_folder(mail_folder_id)
            # Handle broken api call
            elif resp.status_code == 503 or resp.status_code == 504:
                seconds_to_sleep = 1
                logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
                time.sleep(seconds_to_sleep)
                status = self.move_email_to_folder(mail_folder_id)
            else:
                logger.error('Unable to move to the specified folder for user {}'.format(self.requested_recipient))
                logger.error("%d:%s" % (resp.status_code, resp.text))
        else:
            check_and_renew_access_token(access_token)
            status = self.move_email_to_folder(mail_folder_id)
        return status

    # Copy mail to a different folder in the same mailbox
    def copy_email(self, mail_folder_id):
        access_token = oauth_access_token
        headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
        user_url = graph_api.format('users/{0}/messages/{1}/copy')
        data = {"destinationId": mail_folder_id}
        session = session_generator()
        query_start_time = datetime.datetime.now()

        # Make the API call if token expiry time is greater than 1 minute
        if int((expiry_time - query_start_time).seconds) > 60:
            resp = session.post(user_url.format(self.requested_recipient, self.id), headers=headers, json=data)
            if resp.status_code == 201:
                logger.info('Mail copied from user {0}'.format(self.requested_recipient))
            # Handle Rate Limiting
            elif resp.status_code == 429:
                seconds_to_sleep = resp.headers['Retry-After']
                logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
                time.sleep(seconds_to_sleep)
                self.copy_email(mail_folder_id)
            # Handle broken api call
            elif resp.status_code or resp.status_code == 504:
                seconds_to_sleep = 1
                logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
                time.sleep(seconds_to_sleep)
                self.copy_email(mail_folder_id)
            # Handle other http error
            else:
                logger.error('Unable to copy to the specified folder for user {}'.format(self.requested_recipient))
                logger.error("%d:%s" % (resp.status_code, resp.text))
        # Create new access token
        else:
            check_and_renew_access_token(access_token)
            self.copy_email(mail_folder_id)

    # Permanently delete the mail
    def delete_mail(self):
        access_token = oauth_access_token
        headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
        user_url = graph_api.format('users/{0}/messages/{1}')
        session = session_generator()
        query_start_time = datetime.datetime.now()

        # Make the API call if token expiry time is greater than 1 minute
        if int((expiry_time - query_start_time).seconds) > 60:
            resp = session.delete(user_url.format(self.requested_recipient, self.id), headers=headers)
            if resp.status_code == 204:
                logger.info('Mail with sub:{1} deleted from user {0}'.format(self.recipient, self.subject))
            # Handle Rate Limiting
            elif resp.status_code == 429:
                seconds_to_sleep = resp.headers['Retry-After']
                logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
                time.sleep(seconds_to_sleep)
                self.delete_mail()
            # Handle broken api call
            elif resp.status_code == 503 or resp.status_code == 504:
                seconds_to_sleep = 1
                logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
                time.sleep(seconds_to_sleep)
                self.delete_mail()
            # Handle other http errors
            else:
                logger.error("%d:%s" % (resp.status_code, resp.text))
                logger.error('Unable to delete email with sub:{1} from user {0}'.format(self.requested_recipient, self.subject))
        # Create new access token
        else:
            check_and_renew_access_token(access_token)
            self.delete_mail()


# Fetches all mail folder ids present in a recipients mailbox
def get_mail_folders(recipient):
    mail_folders = []
    access_token = oauth_access_token
    headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
    url = graph_api.format('users/{0}/mailFolders?$top=999')
    session = session_generator()
    query_start_time = datetime.datetime.now()

    # Make the API call if token expiry time is greater than 1 minute
    if int((expiry_time - query_start_time).seconds) > 60:
        resp = session.get(url.format(recipient), headers=headers)
        if resp.ok:
            response = resp.json()
            if response['value']:
                for mail_folder in response['value']:
                    # Don't fetch folders Trash & Outbox
                    if mail_folder['displayName'] != "Deleted Items" and mail_folder['displayName'] != 'Outbox':
                        mail_folders.append(mail_folder['id'])
        # Handle Rate Limiting
        elif resp.status_code == 429:
            seconds_to_sleep = resp.headers['Retry-After']
            logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
            time.sleep(seconds_to_sleep)
            mail_folders = get_mail_folders(recipient)
        # Handle broken api call
        elif resp.status_code == 503 or resp.status_code == 504:
            seconds_to_sleep = 1
            logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
            time.sleep(seconds_to_sleep)
            mail_folders = get_mail_folders(recipient)
        else:
            logger.error("For recipient %s ERROR %d:%s" % (recipient, resp.status_code, resp.text))
    else:
        check_and_renew_access_token(access_token)
        mail_folders = get_mail_folders(recipient)
    return mail_folders


# Get folder name associated with a folder id in a recipient's mailbox
def get_folder_name(recipient, folder_id):
    folder_name = ""  # Default folder name
    access_token = oauth_access_token
    headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
    url = graph_api.format('users/{0}/mailFolders/{1}')
    session = session_generator()
    query_start_time = datetime.datetime.now()

    # Make the API call if token expiry time is greater than 1 minute
    if int((expiry_time - query_start_time).seconds) > 60:
        resp = session.get(url.format(recipient, folder_id), headers=headers)
        if resp.ok:
            response = resp.json()
            folder_name = response['displayName']
        # Handle Rate Limiting
        elif resp.status_code == 429:
            seconds_to_sleep = resp.headers['Retry-After']
            logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
            time.sleep(seconds_to_sleep)
            folder_name = get_folder_name(recipient, folder_id)
        # Handle broken api call
        elif resp.status_code == 503 or resp.status_code == 504:
            seconds_to_sleep = 1
            logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
            time.sleep(seconds_to_sleep)
            folder_name = get_folder_name(recipient, folder_id)
        # Handle other http errors
        else:
            logger.error("Unable to fetch folder name associated with recipient %s for folder id %s" % (recipient, str(folder_id)))
            logger.error("For recipient %s ERROR %d:%s" % (recipient, resp.status_code, resp.text))
    # Create new access token
    else:
        check_and_renew_access_token(access_token)
        folder_name = get_folder_name(recipient, folder_id)

    return folder_name


# Create filters to be used to search for very specific mails that match the criteria in the get_email function
def build_get_mail_filter(sender, subject, search_start_date, search_end_date):
    # Program breaks for these special character. We would need to replace them when making a request
    special_characters = ['+', '/', '?', '%', '#', '&']

    if sender:
        for character in special_characters:
            sender = sender.replace(character, urllib.parse.quote(character))
        sender_filter = "from/emailAddress/address eq \'%s\'" % sender
    else:
        sender_filter = ""

    if subject:
        for character in special_characters:
            subject = subject.replace(character, urllib.parse.quote(character))

        subject_filter = " and subject eq \'%s\'" % subject
    else:
        subject_filter = ""

    if search_start_date:
        search_start_date_filter = " and receivedDateTime ge %s" % search_start_date
    else:
        search_start_date_filter = ""

    if search_end_date:
        search_end_date_filter = " and receivedDateTime le %s" % search_end_date
    else:
        search_end_date_filter = ""

    # Combined filter of the above filters
    filter_query = "{0}{1}{2}{3}".format(sender_filter, subject_filter, search_start_date_filter, search_end_date_filter)

    # Remove " and " from the start of the combined filter
    if filter_query[:len(' and '):] == ' and ':
        filter_query = filter_query[len(' and ')::]

    return filter_query


parent_folder_id = {}


# Parses emails returned from get_email function to store as new email objects
def parse_good_email_response(mail, recipient):
    emailObj = OutlookEmail()
    emailObj.id = mail['id']
    emailObj.received_date = mail['receivedDateTime']
    emailObj.subject = mail['subject']
    emailObj.sender = mail['from']['emailAddress']['address']

    if len(mail['toRecipients']) == 1:
        emailObj.envelope_recipient = mail['toRecipients'][0]["emailAddress"]["address"]
    else:
        emailObj.envelope_recipient = "  "
    try:
        emailObj.ccrecipients = mail['ccRecipients']
    except Exception as e:
        emailObj.ccrecipients = " "
    try:
        emailObj.bccrecipients = mail['bccRecipients']
    except Exception as e:
        emailObj.bccrecipients = " "
    if mail['hasAttachments']:
        emailObj.has_attachments = True

    try:
        for section in mail['internetMessageHeaders']:
            if section['name'] == "Received":
                if "for <" in section['value']:
                    # print(section['value'])
                    emailObj.recipient = section['value'].split('for <')[1].split('>')[0]
                    break
        if emailObj.recipient is None:
            # logger.warning("Recipient not found in header. Using user email as recipient.")
            emailObj.recipient = recipient
    except Exception as e:
        emailObj.recipient = " "
    emailObj.requested_recipient = recipient
    emailObj.message_id = mail['internetMessageId']
    emailObj.body = mail['body']['content']
    try:
        emailObj.header = mail['internetMessageHeaders']
    except Exception as e:
        logger.info(
            "I am assuming its something to do with a calendar invite for user {0} with subject {1} to sender {2}".format(
                recipient, emailObj.subject, emailObj.sender))
    emailObj.email_read = mail['isRead']

    # Update the parent folder dictionary that stores folder id-> folder name for each recipient.
    # This helps minimizing api calls in case of rate limiting or broken api call or pagination etc.
    # Fetch folder name
    global parent_folder_id
    if emailObj.requested_recipient in parent_folder_id:
        parent_folder_id[emailObj.requested_recipient][mail['parentFolderId']] = get_folder_name(emailObj.requested_recipient, mail['parentFolderId'])
    else:
        parent_folder_id[emailObj.requested_recipient] = {}
        parent_folder_id[emailObj.requested_recipient][mail['parentFolderId']] = get_folder_name(
            emailObj.requested_recipient, mail['parentFolderId'])

    if parent_folder_id[emailObj.requested_recipient][mail['parentFolderId']] == 'Deleted Items' or parent_folder_id[emailObj.requested_recipient][mail['parentFolderId']] == 'Deletions':
        emailObj.in_deleteditems = True
    emails.append(emailObj)


# Search email for user in their mailbox. By default the search date is today and the search end time is None
# By default, it doesn't search recoverableitemsdeletions folder but searches deleted items folder
def get_email(sender, subject, search_start_date, search_end_date, recipients=None, pagination_urls=None, searchFolder=""):
    access_token = oauth_access_token
    headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}

    # Build the filter query to be used along with the url to search for the mail
    filter_query = build_get_mail_filter(sender, subject, search_start_date, search_end_date)

    index = 1  # To be used in the batch request creation
    json_batch_request = {"requests": []}  # Holds all batch requests that are to be sent as payload

    # The below 2 if statements creates the contents of the batch request
    # Change url to query for each request in the batch request based on if we are searching in a particular folder of the mailbox or from all folders in the mailbox
    if recipients is not None:
        for index, recipient in enumerate(recipients, start=index):
            if searchFolder:
                url = "/users/{0}/mailFolders/{1}/messages?$filter={2}&$select=id,receivedDateTime,from,subject,parentFolderId,body,internetMessageHeaders,isRead,toRecipients,internetMessageId,bccRecipients,ccRecipients,hasAttachments,isDraft&$count=true&$top=999".format(recipient, searchFolder, filter_query)
            else:
                url = "/users/{0}/messages?$filter={1}&$select=id,receivedDateTime,from,subject,parentFolderId,body,internetMessageHeaders,isRead,toRecipients,internetMessageId,bccRecipients,ccRecipients,hasAttachments,isDraft&$count=true&$top=999".format(recipient, filter_query)
            json_batch_request["requests"].append({"id": str(index), "method": "GET", "url": url})

    # If the function was called with a pagination url i.e. url pointing to the next set of results, the payload for the batch request
    if pagination_urls is not None:
        for _index, pagination_url in enumerate(pagination_urls, start=index):
            json_batch_request["requests"].append({"id": str(_index), "method": "GET", "url": pagination_url})

    block_pagination_urls = []  # A list to hold blocks of 20 pagination url requests

    session = session_generator()
    query_start_time = datetime.datetime.now()

    # Make the API call if token expiry time is greater than 1 minute
    if int((expiry_time - query_start_time).seconds) > 60:
        '''
        if batch_pagination_url:
            resp = session.post(batch_pagination_url, headers=headers)
        else:
            resp = session.post('https://graph.microsoft.com/v1.0/$batch', headers=headers, json=json_batch_request)
        '''
        resp = session.post('https://graph.microsoft.com/v1.0/$batch', headers=headers, json=json_batch_request)

        # Capture pagination for individual responses in the batch request. Scenario: Multiple recipients, multiple emails.
        # Usually pagination happens if there are 900+ mails returned from 1 recipient
        new_requests = {}
        new_requests["pagination_url"] = []
        new_requests["recipients"] = []

        if resp.ok:
            seconds_to_sleep = 0  # Used in handling rate limiting of the individual request in the batch request or individual broken api call
            for response in resp.json()['responses']:
                # Extract the url corresponding to the request from the batch request. This will contain the recipient.
                url = json_batch_request["requests"][int(response['id']) - 1]["url"]

                # Get the recipient from either the individual request response or from the associated request url
                if "@odata.context" in response["body"]:
                    recipient = urllib.parse.unquote(response["body"]["@odata.context"].replace("https://graph.microsoft.com/v1.0/$metadata#users('", "", 1).split("')/")[0])
                else:
                    recipient = url.split('/')[2]

                if response['status'] == 200:
                    total_number_of_emails_found = int(response["body"]['@odata.count'])
                    # Number of mails in the current page fetched
                    number_of_emails_in_response = len(response["body"]["value"])

                    # Not using the below next link in pagination as not all pages receive this link and it appears to be broken for pagination for batch requests
                    '''
                    if "@odata.nextLink" in response["body"]:
                        print(response["body"]["@odata.nextLink"])
                    '''

                    # Push the items fetched from the individual request to create mail objects populating their respective fields
                    if total_number_of_emails_found > 0:
                        with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                            for mail in response["body"]["value"]:
                                executor.submit(parse_good_email_response, mail, recipient)


                    # If there are no mails found for the recipient and no specific folder is told to be searched in, search the permanent deleted folder
                    if total_number_of_emails_found == 0 and not searchFolder:
                        logger.info("Unable to find mails for user {0}. Checking all mail folders....".format(recipient))

                        users = []  # Send the recipient as a list to fetch mails
                        users.append(recipient)

                        block_pagination_urls.extend(get_email(sender, subject, search_start_date,
                                                         search_end_date, users, pagination_urls,
                                                         searchFolder="recoverableitemsdeletions"))

                    if total_number_of_emails_found > 1 and pagination_urls is None:
                        logger.info("Total number of emails for %s: %d" % (recipient, total_number_of_emails_found))

                    # Create a list of pagination urls as chunks of page_size from the 1st response received
                    # Doing a recursive loop on pagination may end up eating all your resources and difficult to untangle from.
                    if (total_number_of_emails_found > number_of_emails_in_response) and pagination_urls is None:
                        parsed_url = url.split('&$top=')
                        url_without_skip = "{0}{1}{2:d}{3}".format(parsed_url[0], '&$top=', number_of_emails_in_response, "&$skip=")
                        # Found an issue when i scan through just drafts
                        if number_of_emails_in_response != 0:
                            new_requests["pagination_url"] = ["{0}{1:d}".format(url_without_skip, i) for i in range(number_of_emails_in_response, total_number_of_emails_found, number_of_emails_in_response)]

                    # For those emails in which the anticipated number of email responses after splitting them up into batch size request (20) is less than the actual email response
                    if pagination_urls is not None and number_of_emails_in_response != 0:
                        page_size = int(url.split('&$top=')[1].split('&$skip=')[0])
                        if number_of_emails_in_response != page_size:
                            url_without_page_size = url.split('&$top=')[0]
                            current_skip_size = int(url.split('&$skip=')[1])
                            pagination_url = "{0}{1}{2:d}{3}{4:d}".format(url_without_page_size, '&$top=', page_size - number_of_emails_in_response
                                                                       , "&$skip=", current_skip_size + number_of_emails_in_response)
                            new_requests["pagination_url"].append(pagination_url)

                # Handle rate limting for individual response by converting them to a pagination url or send the recipient to be sent in the next iteration
                elif response["status"] == 429:
                    seconds_to_sleep = int(response["headers"]['Retry-After']) - seconds_to_sleep
                    logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
                    if not seconds_to_sleep < 0:
                        time.sleep(seconds_to_sleep)
                    if not pagination_urls:
                        new_requests["recipients"].append(recipient)
                    else:
                        new_requests["pagination_url"].append(url)

                # Handle broken api call for individual response by converting them to a pagination url to be sent in the next iteration
                elif response["status"] == 503 or response["status"] == 504:
                    seconds_to_sleep = 1.5
                    logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
                    time.sleep(seconds_to_sleep)
                    if not pagination_urls:
                        new_requests["recipients"].append(recipient)
                    else:
                        new_requests["pagination_url"].append(url)

                # Handle other indiviudal response htttp error
                else:
                    logger.error("Failed to fetch mail for recipient %s ERROR %d:%s" % (recipient, response["status"], str(response["body"])))
                    logger.error("URL: %s" % url)

            # Only place to use recursive function.
            # Collect all recipients that failed or broken and create a new recursive function call for only those recipients
            recipients = new_requests["recipients"]
            if recipients:
                block_pagination_urls.extend(get_email(sender, subject, search_start_date, search_end_date, recipients))

            # Distribute the pagination urls to blocks of size 20 as size of batch request is 20
            if new_requests["pagination_url"]:
                if len(new_requests["pagination_url"]) > 20:
                    block_pagination_urls.extend(new_requests["pagination_url"][i:i+20] for i in range(0, len(new_requests["pagination_url"]), 20))
                else:
                    block_pagination_urls.append(new_requests["pagination_url"])
        # Handle Rate Limiting
        elif resp.status_code == 429:
            seconds_to_sleep = resp.headers['Retry-After']
            logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
            time.sleep(seconds_to_sleep)
            block_pagination_urls.extend(get_email(sender, subject, search_start_date, search_end_date, recipients, pagination_urls))
        # Handle broken api call
        elif resp.status_code == 503 or resp.status_code == 504:
            seconds_to_sleep = 1.5
            logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
            time.sleep(seconds_to_sleep)
            block_pagination_urls.extend(get_email(sender, subject, search_start_date, search_end_date, recipients, pagination_urls))
        # Handle other http errors
        else:
            logger.error("Error processing batch request to get emails for recipients")
            logger.info("Batch request :%s" % json.dumps(json_batch_request["requests"], indent=4))
            logger.error("ERROR %d:%s" % (resp.status_code, resp.json()))
    # Create new access token
    else:
        check_and_renew_access_token(access_token)
        block_pagination_urls.extend(get_email(sender, subject, search_start_date, search_end_date, recipients, pagination_urls))
    return block_pagination_urls


# Removes duplicate entries from the mail recipients
def remove_duplicate_email_entries(recipients):
    logger.info("Removing duplicate entries from the recipient list")
    return list(set(recipients))


# Print the email objects
def print_all_mails_found(showDeletedMails=False):
    index = 0

    if not showDeletedMails:
        print(
            '\nIndex| Subject| Sender| Mailbox| Header Recipient| Envelope Recipient| Read| Received Date| ccRecipients| bccRecipients| Message ID| hasAttachment')
        for email in filtered_emails:
            index = index + 1
            print("{0}| {1}| {2}| {3}| {4}| {5}| {6}| {7}| {8}| {9}| {10}| {11}".format(index, email.subject,
                                                                                        email.sender,
                                                                                        email.requested_recipient,
                                                                                        email.recipient,
                                                                                        email.envelope_recipient,
                                                                                        email.email_read,
                                                                                        email.received_date,
                                                                                        str(email.ccrecipients),
                                                                                        str(email.bccrecipients),
                                                                                        email.message_id,
                                                                                        email.has_attachments))
    else:
        print(
            '\nIndex| Subject| Sender| Mailbox| Header Recipient| Envelope Recipient| Read| Received Date| ccRecipients| bccRecipients| Message ID| hasAttachment| Deleted')
        for email in emails:
            index = index + 1
            print("{0}| {1}| {2}| {3}| {4}| {5}| {6}| {7}| {8}| {9}| {10}| {11}| {12}".format(index, email.subject,
                                                                                              email.sender,
                                                                                              email.requested_recipient,
                                                                                              email.recipient,
                                                                                              email.envelope_recipient,
                                                                                              email.email_read,
                                                                                              email.received_date,
                                                                                              str(email.ccrecipients),
                                                                                              str(email.bccrecipients),
                                                                                              email.message_id,
                                                                                              email.has_attachments,
                                                                                              email.in_deleteditems))
    print()


# Get all your company employees
def list_all_active_users(pagination_url=""):
    session = session_generator()
    query_start_time = datetime.datetime.now()
    access_token = oauth_access_token
    headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}

    # Make the API call if token expiry time is greater than 1 minute
    if int((expiry_time - query_start_time).seconds) > 60:
        if not pagination_url:
            url = graph_api.format('users/?$select=id,mail,accountEnabled,userPrincipalName&$top=999&$count=true')
        else:
            url = pagination_url

        resp = session.get(url, headers=headers)

        if resp.ok:
            response = resp.json()
            if response['value']:
                for user in response['value']:
                    if user['accountEnabled'] and user['mail']:
                        # Check if end of the string contains company domain
                        if user['userPrincipalName'][-len('@%s' % company_domain)::] == '@%s' % company_domain:
                            active_employee_usernames.append(user['userPrincipalName'])

            # Appears if next page of results are present
            if '@odata.nextLink' in response:
                pagination_url = response['@odata.nextLink']
            else:
                pagination_url = ""
        # Handle rate limiting
        elif resp.status_code == 429:
            seconds_to_sleep = resp.headers['Retry-After']
            logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
            time.sleep(seconds_to_sleep)
            pagination_url = list_all_active_users(pagination_url)
        # Handle broken api call
        elif resp.status_code == 503 or resp.status_code == 504:
            seconds_to_sleep = 1
            logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
            time.sleep(seconds_to_sleep)
            pagination_url = list_all_active_users(pagination_url)
        # Handle other http errors
        else:
            logger.error("Unable to get a list of all company employees ")
            logger.error("ERROR %d:%s" % (resp.status_code, resp.text))
            pagination_url = ""
    # Create new access token
    else:
        check_and_renew_access_token(access_token)
        pagination_url = list_all_active_users(pagination_url)

    return pagination_url


# Map email to UserPrincipal Name. This username will be used while fetching mails
def map_email_to_username(mail):
    userPrincipalName = ""  # Default value
    session = session_generator()
    access_token = oauth_access_token
    headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
    url = graph_api.format('users?$filter=mail eq \'{0}\' or userPrincipalName eq \'{1}\'&$select=id,mail,accountEnabled,userPrincipalName')
    query_start_time = datetime.datetime.now()

    # Make the API call if token expiry time is greater than 1 minute
    if int((expiry_time - query_start_time).seconds) > 60:
        resp = session.get(url.format(mail, mail), headers=headers)
        if resp.ok:
            response = resp.json()
            if response['value']:
                if response['value'][0]['accountEnabled']:
                    userPrincipalName = response['value'][0]['userPrincipalName']
                    if '@%s.onmicrosoft.com' % company_domain.split('.')[0] in userPrincipalName:
                        userPrincipalName = mail
        # Handle rate limiting
        elif resp.status_code == 429:
            seconds_to_sleep = resp.headers['Retry-After']
            logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
            time.sleep(seconds_to_sleep)
            userPrincipalName = map_email_to_username(mail)
        # Handle broken api call
        elif resp.status_code == 503 or resp.status_code == 504:
            seconds_to_sleep = 1
            logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
            time.sleep(seconds_to_sleep)
            userPrincipalName = map_email_to_username(mail)
        # Handle other errors
        else:
            logger.error("Mail {0} is causing issues and cannot be mapped to username".format(mail))
            logger.error("ERROR %d:%s" % (resp.status_code, resp.text))
    # Create new access token
    else:
        check_and_renew_access_token(access_token)
        userPrincipalName = map_email_to_username(mail)

    return userPrincipalName


# Check if the dl exists
def check_if_its_dl(recipient):
    dl_id = ""  # Default value of DL id
    access_token = oauth_access_token
    headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
    url = graph_api.format('groups?$filter=mail eq \'{1}\'&$select=id,displayName,groupTypes,mail')
    session = session_generator()
    query_start_time = datetime.datetime.now()

    # Make the API call if token expiry time is greater than 1 minute
    if int((expiry_time - query_start_time).seconds) > 60:
        resp = session.get(url.format(recipient.split('@')[0], recipient), headers=headers)
        if resp.ok:
            response = resp.json()
            if len(response['value']) == 1:
                dl_id = response['value'][0]['id']
        # Handle Rate limiting
        elif resp.status_code == 429:
            seconds_to_sleep = resp.headers['Retry-After']
            logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
            time.sleep(seconds_to_sleep)
            dl_id = check_if_its_dl(recipient)
        # Handle broken api call
        elif resp.status_code == 503 or resp.status_code == 504:
            seconds_to_sleep = 0.5
            logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
            time.sleep(seconds_to_sleep)
            dl_id = check_if_its_dl(recipient)
        # Handle other errors
        else:
            logger.error("Unable to resolve DL/recipient %s" % recipient)
            logger.error("%d:%s" % (resp.status_code, resp.text))
    # Create new access token to be used
    else:
        check_and_renew_access_token(access_token)
        check_if_its_dl(recipient)

    return dl_id


# Get all members in a DL
def get_recipients_from_dl(dl_id, pagination_url=""):
    recipients = []  # All individual mailboxes/users
    groups = []  # If the DL contains another DL
    access_token = oauth_access_token
    headers = {"Authorization": "Bearer %s" % access_token, "Content-Type": "application/json; charset=utf-8"}
    session = session_generator()
    query_start_time = datetime.datetime.now()

    # Make the API call if token expiry time is greater than 1 minute
    if int((expiry_time - query_start_time).seconds) > 60:
        if not pagination_url:
            url = graph_api.format('groups/{0}/members?$top=999')
            resp = session.get(url.format(dl_id), headers=headers)
        else:
            url = pagination_url
            resp = session.get(url, headers=headers)
        if resp.ok:
            response = resp.json()

            for member in response['value']:
                # If it contains a DL
                if 'group' in member['@odata.type']:
                    groups.append(member['id'])

                # For individual mailbox/users
                if 'user' in member['@odata.type']:
                    recipients.append(member['userPrincipalName'])

            # Make recursive calls if a DL contains another DL
            with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
                fs = [executor.submit(get_recipients_from_dl, group) for group in groups]
                block_of_futures = []
                if len(fs) > 15:
                    block_of_futures = [fs[i:i+15] for i in range(0, len(fs), 15)]
                else:
                    block_of_futures.append(fs)
                for futures in block_of_futures:
                    if futures:
                        for future in concurrent.futures.as_completed(futures):
                            recipients.extend(future.result())

            # Appears if next page of results are present
            if "@odata.nextLink" in response:
                pagination_url = response["@odata.nextLink"]
                recipients.extend(get_recipients_from_dl(dl_id, pagination_url))

        # Handle Rate limiting
        elif resp.status_code == 429:
            seconds_to_sleep = resp.headers['Retry-After']
            logger.warning('Throttling threshold reached. Sleeping for {0:d} seconds'.format(seconds_to_sleep))
            time.sleep(seconds_to_sleep)
            recipients.extend(get_recipients_from_dl(dl_id))
        # Handle broken api call
        elif resp.status_code == 503 or resp.status_code == 504:
            seconds_to_sleep = 0.5
            logger.warning("Experiencing 504 error i.e. connection or gateway timeouts")
            time.sleep(seconds_to_sleep)
            recipients.extend(get_recipients_from_dl(dl_id))
        # Handle other errors
        else:
            logger.error("Unable to get recipients for a DL id %s" % str(dl_id))
            logger.error("%d:%s" % (resp.status_code, resp.text))
    # Create new access token
    else:
        check_and_renew_access_token(access_token)
        recipients.extend(get_recipients_from_dl(dl_id))
    return recipients


# Check start date and end date logic
def check_date(start_date, end_date):
    if not start_date:
        start_date = datetime.date.today().strftime("%Y-%m-%dT%H:%M:%SZ")
    if not end_date:
        end_date = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

    if datetime.datetime.strptime(end_date, "%Y-%m-%dT%H:%M:%SZ") < datetime.datetime.strptime(start_date, "%Y-%m-%dT%H:%M:%SZ"):
        logger.critical("Start date cannot be greater than end date")
        exit(1)


# Check if employee still works in the company
def recipient_exits_check(recipient):
    recipients = []  # A list of recipients that still work in the company

    # Get the username associated with the email address
    username = map_email_to_username(recipient)
    if username:
        recipients.append(username)

    else:
        # Might be a DL
        dl_id = check_if_its_dl(recipient)
        if dl_id:
            recipients_from_dl = get_recipients_from_dl(dl_id)
            if not recipients_from_dl:
                # For DL containing 0 members
                logger.info("No recipients found for {0}".format(recipient))
            else:
                recipients.extend(recipients_from_dl)  # Add members of dl
        else:
            logger.info("{0} not a Email DL nor a user".format(recipient))

    return recipients


# Convert user input entered from command line or from another script into a consumable variable
def format_user_input(recipients, start_date, end_date):
    # Add the Time stamp as that is not added from the command line
    if start_date:
        start_date = "{}T00:00:00Z".format(start_date)
    if end_date:
        end_date = "{}T23:59:59Z".format(end_date)

    # Verify if the start date is less than the end_date
    if start_date and end_date:
        check_date(start_date, end_date)  # Check if start date < end_date

    # Parse recipients into a list
    if isinstance(recipients, str):
        # Check if there is a single recipient or multiple recipients in the recipients string
        if ', ' in recipients:
            recipients = recipients.strip('\n').split(', ')
        elif ',' in recipients:
            recipients = recipients.strip('\n').split(',')
        elif ' ' in recipients:
            recipients = recipients.strip('\n').split(' ')
        else:
            # Convert recipients string to list for one member
            recipient = recipients
            recipients = []
            recipients.append(recipient)
    elif not isinstance(recipients, list):
        logger.critical("Recipients should be either a list or string. Exiting.")
        recipients = None

    return recipients, start_date, end_date


# Pull Emails
def email_pull(sender, recipients, subject, start_date="", end_date="", skip_recipient_check=False):
    # Purge global lists so that re-using the script doesn't cause conflict and display weird behavior
    global emails
    emails = []
    global filtered_emails
    filtered_emails = []

    # Format user input into lists and add time info to start and end date if necessary
    recipients, start_date, end_date = format_user_input(recipients, start_date, end_date)

    # Skip the recipient verification if you are sure that the recipients exist and have a mailbox
    if not skip_recipient_check:
        # Verify and generate recipient list including resolving dls. It has to be a list even if its a single recipient
        if len(recipients) == 1 and not recipients[0]:  # If recipients input is blank, get all users in company that have a mailbox
            logger.info("Generating list of all active users")
            pagination_url = list_all_active_users()
            while pagination_url:
                pagination_url = list_all_active_users(pagination_url)
            logger.info("Completed acquiring all active employee usernames")
            recipients = active_employee_usernames
        else:
            new_recipient_list = []
            for recipient in recipients:
                recipients_from_check = recipient_exits_check(recipient)
                if recipients_from_check:
                    new_recipient_list.extend(recipients_from_check)
            if new_recipient_list:  # Overwrite recipients variable from user input with the verified recipients
                recipients = new_recipient_list
            else:
                recipients = []

    if not recipients:
        logger.info("No recipients received the mail with subject {0}".format(subject))
        success_failure_message = "No recipients received the mail with subject '{0}'".format(subject)
        return success_failure_message

    # Remove duplicate entries of recipients from user input
    recipients = remove_duplicate_email_entries(recipients)  # remove duplicate entries of recipients from user input

    # Divide the recipient list into blocks of 20 recipients. 20 is the limit for batch requests to Microsoft Graph API
    if len(recipients) > 20:
        blocks_of_recipients = [recipients[i:i + 20] for i in range(0, len(recipients), 20)]
    else:
        blocks_of_recipients = []
        blocks_of_recipients.append(recipients)

    logger.info("Getting mails")

    # Store blocks of 20 pagination urls that is returned from the get email function to call them again to get new mails
    block_pagination_urls = []

    # Fetch mails concurrently for each block of 20 from the block recipients list
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        fs = [executor.submit(get_email, sender, subject, start_date, end_date, recipient)
              for recipient in blocks_of_recipients]
        block_of_futures = []
        if len(fs) > 15:
            block_of_futures = [fs[i:i+15] for i in range(0, len(fs), 15)]
        else:
            block_of_futures.append(fs)
        for futures in block_of_futures:
            if futures:
                for future in concurrent.futures.as_completed(futures):
                    block_pagination_urls.extend(future.result())

    # If pagination urls are returned by the above execution, run them to fetch more mails until they stop returning pagination urls
    while block_pagination_urls:
        another_pagination_urls = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            fs = [executor.submit(get_email, sender, subject, start_date, end_date, None, pagination_urls) for pagination_urls in block_pagination_urls]
            block_of_futures = []
            if len(fs) > 15:
                block_of_futures = [fs[i:i+15] for i in range(0, len(fs), 15)]
            else:
                block_of_futures.append(fs)
            for futures in block_of_futures:
                if futures:
                    for future in concurrent.futures.as_completed(futures):
                        another_pagination_urls.extend(future.result())

            block_pagination_urls = another_pagination_urls

    # If no mails were fetched, exit
    if not emails:
        logger.info("Email not found in the recipients mailboxes")
        success_failure_message = "Email not found in the {0} mailboxes for subject '{1}'".format(str(recipients).strip('[').strip(']'), subject)
        return success_failure_message

    # Remove emails from the list that are already deleted or are in recipient Trash and push them to filtered_emails list
    logger.info("Filtering emails that are not deleted.")
    filtered_emails = [mail for mail in emails if not mail.in_deleteditems]

    if filtered_emails:
        emails = []
        try:
            initial_filtered_email_recipients = []
            recipients_with_deleted_mail = []

            # Start sending the mails to Trash for each recipient
            for mail in filtered_emails:
                initial_filtered_email_recipients.append(mail.requested_recipient)
                if mail.move_email_to_folder('deleteditems'):
                    recipients_with_deleted_mail.append(mail.requested_recipient)

            logger.info("Email Pull successful. Mail is present in {0} Deleted Folder".format(str(recipients_with_deleted_mail).strip('[').strip(']')))
            success_failure_message = "Email Pull successful for subject '{0}'. Mail is present in {1} Deleted Folder".format(subject, str(recipients_with_deleted_mail).strip('[').strip(']'))

            if recipients.sort() != initial_filtered_email_recipients.sort():
                recipients_with_deleted_mail_from_start = []
                for recipient in recipients:
                    if recipient not in initial_filtered_email_recipients:
                        recipients_with_deleted_mail_from_start.append(recipient)
                logger.info("Recipients {0} has already deleted the mail.".format(str(recipients_with_deleted_mail_from_start).strip('[').strip(']')))
                success_failure_message = success_failure_message + ".\nRecipients {0} have already deleted the mail for subject '{1}'.".format(str(recipients_with_deleted_mail_from_start).strip('[').strip(']'), subject)

            if initial_filtered_email_recipients.sort() != recipients_with_deleted_mail.sort():
                unable_to_delete_mail_recipients = list(set(initial_filtered_email_recipients) - set(recipients_with_deleted_mail))
                logger.error('Unable to delete mail for {0}'.format(str(unable_to_delete_mail_recipients).strip('[').strip(']')))
                success_failure_message = success_failure_message + ".\nUnable to delete mail for {0}".format(str(unable_to_delete_mail_recipients).strip('[').strip(']'))

            return success_failure_message

        except Exception as e:
            logger.error(e)
            logger.critical("Ran into error. Run pull script again for recipient {0}".format(str(recipients).strip('[').strip(']')))
            success_failure_message = "Ran into error. Run pull script again for recipient {0} with subject '{1}'".format(str(recipients).strip('[').strip(']'), subject)
            return success_failure_message

    else:
        # All the recipients have already deleted the mail.
        recipients = []  # The recipient list is not required anymore, and hence is being overwritten to get the list of all recipients whose mail was deleted
        for email in emails:
            recipients.append(email.requested_recipient)
        recipients = list(set(recipients))
        logger.info("Recipients {0} has already deleted the mail.".format(str(recipients).strip('[').strip(']')))
        success_failure_message = "Recipients {0} have already deleted the mail for subject '{1}'.".format(str(recipients).strip('[').strip(']'), subject)
        return success_failure_message


# Pull Emails
def email_restore(sender, recipients, subject, start_date="", end_date="", skip_recipient_check=False):
    # Purge global lists so that re-using the script doesn't cause conflict and display weird behavior
    global emails
    emails = []
    global filtered_deleted_emails
    filtered_deleted_emails = []

    # Format user input into lists and add time info to start and end date if necessary
    recipients, start_date, end_date = format_user_input(recipients, start_date, end_date)

    # Skip the recipient verification if you are sure that the recipients exist and have a mailbox
    if not skip_recipient_check:
        # Verify and generate recipient list including resolving dls. It has to be a list even if its a single recipient
        if len(recipients) == 1 and not recipients[0]:  # If recipients input is blank, get all users in your company that have a mailbox
            logger.info("Generating list of all active users")
            pagination_url = list_all_active_users()
            while pagination_url:
                pagination_url = list_all_active_users(pagination_url)
            logger.info("Completed acquiring all active employee usernames")
            recipients = active_employee_usernames
        else:
            new_recipient_list = []
            for recipient in recipients:
                recipients_from_check = recipient_exits_check(recipient)
                if recipients_from_check:
                    new_recipient_list.extend(recipients_from_check)
            if new_recipient_list:  # Overwrite recipients variable from user input with the verified recipients
                recipients = new_recipient_list
            else:
                recipients = []

    if not recipients:
        logger.info("No recipients received the mail with subject {0}".format(subject))
        success_failure_message = "No recipients received the mail with subject '{0}'".format(subject)
        return success_failure_message

    # Remove duplicate entries of recipients from user input
    recipients = remove_duplicate_email_entries(recipients)  # remove duplicate entries of recipients from user input

    # Divide the recipient list into blocks of 20 recipients. 20 is the limit for batch requests to Microsoft Graph API
    if len(recipients) > 20:
        blocks_of_recipients = [recipients[i:i + 20] for i in range(0, len(recipients), 20)]
    else:
        blocks_of_recipients = []
        blocks_of_recipients.append(recipients)

    logger.info("Getting mails")

    # Store blocks of 20 pagination urls that is returned from the get email function to call them again to get new mails
    block_pagination_urls = []

    # Fetch mails concurrently for each block of 20 from the block recipients list
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        fs = [executor.submit(get_email, sender, subject, start_date, end_date, recipient)
              for recipient in blocks_of_recipients]
        block_of_futures = []
        if len(fs) > 15:
            block_of_futures = [fs[i:i+15] for i in range(0, len(fs), 15)]
        else:
            block_of_futures.append(fs)
        for futures in block_of_futures:
            if futures:
                for future in concurrent.futures.as_completed(futures):
                    block_pagination_urls.extend(future.result())

    # If pagination urls are returned by the above execution, run them to fetch more mails until they stop returning pagination urls
    while block_pagination_urls:
        another_pagination_urls = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            fs = [executor.submit(get_email, sender, subject, start_date, end_date, None, pagination_urls) for pagination_urls in block_pagination_urls]
            block_of_futures = []
            if len(fs) > 15:
                block_of_futures = [fs[i:i+15] for i in range(0, len(fs), 15)]
            else:
                block_of_futures.append(fs)
            for futures in block_of_futures:
                if futures:
                    for future in concurrent.futures.as_completed(futures):
                        another_pagination_urls.extend(future.result())

            block_pagination_urls = another_pagination_urls

    # If no mails were fetched, exit
    if not emails:
        logger.info("Email not found in the recipients mailboxes")
        success_failure_message = "Email not found in the {0} mailboxes for subject '{1}'".format(str(recipients).strip('[').strip(']'), subject)
        return success_failure_message

    # Remove emails from the list that are already deleted or are in recipient Trash and push them to filtered_emails list
    logger.info("Filtering emails that are deleted.")
    filtered_deleted_emails = [mail for mail in emails if mail.in_deleteditems]

    if filtered_deleted_emails:
        emails = []
        try:
            initial_filtered_email_recipients = []
            recipients_with_restored_mail = []

            # Start sending the mails from Trash for each recipient
            for mail in filtered_emails:
                initial_filtered_email_recipients.append(mail.requested_recipient)
                if mail.move_email_to_folder('inbox'):
                    recipients_with_restored_mail.append(mail.requested_recipient)

            logger.info("Email Restore successful. Mail is present in {0} Inbox Folder".format(str(recipients_with_restored_mail).strip('[').strip(']')))
            success_failure_message = "Email Restore successful for subject '{0}'. Mail is present in {1} Inbox Folder".format(subject, str(recipients_with_restored_mail).strip('[').strip(']'))

            if recipients.sort() != initial_filtered_email_recipients.sort():
                recipients_with_restored_mail_from_start = []
                for recipient in recipients:
                    if recipient not in initial_filtered_email_recipients:
                        recipients_with_restored_mail_from_start.append(recipient)
                logger.info("Recipients {0} has already restored the mail.".format(str(recipients_with_restored_mail_from_start).strip('[').strip(']')))
                success_failure_message = success_failure_message + ".\nRecipients {0} have already restored the mail for subject '{1}'.".format(str(recipients_with_restored_mail_from_start).strip('[').strip(']'), subject)

            if initial_filtered_email_recipients.sort() != recipients_with_restored_mail.sort():
                unable_to_restore_mail_recipients = list(set(initial_filtered_email_recipients) - set(recipients_with_restored_mail))
                logger.error('Unable to restore mail for {0}'.format(str(unable_to_restore_mail_recipients).strip('[').strip(']')))
                success_failure_message = success_failure_message + ".\nUnable to restore mail for {0}".format(str(unable_to_restore_mail_recipients).strip('[').strip(']'))

            return success_failure_message
        except Exception as e:
            logger.error(e)
            logger.critical("Ran into error. Run restore script again for recipient {0}".format(str(recipients).strip('[').strip(']')))
            success_failure_message = "Ran into error. Run restore script again for recipient {0} with subject '{1}'".format(str(recipients).strip('[').strip(']'), subject)
            return success_failure_message
    else:
        # All the recipients have already deleted the mail.
        recipients = [email.requested_recipient for email in emails]  # The recipient list is not required anymore, and hence is being overwritten to get the list of all recipients whose mail was deleted
        recipients = list(set(recipients))
        logger.info("Recipients {0} has already restored the mail.".format(str(recipients).strip('[').strip(']')))
        success_failure_message = "Recipients {0} have already restored the mail for subject '{1}'.".format(str(recipients).strip('[').strip(']'), subject)
        return success_failure_message
