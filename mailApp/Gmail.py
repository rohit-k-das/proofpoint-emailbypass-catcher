import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import mailApp.Directory as directory_api
import datetime
import json
import logging
import time
import concurrent.futures
import jwt
import re
import ConfigParser
import os

logger = logging.getLogger(__name__)

MAX_THREADS = 14  # Get max number of threads for multi-threading
gmail_api = 'https://www.googleapis.com/gmail/v1/users'

Config = ConfigParser.ConfigParser()
Config.read(os.path.join(os.path.abspath(os.path.dirname(__file__)),'Mail_creds'))
google_service_account_secret = Config.get('Settings', 'Google_Service_Account_Secret')
google_service_account_id = Config.get('Settings', 'Google_Service_Account_ID')
google_user_for_service_account = Config.get('Settings', 'Google_User_For_Project')

gmail_emails = []
gmail_filtered_emails = []
gmail_filtered_deleted_emails = []
access_tokens = {}


# Generate session with max of 3 retries and interval of 1 second
def session_generator():
    session = requests.Session()
    retry = Retry(connect=3, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session


class Gmail:
    def __init__(self):
        self.sender = None
        self.requested_recipient = None
        self.recipient = None
        self.envelope_recipient = None
        self.in_deleteditems = False
        self.body = None
        self.ccrecipients = None
        self.bccrecipients = None
        self.message_id = None
        self.has_attachments = False
        self.received_date = None
        self.id = None
        self.email_read = False
        self.subject = None
        self.header = None

    # Send mail to Trash
    def delete_mail(self):
        status = False
        access_token = access_tokens[self.requested_recipient]['access_token']
        expiry = access_tokens[self.requested_recipient]['expiry']
        headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer %s' % access_token}
        session = session_generator()
        query_start_time = time.time()

        # Check if there is more than a minute left for the access token to expire
        if (expiry - query_start_time) > 60:  # Check if there is more than a minute left for the access token to expire
            resp = session.post("%s/%s/messages/%s/trash" % (gmail_api, self.requested_recipient, self.id), headers=headers)
            if resp.ok:
                status = True
            # Rate limiting
            elif resp.status_code == 429:
                logger.error('Too many requests. Sleeping %s' % resp.json()['error']['message'])
                time.sleep(1)
                status = self.delete_mail()
            # Handle other http errors
            else:
                logger.error("Unable to delete mail for %s with subject %s" % (self.requested_recipient, self.subject))
                logger.error("%d:%s" % (resp.status_code, resp.text))

        # Create new access token to be used by the recipient
        else:
            access_token, expiry = generate_access_token(self.requested_recipient)
            if access_token is not None and expiry is not None:
                access_tokens[self.requested_recipient]['access_token'] = access_token
                access_tokens[self.requested_recipient]['expiry'] = expiry
                status = self.delete_mail()
        return status

    # Recover mail from Trash
    def undelete_mail(self):
        status = False
        access_token = access_tokens[self.requested_recipient]['access_token']
        expiry = access_tokens[self.requested_recipient]['expiry']
        headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer %s' % access_token}
        session = session_generator()
        query_start_time = time.time()

        # Check if there is more than a minute left for the access token to expire
        if (expiry - query_start_time) > 60:
            resp = session.post("%s/%s/messages/%s/untrash" % (gmail_api, self.requested_recipient, self.id), headers=headers)
            if resp.ok:
                status = True
            # Rate limiting
            elif resp.status_code == 429:
                logger.error('Too many requests. Sleeping %s' % resp.json()['error']['message'])
                time.sleep(1)
                status = self.undelete_mail()
            else:
                logger.error("Unable to recover mail for %s with subject %s from Trash" % (self.requested_recipient, self.subject))
                logger.error("%d:%s" % (resp.status_code, resp.text))

        # Create new access token to be used by the recipient
        else:
            access_token, expiry = generate_access_token(self.requested_recipient)
            if access_token is not None and expiry is not None:
                access_tokens[self.requested_recipient]['access_token'] = access_token
                access_tokens[self.requested_recipient]['expiry'] = expiry
                status = self.undelete_mail()
        return status


# Create OAuth token per requirement for each recipient
def generate_access_token(recipient, need_write_access=False):
    access_token = None
    expiry = None
    jwt_header = {"alg": "RS256", "typ": "JWT"}
    iat = time.time()
    exp = iat + 3600
    jwt_claim_set = {
        'iss': google_service_account_id,
        'scope': 'https://www.googleapis.com/auth/gmail.readonly',
        'sub': recipient,
        'aud': 'https://www.googleapis.com/oauth2/v4/token',
        'iat': iat,
        'exp': exp}
    if need_write_access:
        jwt_claim_set['scope'] = 'https://www.googleapis.com/auth/gmail.modify'
    secret = bytes(google_service_account_secret.replace('\\n', '\n'), 'utf-8')
    signed_jwt = jwt.encode(jwt_claim_set, secret, headers=jwt_header, algorithm='RS256')

    headers = {"Content-Type": "application/json; charset=utf-8"}
    data = {'grant_type': 'urn:ietf:params:oauth:grant-type:jwt-bearer', 'assertion': signed_jwt.decode('utf-8').replace("'", '"')}
    url = 'https://www.googleapis.com/oauth2/v4/token'
    session = session_generator()
    resp = session.post(url, headers=headers, data=json.dumps(data))
    if resp.ok:
        response = resp.json()
        access_token = response['access_token']
        expiry = time.time() + response['expires_in']
    elif resp.status_code == 400 and "Invalid email" in resp.json()['error_description']:
        logger.info("Recipient %s not found" % recipient)
    elif resp.status_code == 429:
        logger.error('Too many requests. Sleeping %s' % resp.json()['error']['message'])
        time.sleep(1)
        access_token, expiry = generate_access_token(recipient, need_write_access)
    else:
        logger.error('Failed to generate access token')
        logger.error("%d:%s" % (resp.status_code, resp.text))
    return access_token, expiry


# Populate the email obj with details of the mail from mail id
def populate_emails(mail):
    access_token = access_tokens[mail.requested_recipient]['access_token']
    expiry = access_tokens[mail.requested_recipient]['expiry']
    query_start_time = time.time()

    # Make the API call if token expiry time is greater than 1 minute
    if (expiry - query_start_time) > 60:
        headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer %s' % access_token}
        params = {'format': 'metadata', 'metadataHeaders': ['Received', 'From', 'To', 'Subject', 'Date', 'X-MS-Has-Attach', 'Message-ID']}
        session = session_generator()
        resp = session.get("%s/%s/messages/%s" % (gmail_api, mail.requested_recipient, mail.id), headers=headers, params=params)
        if resp.ok:
            response = resp.json()

            # Fill in the parameters of the email object from the response
            mail.header = response['payload']['headers']
            if 'TRASH' in response['labelIds']:
                mail.in_deleteditems = True
            if 'UNREAD' in response['labelIds']:
                mail.email_read = False
            for section in mail.header:
                if section['name'] == 'Received' and 'for <' in section['value']:
                    mail.recipient = section['value'].split('for <')[1].split('>')[0]
                if section['name'] == 'From':
                    sender = section['value'].split('<')[1].split('>')[0]
                    if mail.sender is None or mail.sender != sender:
                        mail.sender = sender
                if section['name'] == 'To':
                    mail.envelope_recipient = str(re.findall(r"\<(\S+)\>", section['value'], flags=re.I)).strip('[').strip(']')
                if section['name'] == 'Subject':
                    mail.subject = section['value']
                if section['name'] == 'Date':
                    mail.received_date = section['value'].split(',')[1]
                if section['name'] == 'X-MS-Has-Attach' and section['value'] == 'yes':
                    mail.has_attachments = True
                if section['name'] == 'Message-ID':
                    mail.message_id = section['value'].strip('<').strip('>')
        # Rate limiting
        elif resp.status_code == 429:
            logger.error('Too many requests. Sleeping %s' % resp.json()['error']['message'])
            time.sleep(1)
            populate_emails(mail)
        # Handle other http errors
        else:
            logger.error("Unable to get mail for %s" % mail.requested_recipient)
            logger.error("%d:%s" % (resp.status_code, resp.text))
    # Create new access token to be used by the recipient
    else:
        access_token, expiry = generate_access_token(mail.requested_recipient)
        if access_token is not None and expiry is not None:
            access_tokens[mail.requested_recipient]['access_token'] = access_token
            access_tokens[mail.requested_recipient]['expiry'] = expiry
            populate_emails(mail)


# Print the email objects
def print_all_mails_found(emails):
    if emails:
        print(
            '\nIndex| Subject| Sender| Requested Recipient| Header Recipient| Envelope Recipient| Read| Received Date| ccRecipients| bccRecipients| Message ID| hasAttachment')
        for index, email in enumerate(emails, start=1):
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

    print()


# Check input date
def check_date(start_date, end_date):
    if datetime.datetime.strptime(end_date, "%Y/%m/%d") < datetime.datetime.strptime(start_date, "%Y/%m/%d"):
        logger.critical("Start date cannot be greater than end date")
        exit(1)


# Convert recipients into a list
def format_user_input(recipients, start_date, end_date):
    if start_date:
        start_date = start_date.replace('-', '/')
    if end_date:
        end_date = end_date.replace('-', '/')

    if start_date and end_date:
        check_date(start_date, end_date)  # Check if start date < end_date

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


# Get mails with the filter criteria and return mail_ids corresponding to them. PageToken is used for pagination
def get_mail_ids(recipient="", subject="", start_date="", end_date="", sender="", pagination_url="", only_has_attachments=False):
    if not recipient:
        if pagination_url:
            recipient = pagination_url.split('users/')[1].split('/messages')[0]
        else:
            logger.error('Wrong usage of function. Exiting..')
            exit(-1)
    access_token = access_tokens[recipient]['access_token']
    expiry = access_tokens[recipient]['expiry']
    query_start_time = time.time()

    # Make the API call if token expiry time is greater than 1 minute
    if (expiry - query_start_time) > 60:
        headers = {'Content-Type': 'application/json', 'Authorization': 'Bearer %s' % access_token}

        # Create request based on whether its a pagination or not
        if not pagination_url:
            params = {'maxResults': 1000, 'includeSpamTrash': True}

            # Create Filter to search for specific mails that fit the criteria
            if subject or sender or start_date or end_date or only_has_attachments:
                filter = ''
                if subject:
                    filter = '%s subject:%s' % (filter, subject)

                if sender:
                    filter = '%s {from:%s list:%s}' % (filter, sender, sender)

                if start_date:
                    filter = '%s after:%s' % (filter, start_date)

                if end_date:
                    filter = '%s before:%s' % (filter, end_date)

                if only_has_attachments:
                    filter = '%s has:attachment'

                if filter:
                    params['q'] = filter[1:]

            session = session_generator()
            resp = session.get("%s/%s/messages" % (gmail_api, recipient), headers=headers, params=params)
        else:
            session = session_generator()
            resp = session.get(pagination_url, headers=headers)

        if resp.ok:
            response = resp.json()
            if response['resultSizeEstimate'] != 0:
                if response['messages']:
                    for mail in response['messages']:
                        mail_id = mail['id']
                        gmail_obj = Gmail()
                        gmail_obj.id = mail_id
                        gmail_obj.requested_recipient = recipient
                        gmail_emails.append(gmail_obj)

                if not pagination_url and not response['messages']:
                    logger.info("0 mails found for %s" % recipient)

            # Pagination
            if 'nextPageToken' in response:
                pageToken = response['nextPageToken']
                if 'pageToken' in resp.url:
                    pagination_url = '{0}&pageToken={1}'.format(resp.url.split('&pageToken')[0], pageToken)
                else:
                    pagination_url = '{0}&pageToken={1}'.format(resp.url, pageToken)
            else:
                pagination_url = ''
        # Rate limiting
        elif resp.status_code == 429:
            logger.error('Too many requests. Sleeping %s' % resp.json()['error']['message'])
            time.sleep(1)
            pagination_url = get_mail_ids(recipient, subject, start_date, end_date, sender, pagination_url, only_has_attachments)
        # Handle other http errors
        else:
            logger.error("Unable to get mail for %s" % recipient)
            logger.error("%d:%s" % (resp.status_code, resp.text))
    # Create new access token to be used by the recipient
    else:
        access_token, expiry = generate_access_token(recipient)
        if access_token is not None and expiry is not None:
            access_tokens[recipient]['access_token'] = access_token
            access_tokens[recipient]['expiry'] = expiry
            pagination_url = get_mail_ids(recipient, subject, start_date, end_date, sender, pagination_url, only_has_attachments)
    return pagination_url


# Removes duplicate entries from among the mail recipients
def remove_duplicate_email_entries(recipients):
    logger.info("Removing duplicate entries from the recipient list")
    return list(set(recipients))


# Get actual list of recipients
def get_users(recipients, subject):
    access_token, expiry = directory_api.generate_directory_api_access_token(google_user_for_service_account)
    query_start_time = time.time()
    # Temp code to remove microsoft domain on searched emails
    recipients_copy = recipients.copy()
    for recipient in recipients:
        if 'onmicrosoft' in recipient:
            recipients_copy.append(recipient.replace('onmicrosoft.', ''))
    recipients = recipients_copy
    # Make the API call if token expiry time is greater than 1 minute
    if (expiry - query_start_time) > 60:
        # Verify and generate recipient list including resolving dls. It has to be a list even if its a single recipient
        if len(recipients) == 1 and not recipients[0]:  # If recipients input is blank, get all users in Company that have a mailbox
            logger.info("Generating list of all active users")
            recipients = directory_api.list_all_active_users(access_token)
        else:
            new_recipient_list = []
            for recipient in recipients:
                recipients_from_check = directory_api.recipient_exits_check(recipient, access_token)
                if recipients_from_check:
                    new_recipient_list.extend(recipients_from_check)
            if new_recipient_list:  # Overwrite recipients variable from user input with the verified recipients
                recipients = new_recipient_list
            else:
                recipients = []

        if recipients:
            # Remove duplicate entries of recipients from user input
            recipients = remove_duplicate_email_entries(recipients)
            recipients.sort()
            logger.info("Total number of recipients entered: {0:d}".format(len(recipients)))
        else:
            logger.info("No recipients received the mail with subject {0}".format(subject))
    else:
        logger.warning('Unable to verify recipients as the access token for Directory API was not created')
        recipients = []  # Send empty recipient list to kill Gmail

    return recipients


# Generate token and check if the user exists
def user_token(recipients, write_permissions=False):
    # Generate Access token for each recipient
    for recipient in recipients:
        if write_permissions:
            access_token, expiry = generate_access_token(recipient, need_write_access=True)
        else:
            access_token, expiry = generate_access_token(recipient)
        if access_token is not None and expiry is not None:
            access_tokens[recipient] = {}
            access_tokens[recipient]['access_token'] = access_token
            access_tokens[recipient]['expiry'] = expiry


# Fetch all mails that match the criteria
def get_emails(recipients, subject, start_date, end_date, sender, get_filtered_mails=False):
    global gmail_filtered_deleted_emails
    global gmail_filtered_emails

    # Fetch mail id for each recipient with the given set of conditions concurrently
    pagination_urls = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        fs = [executor.submit(get_mail_ids, recipient, subject, start_date, end_date, sender) for recipient in
              recipients]
        block_of_futures = []
        if len(fs) > 15:
            block_of_futures = [fs[i:i+15] for i in range(0, len(fs), 15)]
        else:
            block_of_futures.append(fs)
        for futures in block_of_futures:
            if futures:
                for future in concurrent.futures.as_completed(futures):
                    if future.result():
                        pagination_urls.append(future.result())

    # If pagination urls are returned by the above execution, run them to fetch more mail ids until they stop returning pagination urls
    while pagination_urls:
        paginations = []
        with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            fs = [executor.submit(get_mail_ids, pagination_url=pagination_url) for pagination_url in pagination_urls]
            block_of_futures = []
            if len(fs) > 15:
                block_of_futures = [fs[i:i+15] for i in range(0, len(fs), 15)]
            else:
                block_of_futures.append(fs)
            for futures in block_of_futures:
                if futures:
                    for future in concurrent.futures.as_completed(futures):
                        if future.result():
                            paginations.append(future.result())
            pagination_urls = paginations

    # If no mails were fetched, exit
    if not gmail_emails:
        logger.info("Email not found in the recipients mailboxes")
        return

    # Get mail information including the metadata from their mail id fetched from the previous execution
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        for mail in gmail_emails:
            executor.submit(populate_emails, mail)

    global matching_subject_mails
    matching_subject_mails = [email for email in gmail_emails if email.subject == subject]

    if get_filtered_mails:
        # Remove emails from the list that are not deleted or are in recipient Trash; and push them to filtered_emails list
        logger.info("Filtering emails that are deleted.")
        gmail_filtered_deleted_emails = [mail for mail in matching_subject_mails if mail.in_deleteditems]
    else:
        # Remove emails from the list that are already deleted or are in recipient Trash and push them to filtered_emails list
        logger.info("Filtering emails that are not deleted.")
        gmail_filtered_emails = [mail for mail in matching_subject_mails if not mail.in_deleteditems]


# Restore mail from trash
def restore_mails(sender, recipients, subject, start_date, end_date):
    # Purge global lists so that re-using the script doesn't cause conflict and display weird behavior
    global gmail_emails
    gmail_emails = []
    global gmail_filtered_deleted_emails
    gmail_filtered_deleted_emails = []

    # Format user input into lists and add time info to start and end date if necessary
    recipients, start_date, end_date = format_user_input(recipients, start_date, end_date)

    # Get actual list of recipients
    recipients = get_users(recipients, subject)

    if recipients:
        # Generate ouath token and check existence of users
        user_token(recipients, write_permissions=True)

        # Fetch all mails that match the criteria
        get_emails(recipients, subject, start_date, end_date, sender, get_filtered_mails=True)

        # Print out all mails that were not deleted and push them to recipient's Trash
        if gmail_filtered_deleted_emails:
            gmail_emails = []
            initial_filtered_email_recipients = []
            recipients_with_restored_mail = []
            try:
                for mail in gmail_filtered_deleted_emails:
                    initial_filtered_email_recipients.append(mail.requested_recipient)

                    if mail.undelete_mail():
                        recipients_with_restored_mail.append(mail.requested_recipient)

                logger.info("Email Restore successful. Mail is present in {0} Deleted Folder".format(str(recipients_with_restored_mail).strip('[').strip(']')))
                success_failure_message = "Email Restore successful for subject '{0}'. Mail is present in {1} Deleted Folder".format(subject, str(recipients_with_restored_mail).strip('[').strip(']'))

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

                #return success_failure_message
            except Exception as e:
                logger.error(e)
                logger.critical("Ran into error. Run restore script again for recipient {}".format(str(recipients).strip('[').strip(']')))
        else:
            # All the recipients have already deleted the mail.
            recipients = []  # The recipient list is not required anymore, and hence is being overwritten to get the list of all recipients whose mail was deleted
            for email in gmail_emails:
                recipients.append(email.requested_recipient)
            recipients = list(set(recipients))
            logger.info("Recipients {0} have already restored the mail.".format(str(recipients).strip('[').strip(']')))
            success_failure_message = "Recipients {0} have already restored the mail for subject '{1}'.".format(str(recipients).strip('[').strip(']'), subject)
            #return success_failure_message


# Delete mail to trash
def remove_mails(sender, recipients, subject, start_date, end_date):
    # Purge global lists so that re-using the script doesn't cause conflict and display weird behavior
    global gmail_emails
    gmail_emails = []
    global gmail_filtered_emails
    gmail_filtered_emails = []

    # Format user input into lists and add time info to start and end date if necessary
    recipients, start_date, end_date = format_user_input(recipients, start_date, end_date)

    # Get actual list of recipients
    recipients = get_users(recipients, subject)

    if recipients:
        # Generate ouath token and check existence of users
        user_token(recipients, write_permissions=True)

        # Fetch all mails that match the criteria
        get_emails(recipients, subject, start_date, end_date, sender)

        # Print out all mails that were not deleted and push them to recipient's Trash
        if gmail_filtered_emails:
            gmail_emails = []
            initial_filtered_email_recipients = []
            recipients_with_deleted_mail = []
            try:
                for mail in gmail_filtered_emails:
                    initial_filtered_email_recipients.append(mail.requested_recipient)

                    if mail.delete_mail():
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

                #return success_failure_message
            except Exception as e:
                logger.error(e)
                logger.critical("Ran into error. Run pull script again for recipient {}".format(str(recipients).strip('[').strip(']')))
        else:
            # All the recipients have already deleted the mail.
            recipients = []  # The recipient list is not required anymore, and hence is being overwritten to get the list of all recipients whose mail was deleted
            for email in gmail_emails:
                recipients.append(email.requested_recipient)
            recipients = list(set(recipients))
            logger.info("Recipients {0} have already deleted the mail.".format(str(recipients).strip('[').strip(']')))
            success_failure_message = "Recipients {0} have already deleted the mail for subject '{1}'.".format(str(recipients).strip('[').strip(']'), subject)
            #return success_failure_message
