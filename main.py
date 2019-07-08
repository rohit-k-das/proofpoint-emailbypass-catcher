import proofpointTAP.proofpointTAP as proofpointTAP
import mailApp.MicrosoftOutlookMail as MicrosoftOutlookMail
import mailApp.Gmail as Gmail
import mailApp.Directory as directory_api
import logging
import datetime
import argparse
import requests
import os
import json
import urllib.parse
import ConfigParser

active_users = []
email_dls = {}
aliases = []

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(name)-15s [%(levelname)-8s]: %(message)s',
                    datefmt='%m/%d/%Y %I:%M:%S %p')
logger = logging.getLogger(__name__)

Config = ConfigParser.ConfigParser()
Config.read(os.path.join(os.path.abspath(os.path.dirname(__file__)), 'settings.ini'))
google_user_for_service_account = Config.get('Settings', 'Google_User_For_Project')

# Create your own slackbot
hubot_webhook_url = Config.get('Settings', 'Slackbot_Url')


# Current configuration checks for users in the company from Microsoft. You can switch to Google by uncommenting the lines below
def filter_recipient(recipients, access_token):
    # If the recipients variable is a list iterate through
    if isinstance(recipients, list):
        recipients_copy = []
        for recipient in recipients:
            recipients_from_check = MicrosoftOutlookMail.recipient_exits_check(recipient)
            if recipients_from_check:
                recipients_copy.extend(recipients_from_check)

            '''
            recipints_from_check = directory_api.recipient_exits_check(recipient, access_token)
            recipients.remove(recipient)
            if recipints_from_check:
                recipients.extend(recipints_from_check)
            '''
        recipients = recipients_copy

    else:
        recipients = MicrosoftOutlookMail.recipient_exits_check(recipients)

        '''
        recipients = directory_api.recipient_exits_check(recipient, access_token)
        '''
    return recipients


# Collect the extracted malicious emails as per Proofpoint that were allowed
def parse_emails_per_threat_association(events):
    threats = []  # A list to hold ProofpointEmailDetail objects

    if 'messagesDelivered' in events or 'clicksPermitted' in events:
        # Generate oauth tokens to be used by Outlook API or Google API
        MicrosoftOutlookMail.oauth_access_token, MicrosoftOutlookMail.expiry_time = MicrosoftOutlookMail.generate_access_token()

        if not MicrosoftOutlookMail.oauth_access_token and MicrosoftOutlookMail.expiry_time > datetime.datetime.now():
            logger.critical('Unable to generate access token. Exiting..')
            return

        access_token, expiry = directory_api.generate_directory_api_access_token(google_user_for_service_account)

        # Mails with attachments that were not blocked
        if 'messagesDelivered' in events:
            if events['messagesDelivered']:
                logger.info('Extracting messages delivered in Proofpoint.')
                for event in events['messagesDelivered']:
                    Pobj = proofpointTAP.ProofpointEmailDetail()
                    # Only parse if threatInfoMap has some data
                    if event['threatsInfoMap']:

                        # Get Campaign Name and Threat ID
                        Pobj.get_campaign_name_from_message(event['threatsInfoMap'])

                        # Get receiver of mail
                        recipient = filter_recipient(event['recipient'], access_token)
                        if recipient:
                            Pobj.recipient.extend(recipient)
                        if 'ccAddresses' in event:
                            recipients = filter_recipient(event['ccAddresses'], access_token)
                            if recipients:
                                Pobj.recipient.extend(recipients)

                        # Get email subject
                        Pobj.subject = event['subject']

                        # Get email sender
                        Pobj.sender = event['headerFrom'].split('<')
                        Pobj.sender = Pobj.sender[len(Pobj.sender) - 1].strip('>')
                        Pobj.sender_IP = event['senderIP']
                        if Pobj.has_attachments:
                            for attachment in event['messageParts']:
                                if attachment['disposition'] == 'attached':
                                    Pobj.attachments[attachment['filename']] = attachment['contentType']
                        threats.append(Pobj)

        # Mails that are mostly phishing links which were not blocked
        if 'clicksPermitted' in events:
            if events['clicksPermitted']:
                logger.info('Extracting clicks permitted in Proofpoint.')
                for event in events['clicksPermitted']:
                    Pobj = proofpointTAP.ProofpointEmailDetail()
                    # Only parse if mail is associated with a threatID
                    if 'threatID' in event:
                        Pobj.get_campaign_name_from_clicks(event)
                        recipient = filter_recipient(event['recipient'], access_token)
                        if recipient:
                            Pobj.recipient.append(recipient)
                        Pobj.sender = event['sender']
                        Pobj.sender_IP = event['senderIP']
                        threats.append(Pobj)
    return threats


# Send alert with msg to slack via hubot
'''
Use POST with Body
'''
def send_alert_via_hubot(campaign, campaign_fields, number_of_users):
    alerts = []
    email_pull_messages = []
    if campaign_fields['EmailPulls']:
        for message in campaign_fields['EmailPulls']:
            email_pull_messages.append(message)
            if 'Email Pull successful' in message:
                alerts.append(':green-alert: Email Pull Successful')
            if 'already deleted the mail' in message:
                alerts.append(':green-alert: Already Deleted')
            if 'Unable to delete mail' in message:
                alerts.append(":red-alert: Didn't Delete/Pull")
            if 'Ran into error.' in message:
                alerts.append(":red-alert: Script Failed")
            if 'Email not found' in message:
                alerts.append(":amber-alert: Not Found")

    if campaign_fields['Attachments']:
        alerts.append(":mail-attachment: Attachment Found")

    alerts = list(set(alerts))
    # Message to send in the alert
    # Filter charactes & < > from message as slack will not be able to handle threat
    message_to_send = ":malicious-email: Alert: <https://threatinsight.proofpoint.com/14f465d2-8daf-445c-52a9-fec245f2d609/threat/email/%s|%s> ---> %s\n" % (campaign_fields['ThreatID'], campaign, str(alerts).strip('[').strip(']').replace('\'', ''))
    if campaign_fields['Subject']:
        message_to_send = "%sSubject: %s\n" % (message_to_send, str(campaign_fields['Subject']).strip('[').strip(']').encode('utf-8').decode('utf-8').replace('\'', '').replace('&', '%26amp;').replace('<', '%26lt;').replace('>', '%26gt;'))
    message_to_send = "%sTotal Recipients: %d\n" % (message_to_send, number_of_users)
    message_to_send = "%sRecipients: %s\n" % (message_to_send, str(campaign_fields['Recipients']).strip('[').strip(']').replace('\'', ''))
    message_to_send = "%sSenders: %s\n" % (message_to_send, str(campaign_fields['Senders']).strip('{').strip('}').replace('\'', ''))
    if campaign_fields['Attachments']:
        message_to_send = "%sAttachments: %s\n" % (message_to_send, str(campaign_fields['Attachments']).strip('{').strip('}').replace("'", "").replace('\\\'', '\'').replace('\\\"', '\"').replace('&', '%26amp;').replace('<', '%26lt;').replace('>', '%26gt;'))
    if email_pull_messages:
        message_to_send = "\n%sEmail Pull Report:\n" % message_to_send
        for message in email_pull_messages:
            message = message.replace('&', '%26amp;').replace('<', '%26lt;').replace('>', '%26gt;').replace('", "', "\n").replace('\\\'', '\'').replace('\\\"', '\"')
            if message[0] == '"':
                message = message[1:]
            if message[len(message)-1] == '"':
                message = message[:len(message)-1]
            message_to_send = "%s%s\n" % (message_to_send, message)
    else:
        message_to_send = "%s%s\n" % (message_to_send, 'No email pull messages unfortunately. Something is wrong.')
    if campaign_fields['IOCs clicked or downloaded']:
        message_to_send = "\n%sIOCs Report:\n" % message_to_send
        for message in campaign_fields['IOCs clicked or downloaded']:
            if message:
                message_to_send = "%s%s\n" % (message_to_send, message.replace('&', '%26amp;').replace('<', '%26lt;').replace('>', '%26gt;'))

    # Whom to send the alert
    send_to = 'Your channel or username'
    data = {'message': message_to_send, 'users': send_to}
    data = urllib.parse.urlencode(data)

    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    resp = requests.post(hubot_webhook_url, headers=headers, data=data)
    if resp.ok:
        logger.info("Sent alert to user/channel %s" % send_to)
    else:
        logger.critical("Unable to connect to hubot.")
        logger.info("Hubot Error %d:%s" % (resp.status_code, resp.text))
        exit(-1)


# Parse command line arguments
def parse_options():
    parser = argparse.ArgumentParser(description='This is a proofpoint alerter script')
    parser.add_argument("-i", "--interval", action="store", default=15, type=int,
                        dest='interval', help="Interval in minutes to look back to fetch mails")
    parser.add_argument("-t", "--threshold", action="store", default=4, type=int,
                        dest='threshold', help="Threshold after which alert will be generated per campaign")
    arguments = parser.parse_args()

    # Check if arguments are given values from cli
    if not arguments.interval:
        logger.info("Usage of script:")
        parser.usage()
        logger.warning("Going with the default value for interval.")

    if not arguments.threshold:
        logger.info("Usage of script:")
        parser.usage()
        logger.warning("Going with the default value for threshold.")

    return arguments


# Check if alert is already sent
def check_event(campaign):
    present = False
    if os.path.isfile('./Logged_event'):
        with open('Logged_event') as f:
            '''
            if str(campaign) in events:
                present = True
            '''
            # print(events)
            # print(campaign)
            for event in f.read().splitlines():
                event = event.replace("'", "\"")
                event = json.loads(event)
                if event['Subject'] == campaign['Subject'] and event['Recipients'] == campaign['Recipients'] and sorted(event['Senders']) == sorted(campaign['Senders']):
                    logger.info("Campaign %s present" % str(event))
                    present = True
                    break
    else:
        logger.warning('Logged_event file not created')
    return present


def email_pull_action_based_on_return_message(sender, recipients, subject, start_date, message):
    if 'Email not found in the' in message and sender:
        logger.info('Retrying to pull email by using only the sender, recipient and start date and skipping subject %s for recipient %s' % (subject, recipients))
        message = MicrosoftOutlookMail.email_pull(sender, recipients, "", start_date, skip_recipient_check=True)
        Gmail.remove_mails(sender, recipients, "", start_date, end_date="")
    if 'Email not found in the' in message and sender:
        logger.info('Retrying to pull email by using only the subject %s, recipient and start date and skipping sender for recipient %s' % (subject, recipients))
        message = MicrosoftOutlookMail.email_pull("", recipients, subject, start_date, skip_recipient_check=True)
        Gmail.remove_mails("", recipients, subject, start_date, end_date="")
    if 'Email not found in the' in message and sender:
        if '?utf-8' not in subject:
            logger.info('Retrying to pull email by using only the subject %s, recipient %s since the last 14 days' % (subject, recipients))
            message = MicrosoftOutlookMail.email_pull(sender, recipients, subject, (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), skip_recipient_check=True)
            Gmail.remove_mails(sender, recipients, subject, (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), end_date="")
        else:
            logger.info('Retrying to pull email by using only recipient %s and sender %s since the last 14 days' % (recipients, sender))
            message = MicrosoftOutlookMail.email_pull(sender, recipients, "", (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), skip_recipient_check=True)
            Gmail.remove_mails(sender, recipients, "", (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), end_date="")

    return message


def email_restore_action_based_on_return_message(sender, recipients, subject, start_date, message):
    if 'Email not found in the' in message and sender:
        logger.info('Retrying to restore email by using only the sender, recipient and start date and skipping subject %s for recipient %s' % (subject, recipients))
        message = MicrosoftOutlookMail.email_restore(sender, recipients, "", start_date, skip_recipient_check=True)
        Gmail.restore_mails(sender, recipients, "", start_date, end_date="")
    if 'Email not found in the' in message and sender:
        logger.info('Retrying to restore email by using only the subject %s, recipient and start date and skipping sender for recipient %s' % (subject, recipients))
        message = MicrosoftOutlookMail.email_restore("", recipients, subject, start_date, skip_recipient_check=True)
        Gmail.restore_mails("", recipients, subject, start_date, end_date="")
    if 'Email not found in the' in message and sender:
        if '?utf-8' not in subject:
            logger.info('Retrying to restore email by using only the subject %s, recipient %s since the last 14 days' % (subject, recipients))
            message = MicrosoftOutlookMail.email_restore(sender, recipients, subject, (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), skip_recipient_check=True)
            Gmail.restore_mails(sender, recipients, subject, (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), end_date="")
        else:
            logger.info('Retrying to restore email by using only recipient %s and sender %s since the last 14 days' % (recipients, sender))
            message = MicrosoftOutlookMail.email_restore(sender, recipients, "", (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), skip_recipient_check=True)
            Gmail.restore_mails(sender, recipients, "", (datetime.date.today() - datetime.timedelta(days=14)).isoformat(), end_date="")

    return message


def main():
    # Get arguments parsed via command line
    arguments = parse_options()
    interval = arguments.interval
    threshold = arguments.threshold

    events = proofpointTAP.get_emails(interval)
    if events is None:
        logger.info("No mail received via Proofpoint in the past %d minutes" % interval)
        exit(0)

    # Run sql query to fill in active employees and distribution list

    threats = parse_emails_per_threat_association(events)

    if not threats:
        logger.info("No mails allowed by proofpoint in the past %d minutes" % interval)
        exit(0)

    malicious_threats = []

    # Filter all malicious mails
    for threat in threats:
        if threat.malicious and not threat.false_positive:
            malicious_threats.append(threat)

    false_positive_threats = []
    for threat in threats:
        if threat.false_positive:
            false_positive_threats.append(threat)

    if not malicious_threats:
        logger.info("No mails that were allowed had any malicious content in the past %d minutes" % interval)
        exit(0)

    if false_positive_threats:
        logger.info("False positives found. Restoring them from the mailbox Trash folder.")

    email_restore_messages = []
    for threat in false_positive_threats:
        if threat.recipients:
            start_date = datetime.date.today().isoformat()
            subject = threat.subject
            recipients = threat.recipient
            sender = threat.sender
            logger.info("Restoring mail %s for recipient %s from sender %s" % (subject, recipients, sender))
            message = MicrosoftOutlookMail.email_restore(sender, recipients, subject, start_date, skip_recipient_check=True)
            Gmail.restore_mails(sender, recipients, subject, start_date, end_date="")
            if message:
                if subject not in message:
                    message = message.replace("''", "'%s'" % subject)

                while True:
                    if 'Unable to restore mail for' in message:
                        divide_message = message.split('\nUnable to restore mail for ')
                        recipients = divide_message[1]
                        message = MicrosoftOutlookMail.email_restore(sender, recipients, subject, start_date, skip_recipient_check=True)
                        Gmail.restore_mails(sender, recipients, subject, start_date, end_date="")

                    message = email_restore_action_based_on_return_message(sender, recipients, subject, start_date, message)
                    if not 'Unable to restore mail for' in message:
                        break
                email_restore_messages.append(message)

    malicious_campaigns = {}
    # Get all threats under their respective campaigns
    for threat in malicious_threats:
        campaign_name = threat.campaign_name
        # If campaign already exists in the dict, just add current Subject, Recipients, Senders to the existing ones
        if threat.recipient:
            if campaign_name in malicious_campaigns:
                if threat.subject is not None:
                    if threat.subject not in malicious_campaigns[campaign_name]['Subject']:
                        malicious_campaigns[campaign_name]['Subject'].append(threat.subject)

                malicious_campaigns[campaign_name]['Recipients'].extend(threat.recipient)

                if threat.sender is not None:
                    if threat.sender not in malicious_campaigns[campaign_name]['Senders']:
                        malicious_campaigns[campaign_name]['Senders'][threat.sender] = threat.sender_IP

                if threat.attachments:
                    for attachment in threat.attachments:
                        if attachment not in malicious_campaigns[campaign_name]['Attachments']:
                            malicious_campaigns[campaign_name]['Attachments'][attachment] = threat.attachments[attachment]

                if threat.hash_of_attachment is not None:
                    if threat.hash_of_attachment not in malicious_campaigns[campaign_name]['IOCs']:
                        malicious_campaigns[campaign_name]['IOCs'].append(threat.hash_of_attachment)

                if threat.malicious_url is not None:
                    if threat.malicious_url not in malicious_campaigns[campaign_name]['IOCs']:
                        malicious_campaigns[campaign_name]['IOCs'].append(threat.malicious_url)

            # Else Create a new dict
            else:
                malicious_campaigns[campaign_name] = {}
                malicious_campaigns[campaign_name]['Subject'] = []
                malicious_campaigns[campaign_name]['Recipients'] = []
                malicious_campaigns[campaign_name]['Senders'] = {}
                malicious_campaigns[campaign_name]['Attachments'] = {}
                malicious_campaigns[campaign_name]['IOCs'] = []
                malicious_campaigns[campaign_name]['EmailPulls'] = []
                # malicious_campaigns[campaign_name]['ThreatID'] = 0

                if threat.subject is not None:
                    malicious_campaigns[campaign_name]['Subject'].append(threat.subject)
                if threat.sender is not None:
                    malicious_campaigns[campaign_name]['Senders'][threat.sender] = threat.sender_IP
                if threat.attachments:
                    malicious_campaigns[campaign_name]['Attachments'].update(threat.attachments)
                if threat.hash_of_attachment is not None:
                    malicious_campaigns[campaign_name]['IOCs'].append(threat.hash_of_attachment)
                if threat.malicious_url is not None:
                    malicious_campaigns[campaign_name]['IOCs'].append(threat.malicious_url)
                malicious_campaigns[campaign_name]['Recipients'].extend(threat.recipient)
                malicious_campaigns[campaign_name]['ThreatID'] = threat.threat_id

            # Pull mails from MicrosoftOutlookMail and Gmail
            if int((MicrosoftOutlookMail.expiry_time - datetime.datetime.now()).seconds) > 60:
                start_date = datetime.date.today().isoformat()
                subject = threat.subject
                recipients = threat.recipient
                sender = threat.sender
                logger.info("Pulling email %s for recipient %s from sender %s" % (subject, recipients, sender))
                message = MicrosoftOutlookMail.email_pull(sender, recipients, subject, start_date, skip_recipient_check=True)
                Gmail.remove_mails(sender, recipients, subject, start_date, end_date="")
                if message:
                    if subject not in message:
                        message = message.replace("''", "'%s'" % subject)

                    while True:
                        if 'Unable to delete mail for' in message:
                            divide_message = message.split('\nUnable to delete mail for ')
                            recipients = divide_message[1]
                            message = MicrosoftOutlookMail.email_pull(sender, recipients, subject, start_date, skip_recipient_check=True)
                            Gmail.remove_mails(sender, recipients, subject, start_date, end_date="")

                        message = email_pull_action_based_on_return_message(sender, recipients, subject, start_date, message)
                        if not 'Unable to delete mail for' in message:
                            break
                    malicious_campaigns[campaign_name]['EmailPulls'].append(message)

    number_of_campaigns = len(malicious_campaigns)

    # Send alerts based on number of mails received per campaign in the last interval
    for index,campaign in enumerate(malicious_campaigns, start=1):
        # number_of_users = 0
        # if there are recipients
        if malicious_campaigns[campaign]['Recipients']:
            # Remove duplicate entries
            malicious_campaigns[campaign]['Recipients'].sort()
            malicious_campaigns[campaign]['Recipients'] = list(set(malicious_campaigns[campaign]['Recipients']))

            number_of_users = len(malicious_campaigns[campaign]['Recipients'])
            malicious_campaigns[campaign]['Subject'].sort()
            malicious_campaigns[campaign]['EmailPulls'].sort()
            malicious_campaigns[campaign]['IOCs'].sort()
            malicious_campaigns[campaign]['IOCs clicked or downloaded'] = []

            if index == number_of_campaigns and email_restore_messages:
                malicious_campaigns[campaign]['EmailPulls'] = malicious_campaigns[campaign]['EmailPulls'].extend(email_restore_messages).sort()


            # Check if the event is already alerted
            # event_already_alerted = check_event(malicious_campaigns[campaign])
            # logger.info(event_already_alerted)
            # if (number_of_users >= threshold or shared_mailbox_present) and not event_already_alerted:
            if number_of_users >= threshold:
                try:
                    logger.info("Sending hubort alert")
                    send_alert_via_hubot(campaign, malicious_campaigns[campaign], number_of_users)
                    with open('Logged_event', 'a+') as f:
                        f.write('%s\n' % str(malicious_campaigns[campaign]))
                except Exception as e:
                    logger.error(e)
                    logger.error(campaign, malicious_campaigns[campaign])


if __name__ == '__main__':
    main()
