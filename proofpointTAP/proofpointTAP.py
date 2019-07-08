import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import datetime
import logging
import ConfigParser
import os
from base64 import b64encode

logger = logging.getLogger(__name__)

Config = ConfigParser.ConfigParser()
Config.read(os.path.join(os.path.abspath(os.path.dirname(__file__)),'Proofpoint_creds'))
principal = Config.get('Settings', 'Proofpoint_Service_Principal')
proofpoint_secret = Config.get('Settings', 'Proofpoint_Secret')


# Generate session with max of 3 retries and interval of 1 second
def session_generator():
    session = requests.Session()
    retry = Retry(connect=3, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    return session


# A class to fetch email details with respect to threats
class ProofpointEmailDetail:
    def __init__(self):
        self.campaign_name = None
        self.threat_id = None  # used to call Forensic API to IOCS
        self.malicious = False
        self.recipient = []
        self.subject = None
        self.sender = None
        self.sender_IP = None
        self.hash_of_attachment = None
        self.malicious_url = None
        self.attachments = {}
        self.false_positive = False
        self.has_attachments = False

    # Fetches the campaign name associated with a campaign ID
    def get_campaign_name(self, campaignID):
        headers = {'Authorization': 'Basic %s' % b64encode(("%s:%s" % (principal, proofpoint_secret)).encode()).decode(),
                   'Content-Type': 'application/json'}
        session = session_generator()
        resp = session.get("https://tap-api-v2.proofpoint.com/v2/campaign/%s" % campaignID, headers=headers)
        if resp.status_code == 200:
            self.campaign_name = resp.json()['name']
        else:
            logger.warning('Unable to connect to the campaign API.')

    # Fetch campaign name associated with a message event
    def get_campaign_name_from_message(self, events):
        if events:
            for threat in events:
                # Filter based on not being a false positive
                if threat['threatStatus'] != 'falsepositive':
                    # Use the campaign ID to get campaign name
                    if threat['campaignID'] is not None:
                        self.get_campaign_name(threat['campaignID'])
                        self.malicious = True
                    else:
                        self.campaign_name = threat['threatID']

                    # Fetch the threat ID to be used in Forensics to get IOCS
                    if threat['threatID']:
                        self.threat_id = threat['threatID']
                        self.malicious = self.is_malicious()

                    if threat['threatType'].upper() == 'ATTACHMENT':
                        self.has_attachments = True

                    # Get hash of attached malicious document
                    if self.malicious and self.has_attachments:
                        self.hash_of_attachment = threat['threat']

                    if self.malicious and threat['threatType'].upper() == 'URL':
                        self.malicious_url = threat['threat']
                elif threat['threatStatus'] == 'falsepositive':
                    self.false_positive = True

    # Fetch campaign associated with a click event
    def get_campaign_name_from_clicks(self, event):
        if event['campaignID'] is not None:
            self.get_campaign_name(event['campaignID'])
        else:
            self.campaign_name = event['threatID']
        self.threat_id = event['threatID']
        self.malicious = self.is_malicious()
        if self.malicious:
            self.malicious_url = event['url']

    # Use the Forensic API to confirm maliciousness of threat
    def is_malicious(self):
        headers = {'Authorization': 'Basic %s' % b64encode(("%s:%s" % (principal, proofpoint_secret)).encode()).decode(),
                   'Content-Type': 'application/json'}
        session = session_generator()
        resp = session.get("https://tap-api-v2.proofpoint.com/v2/forensics?threatId=%s" % self.threat_id,
                            headers=headers)
        if resp.status_code == 200:
            response = resp.json()
            for report in response['reports']:
                for event in report['forensics']:
                    # If event was malicious in the sandbox, return true
                    if event['malicious']:
                        return True
        else:
            logger.warning('Unable to connect to the forensic API.')


# Fetches emails from Proofpoint TAP from a certain time today
def get_emails(interval):
    logger.info('Requesting events from Proofpoint.')
    from_time = str(
        (datetime.datetime.utcnow() - datetime.timedelta(minutes=interval)).isoformat(sep='T', timespec='seconds'))
    headers = {'Authorization': 'Basic %s' % b64encode(("%s:%s" % (principal, proofpoint_secret)).encode()).decode(),
               'Content-Type': 'application/json'}
    session = session_generator()
    resp = session.get("https://tap-api-v2.proofpoint.com/v2/siem/all?format=json&sinceTime=%sZ" % from_time, headers=headers)

    if resp.status_code == 200:
        events = resp.json()
        return events
    elif resp.status_code == 429:
        logger.critical('Throttle limit reached. Wait for 24 hrs.')
        exit(1)
    elif resp.status_code == 401:
        logger.critical('Somebody removed your credentials or TAP is down.')
        return None
    else:
        logger.critical('Proofpoint error %d:%s', resp.status_code, resp.text)
        return None
