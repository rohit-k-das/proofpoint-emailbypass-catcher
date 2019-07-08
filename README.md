Proofpoint Email Bypass Catcher
-------------------------------

When Proofpoint takes more time to analyse the mail, other similar mails are usually allowed through until a conclusion is reached.

The script uses Proofpoint SIEM API to find malicious mails that slipped through and pull them from the user's mailbox. An alert is sent to your slack channel via slackbot.

Fill in the creds files and settings.ini files.

For details on how to what to fill for the Mail Creds see https://github.com/rohit-k-das/pymail

Install requirements: `pip install -r requirements.lock`

Run the script using python 3:
    ``python3 -B main.py -i 14 -t 1`` where 
    i = interval to scan from the proofpoint API
    t = threshold of number of malicious mails after which slack alert would be sent
    
The above script has the provision to use Microsoft & Gmail API. Use according to what you use in your company.
    
    
