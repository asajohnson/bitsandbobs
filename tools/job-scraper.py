# Lambda function code to automate looking for posted jobs with given keywords.

import requests
from bs4 import BeautifulSoup
import smtplib
from email.message import EmailMessage
import boto3
import json

# Keywords and URLs to scan
URLS = ['**Sites**']
KEYWORDS = ['**keywords**']

def get_email_secrets():
    secret_name = "**SecretName**"
    region_name = "**awsregion**"

    client = boto3.client('secretsmanager', region_name=region_name)
    get_secret_value_response = client.get_secret_value(SecretId=secret_name)
    secret = json.loads(get_secret_value_response['SecretString'])

    return secret['EMAIL_ADDRESS'], secret['EMAIL_PASSWORD'], secret['SEND_TO']

EMAIL_ADDRESS, EMAIL_PASSWORD, SEND_TO = get_email_secrets()

def find_keywords_in_urls(urls, keywords):
    matches = []
    for url in urls:
        try:
            response = requests.get(url, timeout=10) # 10 sec per request
            soup = BeautifulSoup(response.text, 'html.parser')
            text = soup.get_text().lower()
            if any(keyword.lower() in text for keyword in keywords):
                matches.append(url)
        except Exception as e:
            print(f'Error scraping {url}: {e}')
    return matches

def send_email(matched_urls):
    msg = EmailMessage()
    msg['Subject'] = 'Keyword Match Alert'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = SEND_TO
    msg.set_content('Keywords were found on the following pages:\n\n' + '\n'.join(matched_urls))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

def lambda_handler(event, context):
    matches = find_keywords_in_urls(URLS, KEYWORDS)
    if matches:
        send_email(matches)
    return {
        'statusCode': 200,
        'body': f'Matches found: {len(matches)}'
    }
