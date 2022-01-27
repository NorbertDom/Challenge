from email.mime.text import MIMEText
from playwright.sync_api import sync_playwright
import requests
from os.path import basename, abspath, dirname
import pandas
import configparser
import smtplib
import datetime
import logging
import base64

logging.basicConfig(filename="report.log",
                    format='%(asctime)s %(message)s',
                    filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.INFO)


class ChallengeWebsite:
    def __init__(self, page):
        self.page = page
        page.click("text=\"Start\"")

    def challenge_website_submit(self, data_array):
        self.page.fill("input[ng-reflect-name=\"labelFirstName\"]", data_array[0])
        self.page.fill("input[ng-reflect-name=\"labelLastName\"]", data_array[1])
        self.page.fill("input[ng-reflect-name=\"labelCompanyName\"]", data_array[2])
        self.page.fill("input[ng-reflect-name=\"labelRole\"]", data_array[3])
        self.page.fill("input[ng-reflect-name=\"labelAddress\"]", data_array[4])
        self.page.fill("input[ng-reflect-name=\"labelEmail\"]", data_array[5])
        self.page.fill("input[ng-reflect-name=\"labelPhone\"]", data_array[6])
        self.page.click("text=\"Submit\"")


class RoboFormWebsite:
    def __init__(self, page):
        self.page = page

    def robo_form_website_submit(self, data_array):
        self.page.fill("input[name=\"02frstname\"]", data_array[0])
        self.page.fill("input[name=\"04lastname\"]", data_array[1])
        self.page.locator("input[value='Reset']").click()


def get_input_file(path_to_online_file):
    req = requests.get(path_to_online_file, allow_redirects=True)
    file_name = path_to_online_file.rsplit('/', 1)[1]
    open(file_name, 'wb').write(req.content)
    source_file_name = ("%s\\%s" % ((dirname(abspath(__file__))), (basename(req.url))))
    source_data = pandas.read_excel(source_file_name)
    return source_data


def row_to_array(row):
    output = []
    output.insert(0, row['First Name'])
    output.insert(1, row['Last Name '])
    output.insert(2, row['Company Name'])
    output.insert(3, row['Role in Company'])
    output.insert(4, row['Address'])
    output.insert(5, row['Email'])
    output.insert(6, str(row['Phone Number']))
    return output


class MailNotification:
    def __init__(self, smtp_server, smtp_port):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port

    def send_confirmation(self, user, b64code, receiver, subject, body):

        msg = MIMEText(body)
        msg['From'] = user
        msg['To'] = receiver
        msg['Subject'] = subject

        try:
            server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            server.ehlo()
            server.login(user, b64code)
            server.sendmail(user, receiver, msg.as_string())
            server.close()

        except:
            logger.error("Problem z wysłaniem maila. Sprawdź plik konfiguracyjny - conf.ini")


def main():
    config = configparser.ConfigParser()
    config.read('conf.ini')
    try:
        path_website_rpa = config.get('website_conf', 'path_website_rpa')
        path_website_robo = config.get('website_conf', 'path_website_robo')
        path_to_csv_source = config.get('website_conf', 'path_to_csv_source')
        user = config.get('mail_conf', 'gm_user')
        cred = (base64.b64decode(config.get('mail_conf', 'gm_cred'))).decode("utf-8")
        receiver = config.get('mail_conf', 'gm_receiver')
        body_file = config.get('mail_conf', 'gm_body_path')
        subject = "RPA Summary %s" % (datetime.datetime.now())
        smtp_server = config.get('mail_conf', 'gm_smtp_server')
        smtp_port = config.get('mail_conf', 'gm_smtp_port')
    except:
        logger.error('Problem z plikami konfiguracyjnymi. Sprawdź poprawność conf.ini.')

    source_data = get_input_file(path_to_csv_source)

    mail = MailNotification(smtp_server, smtp_port)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page_challenge = context.new_page()
        page_robo_form = context.new_page()
        page_challenge.goto(path_website_rpa)
        page_robo_form.goto(path_website_robo)

        challenge = ChallengeWebsite(page_challenge)

        robo = RoboFormWebsite(page_robo_form)

        for index, row in source_data.iterrows():
            try:
                challenge.challenge_website_submit(row_to_array(row))
                robo.robo_form_website_submit(row_to_array(row))
                logger.info("Obsługa sprawy: %s - Klient: %s" % (index + 1, row_to_array(row)[0]))
            except:
                logger.info("Błąd: %s - Klient: %s" % (index + 1, row_to_array(row)[0]))

        page_robo_form.close()
        page_challenge.close()
        context.close()
        browser.close()
        logger.info("Proces zakończony")

        body = open(body_file)

        msg = body.read()
        mail.send_confirmation(user, cred, receiver, subject, msg)
        body.close()


main()
