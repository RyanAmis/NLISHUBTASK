from datetime import datetime, timedelta
import os
import shutil
import sys
import zipfile
import smtplib


import pysftp


def send_task_completed_email(bacs_file_num, ndf_file_num, email_num):
    subject = 'NLIS Hub Bacs Transfer Task'
    body = f'''The NLIS Hub Bacs Transfer Task has been successfully completed for today.
    BACs files moved: {bacs_file_num}
    NDF files moved: {ndf_file_num}
    Emails moved: {email_num}
    '''
    message = f'Subject: {subject}\n\n{body}'

    with smtplib.SMTP(SERVER) as server:
        server.sendmail(from_addr=FROM, to_addrs=RECIPIENTS, msg=message)


def send_task_failed_email(error_message):
    subject = 'FAILED - NLIS Hub Bacs Transfer Task'
    body = f'The NLIS Hub Bacs Transfer Task has failed for today.\n\nError: {error_message}'
    message = f'''Subject: {subject}\n\n{body}
    '''

    with smtplib.SMTP(SERVER) as server:
        server.sendmail(from_addr=FROM, to_addrs=RECIPIENTS, msg=message)


SFTP_SERVER_IP = os.environ.get('SFTP_SERVER_IP')
SFTP_SERVER_USERNAME = os.environ.get('SFTP_SERVER_USERNAME')
SFTP_SERVER_PASSWORD = os.environ.get('SFTP_SERVER_PASSWORD')
TODAY_DATE = datetime.today().strftime('%d%m%Y')
TODAY_DATE_REVERSE = datetime.today().strftime('%Y%m%d')
YESTERDAY_DATE_REVERSE = (datetime.today() - timedelta(1)).strftime('%Y%m%d')
TODAY = datetime.today().strftime('%A')
MONTH = datetime.today().strftime('%B')
YEAR = datetime.today().strftime('%Y')
SATURDAY_DATE_REVERSE = (datetime.today() - timedelta(2)).strftime('%Y%m%d')
FRIDAY_DATE_REVERSE = (datetime.today() - timedelta(3)).strftime('%Y%m%d')
CATCH_DAY = 'Monday'

HUBACC_FOLDER_PATH = '/data/nlis_data/hub_proddb/settlement/export/hubacc/'
EMAIL_FOLDER_PATH = '/data/nlis_data/hub_proddb/settlement/email/'
CSERIES_FOLDER_PATH = 'J:/'
SI_FOLDER_PATH = 'Y:/'
CORP_HUB_STATEMENTS_PATH = 'H:/'
MDL_HUB_MAIL_PATH = 'K:/MDL Hub/MDH_Users/NLIS Hub/Operations/Settlement/Mail/'
HUB_BACS_FOLDER_PATH = "K:/IT DEPARTMENT/Hub BACS/"
DATE_FOLDER_PATH = HUB_BACS_FOLDER_PATH + TODAY_DATE + '/'
ZIPPED_FOLDER_PATH = DATE_FOLDER_PATH + 'Zipped folders' + '/'
MDL_HUB_MAIL_YEAR_MONTH_FOLDER_PATH = MDL_HUB_MAIL_PATH + YEAR + '/' + MONTH + '/'
CORP_HUB_STATEMENTS_YEAR_MONTH_FOLDER_PATH = CORP_HUB_STATEMENTS_PATH + YEAR + '/' + MONTH + '/'


SERVER = 'smtp.landmark.co.uk'
FROM = 'NLISHubBacsTransferTask@landmark.co.uk'
RECIPIENTS = [
    'jacob.pagan@landmark.co.uk', 'wayne.algar@landmark.co.uk', 'ryan.amis@landmark.co.uk',
    'matthew.hobson@landmark.co.uk', 'matthew.ellis@landmark.co.uk'
]

if not os.path.exists(DATE_FOLDER_PATH):
    os.makedirs(DATE_FOLDER_PATH)
    os.makedirs(ZIPPED_FOLDER_PATH)
else:
    sys.exit(1)

if not os.path.exists(MDL_HUB_MAIL_YEAR_MONTH_FOLDER_PATH):
    os.makedirs(MDL_HUB_MAIL_YEAR_MONTH_FOLDER_PATH)

if not os.path.exists(CORP_HUB_STATEMENTS_YEAR_MONTH_FOLDER_PATH):
    os.makedirs(CORP_HUB_STATEMENTS_YEAR_MONTH_FOLDER_PATH)

try:
    cnopts = pysftp.CnOpts(
        knownhosts='K:/IT DEPARTMENT/Nathan\'s Folder/Automation/NLIS Hub Bacs Transfer Task/Resources/known_hosts'
    )
    sftp = pysftp.Connection(
        host=SFTP_SERVER_IP, username=SFTP_SERVER_USERNAME, password=SFTP_SERVER_PASSWORD, cnopts=cnopts
    )
except Exception as err:
    send_task_failed_email(err)
    sys.exit(1)

hubacc_folder_contents = sftp.listdir(HUBACC_FOLDER_PATH)
for file in hubacc_folder_contents:
    if 'BACS_' in file and TODAY_DATE_REVERSE in file:
        sftp.get(HUBACC_FOLDER_PATH + file, DATE_FOLDER_PATH + file)
    elif 'CASH_SUM_SUN' in file and TODAY_DATE_REVERSE in file:
        sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
    elif 'CASH_SUN' in file and TODAY_DATE_REVERSE in file:
        sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
    elif 'REV_SUM_SUN' in file:
        if TODAY != CATCH_DAY and YESTERDAY_DATE_REVERSE in file:
            sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
        elif TODAY == CATCH_DAY:
            if YESTERDAY_DATE_REVERSE in file:
                sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
            elif SATURDAY_DATE_REVERSE in file:
                sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
            elif FRIDAY_DATE_REVERSE in file:
                sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
    elif 'REV_SUN' in file:
        if TODAY != CATCH_DAY and YESTERDAY_DATE_REVERSE in file:
            sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
        elif TODAY == CATCH_DAY:
            if YESTERDAY_DATE_REVERSE in file:
                sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
            elif SATURDAY_DATE_REVERSE in file:
                sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)
            elif FRIDAY_DATE_REVERSE in file:
                sftp.get(HUBACC_FOLDER_PATH + file, ZIPPED_FOLDER_PATH + file)

zipped_folders_contents = os.listdir(ZIPPED_FOLDER_PATH)
os.chdir(ZIPPED_FOLDER_PATH)
for file in zipped_folders_contents:
    if zipfile.is_zipfile(file):
        with zipfile.ZipFile(file, 'r') as zipped_file:
            zipped_file.extractall(DATE_FOLDER_PATH)

bacs_num = 0
ndf_num = 0
date_folder_contents = os.listdir(DATE_FOLDER_PATH)
os.chdir(DATE_FOLDER_PATH)
for file in date_folder_contents:
    if 'BACS_' in file:
        shutil.copy(file, CSERIES_FOLDER_PATH)
        bacs_num += 1
    elif 'CASH_SUM_SUN' in file:
        os.rename(file, 'HC' + YESTERDAY_DATE_REVERSE + '.ndf')
        shutil.copy('HC' + YESTERDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
        ndf_num += 1
    elif 'CASH_SUN' in file:
        os.rename(file, 'HCC' + YESTERDAY_DATE_REVERSE + '.ndf')
        shutil.copy('HCC' + YESTERDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
        ndf_num += 1
    elif 'REV_SUM_SUN' in file:
        if TODAY != CATCH_DAY and YESTERDAY_DATE_REVERSE in file:
            os.rename(file, 'HR' + YESTERDAY_DATE_REVERSE + '.ndf')
            shutil.copy('HR' + YESTERDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
            ndf_num += 1
        elif TODAY == CATCH_DAY:
            if YESTERDAY_DATE_REVERSE in file:
                os.rename(file, 'HR' + YESTERDAY_DATE_REVERSE + '.ndf')
                shutil.copy('HR' + YESTERDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
                ndf_num += 1
            elif SATURDAY_DATE_REVERSE in file:
                os.rename(file, 'HR' + SATURDAY_DATE_REVERSE + '.ndf')
                shutil.copy('HR' + SATURDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
                ndf_num += 1
            elif FRIDAY_DATE_REVERSE in file:
                os.rename(file, 'HR' + FRIDAY_DATE_REVERSE + '.ndf')
                shutil.copy('HR' + FRIDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
                ndf_num += 1
    elif 'REV_SUN' in file:
        if TODAY != CATCH_DAY and YESTERDAY_DATE_REVERSE in file:
            os.rename(file, 'HRR' + YESTERDAY_DATE_REVERSE + '.ndf')
            shutil.copy('HRR' + YESTERDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
            ndf_num += 1
        elif TODAY == CATCH_DAY:
            if YESTERDAY_DATE_REVERSE in file:
                os.rename(file, 'HRR' + YESTERDAY_DATE_REVERSE + '.ndf')
                shutil.copy('HRR' + YESTERDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
                ndf_num += 1
            elif SATURDAY_DATE_REVERSE in file:
                os.rename(file, 'HRR' + SATURDAY_DATE_REVERSE + '.ndf')
                shutil.copy('HRR' + SATURDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
                ndf_num += 1
            elif FRIDAY_DATE_REVERSE in file:
                os.rename(file, 'HRR' + FRIDAY_DATE_REVERSE + '.ndf')
                shutil.copy('HRR' + FRIDAY_DATE_REVERSE + '.ndf', SI_FOLDER_PATH)
                ndf_num += 1

email_num = 0
email_folder_contents = sftp.listdir(EMAIL_FOLDER_PATH)
for file in email_folder_contents:
    if TODAY_DATE_REVERSE in file:
        sftp.get(EMAIL_FOLDER_PATH + file, MDL_HUB_MAIL_YEAR_MONTH_FOLDER_PATH + file)
        sftp.get(EMAIL_FOLDER_PATH + file, CORP_HUB_STATEMENTS_YEAR_MONTH_FOLDER_PATH + file)
        email_num += 1

sftp.close()
send_task_completed_email(bacs_file_num=bacs_num, ndf_file_num=ndf_num, email_num=email_num)
sys.exit(0)
