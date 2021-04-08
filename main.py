#in commandline call hickory schedule main.py --every=day@9:00am
#to kill bot: hickory kill emailbot.py
import smtplib
import pandas as pd
import gspread
import gspread_dataframe as gd


from email.message import EmailMessage


# use Pandas library, import a csv file of names and emails
# check if the email has been sent
# after it sends an email make a csv summary of when we sent the email or make it output to a google doc
# read in or output to the google spreadsheet
# "pandas dataframe to google spreadsheet"
# g spread
# df to gspread
# record
# checks email for responses and updates spreadsheet
# automated followup

gc = gspread.service_account(filename='robosub-automated-sponsorships-3df1d7362ac9.json')
sh = gc.open_by_key('1VHx5tWv1xxpUJna_QN5fmFDFtWselMU3lYpyRwv38Hw')
worksheet = sh.sheet1
res = worksheet.get_all_records()

df = pd.DataFrame(res)


def send_email(receiver, message):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login('robosub@ucsd.edu', 'W2e6r#$fZK')

    email = EmailMessage()
    email['From'] = 'robosub@ucsd.edu'
    email['To'] = receiver
    email['Subject'] = 'UCSD Triton Robosub Sponsorship Inquiry'
    email.set_content(message)

    with open('Triton Robosub Sponsorship Packet 2021.pdf', 'rb') as f:
        file_data = f.read()
        file_name = f.name

    email.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
    server.send_message(email)


for index, row in df.iterrows():

    name = row['Company']
    body = 'To Whom It May Concern,\n' \
           '\n' \
           'My name is Jeffrey Chen, and I\'m the Corporate Liaison for Triton Robosub at UC San Diego. We are an engineering org focused on building, programming, and testing autonomous underwater vehicles for research and competition applications. We are entering our second annual Robosub competition, an international AUV competition hosted in San Diego. Going into our build season, I’m reaching out to inquire about a potential sponsorship opportunity.\n' \
           '\n' \
           'We are currently focused on rebuilding our robot in order to prepare for this year’s Robosub competition, as well as making it modular to be applicable to research projects. Currently, we are working on FishSense, a method of monitoring fish populations using depth imaging technology, as well as SeaThru, software capable of implementing real-time color correction of underwater images. ' + name + ' can help us take the next step in developing these projects with financial support, mentorship, and career opportunities.\n' \
                                                                                                                                                                                                                                                                                                                                                                                                                           '\n' \
                                                                                                                                                                                                                                                                                                                                                                                                                           'Triton Robosub is a registered non-profit through UCSD, so all support is tax-deductible. In return for sponsorship, we can connect you to the talented engineers on our team through workshops, networking sessions and more, and include your logo on our official team gear and documents! More details can be found in our 2021 sponsorship packet, which is attached. We would really love to work with you, and with your help push the field of underwater robotics to its limits!\n' \
                                                                                                                                                                                                                                                                                                                                                                                                                           '\n' \
                                                                                                                                                                                                                                                                                                                                                                                                                           'Hope to hear from you soon,\n' \
                                                                                                                                                                                                                                                                                                                                                                                                                           '\n' \
                                                                                                                                                                                                                                                                                                                                                                                                                           'Jeffrey Chen\n' \
                                                                                                                                                                                                                                                                                                                                                                                                                           'UCSD Triton Robosub | Corporate Liaison'

    if row['Sent?'] != 'Yes':
        send_email(row['Email'], body)
        row['Sent?'] = 'Yes'

gd.set_with_dataframe(worksheet, df)

