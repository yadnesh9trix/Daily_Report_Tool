import os
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import numpy as np
import pandas as pd
import csv
from tabulate import tabulate

today = datetime.datetime.today()

def send():
    # creates SMTP session
    # text, pickup_df, segment_pickup_df_updated = summary(hf_df, ms_df, date)
    eid_data = pd.read_excel("D:\Daily_Report_Tool\Mapping\mail_mapping/send_email_TESt.xlsx")
    eid_h_to = eid_data['emailid'][eid_data['type'] == 1]
    eid_h_cc = eid_data['emailid'][eid_data['type'] == 2]

    msg = MIMEMultipart()

    text_body = "Dear Team," \
                "\r\n\r\n" \
                "Greetings of the Day!" \
                "\r\n\r\n" \
                "Please find the attached {} daily collection report.".format(today.date())

  #   html = """\
  #   <html>
  #   <head>
  #   <style>
  #   th {{height: 40px; background-color: #606d99 !important; font-weight: bold;}}
  #   td {{vertical-align: center;}}
  #   table {{width: 85%; border-collapse: collapse; height: 120px; font-size: 20px;margin-left: auto;
  # margin-right: auto;}}
  #   table, th, td {{border:1px solid black; text-align: center  !important;}}
  #   </style>
  #   </head>
  #   <body style="padding-left: 80px;
  #               padding-right: 80px;
  #               padding-top: 50px;
  #               padding-bottom: 50px;
  #               font-size: 20px;
  #               text-align: center !important;">
  #   <p>
  #   <br>
  #   {0}
  #   <br>
  #   <br>
  #   <br>
  #   <br>
  #   {1}
  #   <br>
  #   {2}
  #   <br>
  #   {3}
  #   <br>
  #   {4}
  #   <br>
  #   {5}
  #   <br>
  #   <br><br><br><br><br>
  #   {6}
  #   <br>
  #   </p>
  #   </body>
  #   </html>
  #   """.format(text_body, text, month_df.to_html(index=False),month_df_sum_stly.to_html(index=False),pickup_df.to_html(index=False), segment_pickup_df.to_html(index=False),text_note)

    #OLD
    # html = """
    # <html>
    # <head>
    # <style>
    #  table, th, td {{ border: 1px solid black; border-collapse: separate; border-radius: 10px}}
    #   th, td {{ padding: 15px; }}
    # </style>
    # </head>
    # <body>
    # <br>
    # <br>
    # <p>Dear Team,</p>
    # <p>Greetings of the Day!</p>
    # <p>Please find the attached Todays collection report.</p>
    # <p>Summary of today's zone wise collection:</p>
    # <p></p>
    # {table}
    # <p></p>
    # <p></p>
    # </body></html>
    # """

    html = """
    <html>
    <head>
    <style>
     table, th, td {{ border: 1px solid black; border-collapse: separate; border-radius: 10px}}
      th, td {{ padding: 15px; }}
    </style>
    </head>
    <body>
    <p>सर्वांना नमस्कार,</p>
    <p>दिवसभराचा गटनिहाय वसूली तक्ता आणि विभागीय कार्यालयनिहाय मागणी तक्ता सोबत जोडलेला आहे तो पहा.</p>
    <br>
    <br>
    <p>विभागनिहाय संकलनचा सारांश:</p>
    <p>रक्कम रुपये कोटीमध्ये</p>
    <p></p>
    {table}
    <p></p>
    <p></p>
    </body></html>
    """

    # टीममधील सर्वांना नमस्कार,
    # दिवसभराचा संकलन अहवाल सोबत जोडलेला आहे तो पहा.सोबत विभागनिहाय संकलनचा सारांश तक्ता.

    # msg.attach(MIMEText(text_body, 'plain'))
    # msg.attach(MIMEText(html, 'html'))

    # with open(mailreport+f"{today.date()}_collectiondata.csv") as input_file:
    #     reader = csv.reader(input_file)
    #     data = list(reader)
    std_path = r"C:\PTAX Project\Daily_Report_Tool/"
    in_path = std_path + "Input/" + str(today) + "/"
    outpth = std_path + "Output/" + str(today) + "/"
    # mappath = std_path + "Mapping/"
    # logopath = std_path + "logo/"
    mailreport = std_path + "Mail_report/"
    # mailreport = std_path + "Mail_report/"

    try:
        data = pd.read_csv(mailreport+f"{today.date()}_collectiondata.csv",encoding='utf-8-sig')
    except:
        data =pd.DataFrame()

    data = data.replace(np.nan,"")
    text = text_body.format(table=tabulate(data, headers="firstrow", tablefmt="grid"))
    # html = html.format(table=tabulate(data,headers=['अ.क्र.', 'विभागीय कार्यालय', 'वसूली'], tablefmt="html",showindex=False))
    html = html.format(table=tabulate(data, tablefmt="html",showindex=False))

    msg = MIMEMultipart(
        "alternative", None, [MIMEText(text), MIMEText(html, 'html')])
    # msg = MIMEMultipart(
    #     "alternative", None, [MIMEText(text,'plain'), MIMEText(html, 'html')])
    # msg = MIMEMultipart(
    #     "alternative", None, [MIMEText(text,'plain'), MIMEText(html, 'html')])
    # msg.attach(MIMEText(text_body, 'plain'))
    # msg.attach(MIMEText(html, 'html'))

    path = "C:\PTAX Project\Daily_Report_Tool\Output"
    rdate = datetime.datetime.strftime(today, '%Y-%m-%d')
    path = os.path.join(path + "/" + rdate)
    files = os.listdir(path)

    for filename in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(path + "/" + filename, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename={}'.format(filename))
        msg.attach(part)

    s = smtplib.SMTP('smtp.office365.com', 587)
    # start TLS for security
    s.starttls()
    sender = 'yadnesh.k@foxberry.in'

    recipients_to = eid_h_to.to_list()
    recipients_cc = eid_h_cc.to_list()
    rdate1 = datetime.datetime.strftime(today, '%d_%b_%Y')
    # msg['Subject'] = fl_name.upper() + ' ' + msgs + ' ' + rdate
    msg['Subject'] = "PCMC" + " | " + "PTAX" + " | " + "Collection_Report" + " | "  + rdate1
    msg['From'] = sender
    # msg['To'] = ' ,'.join([str(elem) for elem in recipients.split(',')[:1]])
    msg['To'] = ", ".join(recipients_to)
    msg['Cc'] = ", ".join(recipients_cc)
    # msg['Cc'] = ' ,'.join([str(elem) for elem in recipients.split(',')[1:]])
    # Authentication
    s.login("yadnesh.k@foxberry.in", "Kolhe@4321")
    s.sendmail(sender, recipients_to+recipients_cc, msg.as_string())
    # s.sendmail(sender,  recipients.split(','), msg.as_string())
    s.quit()


if __name__ == '__main__':
    send()
#     today = datetime.datetime.today()
#     std_path = r"C:\PTAX Project\PTAx\Manual Daily Report/"
#     in_path = std_path + "Input/" + str(today) + "/"
#     outpth = std_path + "Output/" + str(today) + "/"
#     mappath = std_path + "Mapping/"
#     logopath = std_path + "logo/"
#     mailreport = std_path + "Mail_report/"
#     send(today,mailreport)