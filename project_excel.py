import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import itertools
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import smtplib
from IPython.display import HTML
from IPython.display import display
import numpy as np
from configparser import ConfigParser
import sys
import re

'''
    Docstring
    if not present: Check the Excel sheet exist and if not print the error message
    if present : Creating the data frame for the two sheets

'''

try:
    df1 = pd.read_excel('C:\\Users\\dakshayani\\Desktop\\Book3.xlsx', sheet_name='EmpMaster')
    df2 = pd.read_excel('C:\\Users\\dakshayani\\Desktop\\Book3.xlsx', sheet_name='Monthly_DetailedReport')
except OSError:
    print("Hey!! you have an OS error, please check the excel sheet present in the location")
    sys.exit()

'''

    Docstring
    get employee from the master sheets

'''

def get_employee_detail():
    get_employee_id = pd.DataFrame(df1,columns=["Employee", "Emp ID","Exa Name Byd"])
    return get_employee_id

'''

    search the employee Id exist in Detailed report sheet and extract the
    rows relevant to the Employee ID and display
    else
    if not present print not exist

'''

def search_employee_id(check_employee_id, mailID, name):
    my_result = df2.iloc[:, 10]
    #print(check_employee_id)
    for i in my_result:
        if i == float(check_employee_id):
            for row in range(df2.shape[0]): # df is the DataFrame
                for col in range(df2.shape[1]):
                    if df2.iat[row,col] == i:
                        test = df2.iloc[row+2:row+9,0:col-2]
                        cols = [0,2]
                        result = test.drop(test.columns[cols], axis = 1)
                        result_row = result.drop(result.index[[0,3,4,5]])
                        result_row.columns = [' ', df2.iloc[10,3], df2.iloc[10,4],df2.iloc[10,5],df2.iloc[10,6],df2.iloc[10,7]]
                        final_df = (
                                   result_row.style
                                   .hide_index()
                                   .set_properties(**{'background-color': 'lightblue','color': 'black','border-color': 'black', 'width': '100px', "text-align": "center"})
                                   .render()
                        )
                        send_mail(final_df, mailID, name,check_employee_id)
        else:
            #print("No record matched in the particular cell of the excel sheet")
            pass


'''

    Docstring
    function to send the mail

'''


def send_mail(data_to_mail, recipients_mailID,name,ID):
    #print(recipients_mailID)
    regex = '^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$'
    if(re.search(regex,recipients_mailID)):
        #print("valid email")
        config_object = ConfigParser()
        config_object.read("config.ini")
        userinfo = config_object["EmailInfo"]
        from_email = userinfo["EmailID"]
        msg = MIMEMultipart()
        msg['Subject'] = userinfo["Subject"]
        msg['From'] = userinfo["EmailID"]
        html = """\
        <html>
          <head></head>
          <body>
          Employee Name : {0}<br>
          Employee ID   : {1}<br>
                {2}
          </body>
        </html>
        """.format(name,ID,data_to_mail)

        part1 = MIMEText(html, 'html')
        msg.attach(part1)
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(from_email, userinfo["Password"])
        server.sendmail(msg['From'], recipients_mailID , msg.as_string())
        print("Successfully sent email to the {0}".format(name))
    else:
        print("Invalid email: {0}, Please check the email ID in Master data Sheet".format(recipients_mailID))



if __name__ == '__main__':
    employee_id = get_employee_detail()
    for i, j in employee_id.iterrows():
        employee_id = j[0][3:6]
        employee_mailID = j[1]
        employee_name = j[2]
        #print(employee_name)
        search_employee_id(employee_id, employee_mailID, employee_name)
