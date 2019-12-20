import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import itertools
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import smtplib

'''
    Docstring
    Creating the data frame for the two sheets

'''
df1 = pd.read_excel('C:\\Users\\dakshayani\\Desktop\\Book3.xlsx', sheet_name='EmpMaster')
df2 = pd.read_excel('C:\\Users\\dakshayani\\Desktop\\Book3.xlsx', sheet_name='Monthly_DetailedReport')

'''
    Docstring
    get employee from the master sheets

'''

def get_employee_id():
    get_employee_id = pd.DataFrame(df1,columns=["Employee"])
    return get_employee_id

'''
    search the employee Id exist in AC sheet and extract the
    rows relevant to the person and display

'''

def search_employee_id(check_employee_id):
    my_result = df2.iloc[:, 10]
    #print(check_employee_id)
    for i in my_result:
        if i == float(check_employee_id):
            for row in range(df2.shape[0]): # df is the DataFrame
                for col in range(df2.shape[1]):
                    if df2.iat[row,col] == i:
                        print(row, col)
                        row_data = row + 9
                        test = df2.iloc[row+2:row+9,0:col-2]
                        send_mail(test)
    else:
        pass


'''
    Docstring
    function to send the mail

'''
def send_mail(data_to_mail):
    # get the mail ID from the masterdata
    mailID = pd.DataFrame(df1,columns=["Emp ID"])
    for i, j in mailID.iterrows():
        recipients_mailID = j[0]
        #print(recipients_mailID)
        from_email = "dakshayani.r@exa-ag.com"
        msg = MIMEMultipart()
        msg['Subject'] = "Employee Attendence"
        msg['From'] = 'dakshayani.r@exa-ag.com'
        html = """\
        <html>
          <head></head>
          <body>
            {0}
          </body>
        </html>
        """.format(data_to_mail.to_html())

        part1 = MIMEText(html, 'html')
        msg.attach(part1)
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(from_email, "password")
        server.sendmail(msg['From'], recipients_mailID , msg.as_string())


if __name__ == '__main__':
    employee_id = get_employee_id()
    for i, j in employee_id.iterrows():
        data = j[0][3:6]
        #result = list(data.split("\n"))
        search_employee_id(data)
