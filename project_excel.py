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

def get_employee_detail():
    get_employee_id = pd.DataFrame(df1,columns=["Employee", "Emp ID","Exa Name Byd"])
    return get_employee_id

'''def get_employee_mailID():
    # get the mail ID from the masterdata
    mailID = pd.DataFrame(df1,columns=["Emp ID"])
    for i, j in mailID.iterrows():
        recipients_mailID = j[0]
        #print(recipients_mailID)'''

'''
    search the employee Id exist in AC sheet and extract the
    rows relevant to the Employee ID and display

'''

def search_employee_id(check_employee_id, mailID, name):
    my_result = df2.iloc[:, 10]
    #print(check_employee_id)
    for i in my_result:
        if i == float(check_employee_id):
            for row in range(df2.shape[0]): # df is the DataFrame
                for col in range(df2.shape[1]):
                    if df2.iat[row,col] == i:
                        row_data = row + 9
                        test = df2.iloc[row+2:row+9,0:col-2]
                        cols = [0,2]
                        result = test.drop(test.columns[cols], axis = 1)
                        result_row = result.drop(result.index[[0,3,4,5]])
                        result_row.columns = [' ', df2.iloc[10,3], df2.iloc[10,4],df2.iloc[10,5],df2.iloc[10,6],df2.iloc[10,7]]
                        #header_result = result_row.head()
                        #print(header_result)
                        #result_row['Average'] = result_row.mean(axis=1)
                        #print(result_row)
                        final_df = (
                                   result_row.style
                                   .hide_index()
                                   .set_properties(**{'background-color': 'lightblue','color': 'black','border-color': 'black', 'width': '100px', "text-align": "center"})
                                   .render()
                        )
                        send_mail(final_df, mailID, name,check_employee_id)
    else:
        pass


'''

    Docstring
    function to send the mail

'''


def send_mail(data_to_mail, recipients_mailID,name,ID):
    print(recipients_mailID)
    from_email = "dakshayani.r@exa-ag.com"
    msg = MIMEMultipart()
    msg['Subject'] = "Employee Attendence"
    msg['From'] = 'dakshayani.r@exa-ag.com'
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
    server.login(from_email, "Password")
    server.sendmail(msg['From'], recipients_mailID , msg.as_string())


if __name__ == '__main__':
    employee_id = get_employee_detail()
    for i, j in employee_id.iterrows():
        employee_id = j[0][3:6]
        employee_mailID = j[1]
        employee_name = j[2]
        #print(employee_name)
        search_employee_id(employee_id, employee_mailID, employee_name)
