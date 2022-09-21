import win32com.client as win32
import pandas as pd
import numpy as np
import os
from datetime import datetime
pd.io.formats.excel.ExcelFormatter.header_style = None

def mailer (thisRec):
    recName = thisRec.loc[:,"First Name.1"].to_list()[0]
    recRet = thisRec.loc[:,"Are you a returning TEK Club member?"].to_list()[0]
    recEmail = thisRec.loc[:,'Email Address'].to_list()[0]
    
    mail = outlook.CreateItem(0)
    mail.Subject = "Thank You for Completing the Membership Form!"
    mail.Recipients.Add(recEmail).Type = 3
    
    if recRet == 'Yes':
        recMess = "Welcome back!  We appreciate you returning to join us for another year!"
    else: recMess = "Thank you so much for filling out our membership form! We are so excited to welcome you to the club, and you are almost there!"
    
    mail.HTMLBody = fr"""
    Dear {recName}, <br><br>
    {recMess}<br>
    The last thing you need to do (if you haven't already) is pay the membership dues!<br>
    The cost for the year is $40 USD, and there are three methods of paying, as follows:<br><br>
    1) Venmo: @TEK-club<br>
    2) Paypal: Click <a href="https://www.paypal.com/paypalme/tekclub/40">this</a> link<br>
    3) Cash:  Take the money in cash to Laurie Bragg at SFEBB 3113 (laurie.bragg@eccles.utah.edu), or bring it to one of our events.<br><br>
    Thank you again so much for your participation in our club!  Make sure to join the slack <a href="https://join.slack.com/t/tekclub/shared_invite/zt-goa1r1dw-V~pxALaFTgDMyET1BGOk9w">here</a>, follow us on instagram <a href="https://www.instagram.com/tekclub/">here</a>, and check out our upcoming events on the Google Calendar <a href="https://calendar.google.com/calendar/u/2?cid=dGVrY2x1YnVvZnVAZ21haWwuY29t">here</a>!<br>
    If you have any questions about your membership, feel free to email me at jacob.minson@utah.edu.<br>
    We are looking forward to a great year and are excited to have you join us!<br><br>
    All the best,<br>
    Jacob Minson<br>
    TEK Club Operations Director<br><br>
    """
    mail.Send()

def masterUpdate (thisRec):
    return pd.DataFrame([[thisRec.loc[:,'First Name.1'].to_list()[0],thisRec.loc[:,'Last Name.1'].to_list()[0],thisRec.loc[:,'UNID (U0000000)'].to_list()[0],thisRec.loc[:,'Email Address'].to_list()[0],thisRec.loc[:,'DateSubmitted'].to_list()[0].rsplit(" ")[0],np.nan,np.nan,np.nan]], columns=members.columns)
    
def moveForm():
    newForm = max([f for f in os.scandir(downDir)], key=lambda x: x.stat().st_ctime).name
    if 'FormSubmissions' in newForm:
        os.rename(downDir+"\\"+newForm, formDir+"\\"+"FormSubmission_"+datetime.today().strftime("%Y-%m-%d_%H-%M-%S")+'.csv')
    else: exit()
    

formDir = r"C:\Users\jacob\Documents\TEK\Member Form Records"
downDir = r"C:\Users\jacob\Downloads"
memDir = r"C:\Users\jacob\Box\TEK Club\Membership 2022-23\Member Form"

moveForm()
newForm = max([f for f in os.scandir(formDir)], key=lambda x: x.stat().st_ctime).name

thisForm = pd.read_csv(formDir+"/"+newForm, header=1)
thisForm = thisForm[['Status','Username', 'First Name.1', 'Last Name.1', 'Email Address', 'Are you a returning TEK Club member?','UNID (U0000000)','DateSubmitted']]
email_list = thisForm[thisForm['Status']=='Pending'].loc[:,'Email Address'].tolist()

outlook = win32.Dispatch('outlook.application')
members = pd.read_excel(memDir+"/"+'Member List.xlsx', index_col=0)
for recipient in email_list:
    thisRec = thisForm[thisForm['Email Address'] == recipient]
    mailer(thisRec)
    members = pd.concat([members,masterUpdate(thisRec)])
members.reset_index(drop=True, inplace=True)
members['Dues Paid Date'] = pd.to_datetime(members['Dues Paid Date'])
members['Dues Paid Date'] = members['Dues Paid Date'].dt.strftime('%m/%d/%Y')
members.to_excel(memDir+"/"+'Member List.xlsx')