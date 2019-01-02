
# coding: utf-8

# In[ ]:


import time
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import re
import winreg

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,                          r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
    return winreg.QueryValueEx(key, "Desktop")[0]

def pivot_table(dffiltered):
    result = pd.pivot_table(dffiltered, index=["PORT OF LOAD NAME"], values = ['CONTAINER COUNT'], 
                            aggfunc = [np.sum], margins = True, margins_name = 'GRAND TOTAL')
    result.columns = ['NUMBER OF CONTAINER']
    return result

def cust_report(customer):
    file_loc = r"%s\Dec 12-07.xlsx" %(get_desktop())
    df = pd.read_excel(file_loc, sheet_name='Sheet1',skiprows = 7)
    dffiltered = df.loc[df['CUSTOMER NAME'] == customer]
    return pivot_table(dffiltered)

def style_html(customer):
    htm = cust_report(customer).reset_index().to_html(index = False)
    a = """ cellspacing="0" cellpadding="0" style="text-align: center;border-collapse:collapse;">
  <thead>
  <caption style="text-align: center;font-weight:bold;">December Volume</caption>"""
    s = re.sub(r'class="dataframe">',a,htm)
    s = re.sub(r'style="text-align: right;"','',s)
    b = """
<tr style="font-weight:bold;">
      <td>GRAND TOTAL"""
    s = re.sub(r'<tr>\n      <td>GRAND TOTAL',b,s)
    return s

def generat_hist(customer):
    file_loc = r"%s\Historical.xlsx" %(get_desktop())
    df = pd.read_excel(file_loc, sheet_name='Sheet1')
    df = df.fillna(0)
    filtered = df.loc[df['Customer'] == customer]
    a = filtered.to_records()
    b = []
    for i in range (2,len(a[0])):   
        b.append(a[0][i])
    label = []
    label = df.columns.tolist()
    del label[0]
    for i in range(len(label)):
        label[i] = label[i][0:3]
    x = []
    for i in range(1,len(label)+1):
        x.append(i)
    #label = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov']
    plt.figure(figsize = (8,5))
    plt.bar(x,b,width = 0.5)
    plt.xlabel('Month')
    plt.ylabel('Number of Containers')
    plt.xticks(x,label)
    plt.title('Year to Date Volume')
    for x, y in zip(x, b):
        # ha: horizontal alignment
        # va: vertical alignment
        plt.text(x, y + 0.05, '%.f' % y, ha='center', va='bottom')
    plt.savefig('historical.png')
    plt.show()
    return 1

# my test mail
def send_email(customer):
    mail_username='*****'   # sender's email address
    mail_password='*****'
    from_addr = mail_username
    to_addrs=('*****')  #receiver's email address

    HOST = 'smtp.gmail.com'
    PORT = 25
    #HOST = 'smtp-mail.outlook.com'  # for outlook
    #PORT = 587
    
    # Create SMTP Object
    smtp = smtplib.SMTP()
    print('connecting ...')

    # show the debug log
    smtp.set_debuglevel(1)

    # connet
    try:
        print(smtp.connect(HOST,PORT))
    except:
        print('CONNECT ERROR ****')
    # gmail uses ssl
    smtp.starttls()
    # login with username & password
    try:
        print('loginning ...')
        smtp.login(mail_username,mail_password)
    except:
        print('LOGIN ERROR ****')
    # fill content with MIMEText's object 
    
    msgRoot = MIMEMultipart('related') 
    
    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    #邮件正文内容
    mail_body = """
<html>
      <head></head>
      <body>
<h3>To our valued customer:<br />
          <br><div style="text-align:center;"><br />%s<br /><img src="cid:logo_image" width="120" height="48" align="middle"></h3></div>
<table align="center" cellpadding="0" cellspacing="0"> <tr> 
<td>
    %s
</td> <td><div><img src="cid:ytd_image"></div></td> </tr></table>
          <br>Kind Regards<br />
          <br>ASF Global<br />
      </body>
    </html>
    """ %(customer,style_html(customer))
    
    msgText = (MIMEText(mail_body, 'html', 'utf-8'))
    msgAlternative.attach(msgText)


    # 指定图片为当前目录
    fp = open('historical.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    # 定义图片 ID，在 HTML 文本中引用
    msgImage.add_header('Content-ID', '<ytd_image>')
    msgRoot.attach(msgImage)

    fp = open('ag.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    # 定义图片 ID，在 HTML 文本中引用
    msgImage.add_header('Content-ID', '<logo_image>')
    msgRoot.attach(msgImage)
    
    msgRoot['From'] = from_addr
    msgRoot['To'] = to_addrs
    msgRoot['Subject']='Montly Report to %s' %(customer)
    #print(msgRoot.as_string())
    smtp.sendmail(from_addr,to_addrs,msgRoot.as_string())
    smtp.quit()
    return 1

file_loc = r"%s\customer_table.xlsx" %(get_desktop())
df1 = pd.read_excel(file_loc, sheet_name='Sheet1')
c = df1.to_records()
for i in range(0,len(c)):
    generat_hist(c[i][1])
    time.sleep(1)
    send_email(c[i][2])

