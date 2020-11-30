import pandas as pd
import datetime
import win32com.client as client

df=pd.read_excel('myConcerto Birthday.xlsx',sheet_name='Master')
df=df.loc[df['Country']=='India',['EnterpriseID','ResourceName','DOB(DateOfBirth)','Roll_off_Date']].dropna(subset=['EnterpriseID','ResourceName','DOB(DateOfBirth)'])

df.fillna(datetime.datetime(2050,10,10) , inplace=True)

df['DOB(DateOfBirth)_short'] = df['DOB(DateOfBirth)'].dt.strftime('%B %d')
today_date_short = datetime.date.today().strftime('%B %d')

birthday = df.loc[df['DOB(DateOfBirth)_short'] == today_date_short , :]

birthday_df = birthday.loc[birthday['Roll_off_Date'].dt.date  > datetime.date.today(),:]

input_data = list(zip(birthday_df['ResourceName'],birthday_df['EnterpriseID']))

def email_trigger(name,email):

    outlook = client.Dispatch('Outlook.Application')
    message = outlook.CreateItem(0)
    html_body = """

        <div style="font-family: 'Graphik';" font-size: 15;">
        <div style="background-image: url('https://images.unsplash.com/photo-1513151233558-d860c5398176?crop=entropy&cs=tinysrgb&fit=crop&fm=jpg&h=900&ixid=eyJhcHBfaWQiOjF9&ixlib=rb-1.2.1&q=80&w=1600');">
        <img src="https://i1.fnp.com/assets/images/custom/birthday-micro/top-categories/Birthday-Best-Seller-Gifts-10-sept-2019.jpg" width=900 height=500 px>
        <br><p style="margin-left: 30px;">Happy Birthday <b> {} </b> !<br><br>
        We hope you have a wonderful day and cherish your time with family, friends and dear ones.<br><br>
        Wish you all the best for the year ahead. <br><br>
        <span><b>Regards, <br>Virtuso Team</b><span><br><br>Advanced Technology Centers, India<br><br>
        <img src="https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcQsWgIpifyX0VOHVTOOgyTwywCDyQ-IDryzcQ&usqp=CAU" width=200 height=60>
        </p>
        
    </div><br>
        <p><b>*Please do not reply to this e-mail as it is system generated.</b></p>
        """.format(name.split()[0])

    message.To = email
    message.CC = "a.a.gaurav@accenture.com"
    message.Subject = "Happy Birthday {} !".format(name.split()[0])
    message.HTMLBody = html_body
    message.Send()

for name , email in input_data:
    email_trigger(name , email)


def sat_sun_email_trigger(name,email):

    outlook = client.Dispatch('Outlook.Application')
    message = outlook.CreateItem(0)
    html_body = """

        <div style="font-family: 'Graphik';" font-size: 15;">
        <div style="background-image: url('https://images.unsplash.com/photo-1513151233558-d860c5398176?crop=entropy&cs=tinysrgb&fit=crop&fm=jpg&h=900&ixid=eyJhcHBfaWQiOjF9&ixlib=rb-1.2.1&q=80&w=1600');">
        <img src="https://i1.fnp.com/assets/images/custom/birthday-micro/top-categories/Birthday-Best-Seller-Gifts-10-sept-2019.jpg" width=900 height=500 px>
        <br><p style="margin-left: 30px;">Belated Happy Birthday <b> {} </b> !<br><br>
        We hope you have a wonderful day and cherish your time with family, friends and dear ones.<br><br>
        Wish you all the best for the year ahead. <br><br>
        <span><b>Regards, <br>Virtuso Team</b><span><br><br>Advanced Technology Centers, India<br><br>
        <img src="https://encrypted-tbn0.gstatic.com/images?q=tbn%3AANd9GcQsWgIpifyX0VOHVTOOgyTwywCDyQ-IDryzcQ&usqp=CAU" width=200 height=60>
        </p>
        
    </div><br>
        <p><b>*Please do not reply to this e-mail as it is system generated.</b></p>
        """.format(name.split()[0])

    message.To = email
    message.CC = "a.a.gaurav@accenture.com"
    message.Subject = "Belated Happy Birthday {} !".format(name.split()[0])
    message.HTMLBody = html_body
    message.Send()

if datetime.date.today().strftime('%A') == "Monday":
    
    sat_sun_date = [(datetime.date.today()-datetime.timedelta(1)).strftime('%B %d') , (datetime.date.today()-datetime.timedelta(2)).strftime('%B %d') ]

    birthday_sat_sun_df = df.loc[df['DOB(DateOfBirth)_short'].isin(sat_sun_date) , :]
    birthday_sat_sun_df = birthday_sat_sun_df.loc[birthday_sat_sun_df['Roll_off_Date'].dt.date  > datetime.date.today(),:]
    sat_sun_input_data = list(zip(birthday_sat_sun_df['ResourceName'],birthday_sat_sun_df['EnterpriseID']))

    for name , email in sat_sun_input_data:
        sat_sun_email_trigger(name , email)   
