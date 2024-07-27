import requests
import datetime
import smtplib

url1 = 'https://api.sheety.co/468fa27b225c199db6181089b7355d81/mySpreadsheet/mySheet' # Retrieve rows from your sheet
url2 = 'https://api.sheety.co/468fa27b225c199db6181089b7355d81/mySpreadsheet/mySheet' # Add a row to your sheet
# Nichay wale dono main /[ObjectID] ata hai is ka matlab hota hai jis row ko edir ya delete 
# kerna hai uska row ka number ata hai
url3 = 'https://api.sheety.co/468fa27b225c199db6181089b7355d81/mySpreadsheet/mySheet' # Edit a row in your sheet
url4 = 'https://api.sheety.co/468fa27b225c199db6181089b7355d81/mySpreadsheet/mySheet' # Delete a row in your sheet
spreadsheet_url = '16mwEuj1VgI0SdUHA1G1ezJ-rKOKbf_zawZjLXiYyUD4'    

smtpgmail = 'smtp.gmail.com'
myemail = 'waleedkamal801@gmail.com'
password = 'my_password'
emailport = 587

def getdata():
    take = requests.get(url=url1)
    #take2 = take.json().get('mysheet',[])
    return take.json().get('mySheet',[])

# In this function is the date of interview is after 4 days after the message of test score
def editstatus1(i_id):
    d = datetime.datetime.now()
    e = d + datetime.timedelta(days=4)
    g = e.strftime("%d-%m-%Y")
    #print(e.strftime("%d-%m-%Y"))
    mySpreadsheet = {
        "mySheet": {
            'status':'Passed',
            'dateOfInterview':g
        }
    }

    edit_data = requests.put(url=f"{url3}/{i_id}",json=mySpreadsheet)
    edit_data.raise_for_status()
    print(edit_data.text)

# In this function is the date of interview is after 5 days after the message of test score
def editstatus2(i_id):
    d = datetime.datetime.now()
    e = d + datetime.timedelta(days=5)
    g = e.strftime("%d-%m-%Y")
    #print(e.strftime("%d-%m-%Y"))
    mySpreadsheet = {
        "mySheet": {
            'status':'Passed',
            'dateOfInterview':g
        }
    }

    edit_data = requests.put(url=f"{url3}/{i_id}",json=mySpreadsheet)
    edit_data.raise_for_status()
    print(edit_data.text) 

# In this function is the date of interview is after 6 days after the message of test score
def editstatus3(i_id):
    d = datetime.datetime.now()
    e = d + datetime.timedelta(days=6)
    g = e.strftime("%d-%m-%Y")
    #print(e.strftime("%d-%m-%Y"))
    mySpreadsheet = {
        "mySheet": {
            'status':'Passed',
            'dateOfInterview':g
        }
    }

    edit_data = requests.put(url=f"{url3}/{i_id}",json=mySpreadsheet)
    edit_data.raise_for_status()
    print(edit_data.text) 

# In this function is the date of interview is after 7 days after the message of test score
def editstatus4(i_id):
    d = datetime.datetime.now()
    e = d + datetime.timedelta(days=7)
    g = e.strftime("%d-%m-%Y")
    #print(e.strftime("%d-%m-%Y"))
    mySpreadsheet = {
        "mySheet": {
            'status':'Passed',
            'dateOfInterview':g
        }
    }

    edit_data = requests.put(url=f"{url3}/{i_id}",json=mySpreadsheet)
    edit_data.raise_for_status()
    print(edit_data.text) 

# In this function is the date of interview is after 8 days after the message of test score
def editstatus5(i_id):
    d = datetime.datetime.now()
    e = d + datetime.timedelta(days=8)
    g = e.strftime("%d-%m-%Y")
    #print(e.strftime("%d-%m-%Y"))
    mySpreadsheet = {
        "mySheet": {
            'status':'Passed',
            'dateOfInterview':g
        }
    }      

    edit_data = requests.put(url=f"{url3}/{i_id}",json=mySpreadsheet)
    edit_data.raise_for_status()
    print(edit_data.text)   

# In this function is the date of interview is after 8 days after the message of test score
def editstatus6(i_id):
    d = datetime.datetime.now()
    e = d + datetime.timedelta(days=10)
    g = e.strftime("%d-%m-%Y")
    #print(e.strftime("%d-%m-%Y"))
    mySpreadsheet = {
        "mySheet": {
            'status':'Failed',
            'dateOfInterview':g
        }
    }      

    edit_data = requests.put(url=f"{url3}/{i_id}",json=mySpreadsheet)
    edit_data.raise_for_status()
    print(edit_data.text)  

def sendmail(email,subject,message):
    try:
        s = smtplib.SMTP(smtpgmail,emailport)
        s.starttls()
        s.login(myemail,password)
        email_message = f"Subject:{subject}\n\n{message}"
        s.sendmail(myemail,email,email_message)
        print("Message sent")
    except:
        print("Message is not sent")    

if __name__=="__main__":
    # This is for get data from Google Sheet and Add the Status and Date of Interview in Google Sheet according to Student's Number Scored in the Test.
    takedata = getdata()
    for i in takedata:
        if i['obtainedMarks']>=i['minimumMarks'] and i['obtainedMarks']<=i['maximumMarks']:
            if i['obtainedMarks']>=90 and i['obtainedMarks']<=100:
                editstatus1(i['id'])
            elif i['obtainedMarks']>=80 and i['obtainedMarks']<=89:
                editstatus2(i['id'])
            elif i['obtainedMarks']>=70 and i['obtainedMarks']<=79:
                editstatus3(i['id'])
            elif i['obtainedMarks']>=60 and i['obtainedMarks']<=69:
                editstatus4(i['id'])      
            elif i['obtainedMarks']>=50 and i['obtainedMarks']<=59:
                editstatus5(i['id']) 
            else:
                print("Error!!!")
                
        else:
            editstatus6(i['id'])      

    # This is for get new data (After Entering Status and Date of Interview) and send email to all the Students who gave test in the Campus.
    takedata2 = getdata()
    for j in takedata2:
        print(j)
        if j['status']=='Passed':
            email = j['email']         
            subject = "Test Update"
            message = f"Dear {j['name']},\nyou have scored {j['obtainedMarks']} out of {j['maximumMarks']} and you have been {j['status']} and you have secure your seat of {j['department']} department.\nSo your interview is in {j['dateOfInterview']}"
            sendmail(email,subject,message)
        elif j['status']=='Failed':
            email = j['email']         
            subject = "Test Update"
            message = f"Dear {j['name']},\nUnfortunately you have {j['status']} because you have scored {j['obtainedMarks']} out of {j['maximumMarks']}. So you can visit our campus in {j['dateOfInterview']}, if there is a seat available then we can give you if we want."
            sendmail(email,subject,message)
