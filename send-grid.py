import sendgrid
import xlrd
import os

sg = sendgrid.SendGridAPIClient(api_key=os.environ.get('SENDGRID_API_KEY'))
loc = ("Emails.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
emails = []
counter = 0
for i in range(sheet.nrows):
    email = {"email":sheet.cell_value(i,0)}
    emails.append(email)
    i += 1
message = input("enter your message:\n")
data = {
  "personalizations": [
    {
      "to": emails
      ,
      "subject": "Testing sendGrid"
    }
  ],
  "from": {
    "email": "aatheen.ds@gmail.com"
  },
  "content": [
    {
      "type": "text/plain",
      "value": message
    }
  ]
}
response = sg.client.mail.send.post(request_body=data)
print(response.status_code)
print(response.body)
print(response.headers)
