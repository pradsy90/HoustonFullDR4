# http://itsecmedia.com/blog/post/2016/python-send-outlook-email/
import win32com.client as client

recipient = "john@example.com"
cc = "peter@example.com"


def create_email(recipient, cc, subject,body):
    """Create an email with summary information"""
    outlook = client.Dispatch('Outlook.Application')
    message = outlook.CreateItem(0)
    message.To = recipient
    message.CC = cc
    #    message.Attachments.Add(file) # You can attach a file with this
    message.Subject = subject
    message.HTMLBody = body
    message.Display()


def main():
    create_email(recipient, cc)

if __name__ == "__main__":
    main()

# os.startfile('C:\Program Files\Microsoft Office\\root\Office16\OUTLOOK.EXE')
