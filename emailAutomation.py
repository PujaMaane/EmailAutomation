import win32com.client


class OutlookMails:

    @staticmethod
    def connectingToMail():
        outlookMail = win32com.client.Dispatch('outlook.application')
        mapi = outlookMail.GetNamespace("MAPI")
        inboxMsg = mapi.GetDefaultFolder(6)
        """for account in mapi.Accounts:
            print(account.DeliveryStore.DisplayName)"""
        return inboxMsg

    @classmethod
    def readMails(cls, inbox):
        messages = inbox.Items
        message = messages.GetLast()

        sender = message.SenderName
        sender_address = message.SenderEmailAddress
        sent_to = message.To
        date = message.LastModificationTime
        subject = message.Subject
        body = message.body

        print("Sender:", sender)
        print("Sender Mail:", sender_address)
        print("Sent TO:", sent_to)
        print("Date:", date)
        print("Subject:", subject)
        print("Body:", body)


if __name__ == "__main__":
    outlook = OutlookMails()
    inbox = outlook.connectingToMail()
    outlook.readMails(inbox)
