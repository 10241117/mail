import win32com.client
import json
with open("config.json", encoding="utf-8") as json_file:
    config = json.load(json_file)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
folder = outlook.Folders(config["email"]).Folders('收件匣')

messages = (folder.Items.Restrict("@SQL=(urn:schemas:httpmail:subject LIKE '%" + config["subject"] + "%')")
                        .Restrict("@SQL=(urn:schemas:httpmail:datereceived >= '" + config["start_time"] + "' AND urn:schemas:httpmail:datereceived <= '" + config["end_time"] + "')"))

for message in messages:
    if hasattr(message, 'Sender'):
        print(message.Sender)
    print(message.SenderName)
    print(message.SenderEmailAddress)
    print(message.ReceivedTime)
    print(message.Subject)
    print(message.Body)
    attachments = message.attachments
    for attachment in attachments:
        print(attachment.filename) # 附件的檔名
        try:
            attachment.SaveAsFile(config["save_path"] + attachment.filename)
        except:
            print("路徑不存在")

