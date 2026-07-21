import win32com.client
import pandas as pd

# Outlook connection
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Shared mailbox and Folder name
mailbox_name = "Portale Cardea Tech Supp"
folder_name = 'Inbox'
subfolder_name = '1_Gestite_2024'

# Shared mailbox selection
try:
    shared_mailbox = outlook.Folders(mailbox_name)
except:
    shared_mailbox = outlook.GetDefaultFolder(6)

# Folder selection
inbox_folder = shared_mailbox.Folders(folder_name)
mail_folder = inbox_folder.Folders(subfolder_name)
#mail_folder = shared_mailbox.Folders(folder_name)

# Data list
data = []

for mail in mail_folder.Items:
    if mail.Class == 43:   #Codice identificativo dell'email singola
        if len(data) < 100:
            body = mail.Body.strip()
            if 'name.surname@email.com' in body:
                body = body.split('name.surname@email.com')[-1].strip()
                data.append([
                    mail.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                    mail.SenderName,
                    mail.Subject,
                    mail.To,
                    mail.CC,
                    mail.Body
                ])
            print(len(data))
        else:
            break


df = pd.DataFrame(data, columns=['Data', 'Mittente', 'Oggetto', 'TO', 'CC', 'Corpo'])

df.to_excel('output_name.xlsx', index=False)
