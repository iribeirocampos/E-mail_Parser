import win32com.client
import unicodecsv as csv
import Analise_horas

output_file = open("sent_emails.csv", "wb")
output_writer = csv.writer(output_file, delimiter=";", encoding="latin2")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
outbox = outlook.GetDefaultFolder(5)
# "6" refers to the index of a folder - in this case,the inbox.
# "5" refers to the index of a folder - in this case,the outbox.

messages = outbox.Items
output_writer.writerow(["Date"])
for message in messages:
    try:
        # sender = message.SenderName
        # sender_address = message.sender.address
        # sent_to = message.To
        date = str(message.LastModificationTime)
        # subject = message.subject
        print(date)
        output_writer.writerow([date])
    except Exception:
        pass
output_file.close()

Analise_horas.data_anal()
