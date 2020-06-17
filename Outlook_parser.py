import win32com.client
import unicodecsv as csv

output_file = open("Enviados.csv", "wb")
output_writer = csv.writer(output_file, delimiter=";", encoding="latin2")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
outbox = outlook.GetDefaultFolder(5)
# "6" refers to the index of a folder - in this case,the inbox.
# "5" refers to the index of a folder - in this case,the outbox.

messages = outbox.Items
output_writer.writerow(["Data"])
for message in messages:
    # sender = message.SenderName
    # sender_address = message.sender.address
    # sent_to = message.To
    date = str(message.LastModificationTime)
    # subject = message.subject
    print(date)
    output_writer.writerow([date])

output_file.close()
