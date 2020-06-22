import win32com.client
import unicodecsv as csv
import data_analyze
from enum import Enum

output_file = open("sent_emails.csv", "wb")
output_writer = csv.writer(output_file, delimiter=";", encoding="latin2")


class OutlookFolder(Enum):
    OUTBOX = 5
    INBOX = 6


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
outbox = outlook.GetDefaultFolder(OutlookFolder.OUTBOX.values)

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
        print("Date not aquired.")
output_file.close()

data_analyze.data_anal()
