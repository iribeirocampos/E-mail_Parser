# E-mail_Parser
Tool to calculate overtime worked through sent e-mails timestamp. 

## Installing
Install the required packages:

```
pip3 install -r requirements.txt
```
#### Outlook
You need to disable cached exchange mode in order to get all e-mais from server. You can see how in https://support.microsoft.com/en-us/office/turn-on-cached-exchange-mode-7885af08-9a60-4ec3-850a-e221c1ed0c1c?ui=en-us&rs=en-us&ad=us.

## Deployment

```
Run `python3 Outlook_parser.py`
```

## Output

file "sent_emails.csv", with all the dates and hours of all the sent e-mails in your Outlook server

file "overtime.xlsx", with que overtime arranged by month/year over the weekdays


![image](Preview.PNG)
