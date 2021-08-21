# Stock-Price-Reminder
Stock price reminder to mobile number/WhatsApp if it crosses either certain limits or if it increases by certain percentage, specified by user. Automatic scheduling of task according to requirement, runs continuously in the background with data capture to excel file.
Webscraping using beautiful soup library from money.rediff.com website.
Add the stocks to monitor to Share_list.xlsx with the URL of money.rediff.com website of the stock and also the limit price where you want to get a reminder.
The data of the stocks gets stored into Data.xlsx depending on your requirement use either twilio (sms to mobile number) or whatsapp reminder.
To run the code in background, open cmd and run the python file using the command "pythonw <filename.py>"
To schedule the task running use windows task scheduler, select the path of the python file.
