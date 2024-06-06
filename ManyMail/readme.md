# MailMany Outlook Addin - Send Personalized Mails to a (large) number of Recipients

My main objective when I started dabbling in Outlook Addins was to create a tool that allows me to easily send mails to a potentially large group of people with personal elements - such as their first name and company name in addition to their mail address. Something like (Word's) MailMerge - but then much better. I want to run it entirely from within Outlook (and not create a document in Word). I did reach that goal and this article shows the result achieved and explains how I did it. It was quite a bit harder than I anticipated. For one: the JavaScript API for Office allows addins to do many things, including reading, creating and saving email messages. But it does not allow an email to be sent! In this article how I worked around the limitation of not being allowed to send an email in order to send many mails. One clue: the workaorund only works for the Outlook Web Client. A second clue: browser extension. More on that later.

This article is more about how to *create* an Outlook addin than about how to use or even suggesting that you should use my addin. You will learn about some of the challenges I encountered and the workarounds I implemented. As an addin, it is not great - just a prototype. It works and it can provide some inspiration for functionality for you to implement.

## Using the Addin

Once the Addin is installed (and the Browser Extension is enabled), I can send a personalized email to many recipients by going through these steps:
* write an email with placeholders that can be personalized (a placeholder is written as `{{property}}`)
![](images/email-to-me.png)
* send the email to myself
* open the email in read mode
* activate the addin, open the taskpane and paste the recipients data in CSV format
![](images/paste-csv-process-data.png)
* press the *Process Data* button to process the data and get a structured overview 
 ![](images/data-processed-in-tabvle.png) 
 * When the data is correct, press the button *Send email to All Recipients*
 * The addin will now open new windows for all all recipients, containing the personalized emails
  ![](images/send-emails.png)
 * the Chrome Browser extension will identify each of these windows and it will send the emails
  ![](images/receive-personalized-mail.png) 



This Addin is activated when reading a message.

The current message can be sent to a list of recipients - and can be personalized before it is sent.

The Addin opens a new message form for every recipient, with the personalized to, subject and message body in the form.

In the desktop client, you have to send each mail message manually. In the web client, there is an additional automation available: The addin defines a special bcc value: autosend@contoso.com. If you use the Outlook WebClient and you have installed and enabled the Chrome Extension Outlook New Mail Auto Sender, then this value will trigger the automatically sending of the mail message (the extension will locate and click the Send button in the new message form). 