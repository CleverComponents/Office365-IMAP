# Connect to Office365 IMAP using OAuth 2.0

<img align="left" src="https://www.clevercomponents.com/images/office365-logo-250.jpg" />

Microsoft provides two different ways to access Outlook / Office365 mail: Outlook REST APIs and Live Connect APIs. The Outlook REST APIs allows you to manage mail / calendar / contacts in the same way as Google APIs does. This approach requires a special client library that implements necessary APIs REST commands and handles server responses. You can still get access to Microsoft Outlook mail using the standard IMAP4 protocol. Live Connect APIs provides you with this functionality. Microsoft has deprecated the non-secure user / password authorization algorithm. It may stop working in the nearest future. However, if your app is using IMAP with OAUTH 2.0, it will continue working.

The introduced program shows how to connect to Microsoft Outlook.com IMAP using OAUTH 2.0. In addition, this program utilizes a fast and easy algorithm of retrieving information about mailbox messages in one single IMAP command.

The program utilizes the following Internet components from the [Clever Internet Suite library](https://www.clevercomponents.com/products/inetsuite/): TclIMAP4 and TclOAUTH.

[Read the Article](https://www.clevercomponents.com/articles/article049/)

Join us on   [Facebook](http://www.facebook.com/clevercomponents)   [Twitter](https://twitter.com/CleverComponent)   [Telegram](https://t.me/clevercomponents)   [Newsletter](https://www.clevercomponents.com/home/maillist.asp)
