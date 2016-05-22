# OneContact
Chrome extension for Outlook.com that allows searching for contacts in another outlook.com account.

I use multiple Outlook.com accounts but don't want to manage my contacts in all of them so I wrote this extension to allow me to search for contacts. The extension needs to be started manually, in other words it is not started via content scripts. Once loaded it uses the To input box to trigger the search and essentially mimics default Outlook.com behaviour, except that the contacts are loaded from another account :)

P.S.
Make sure to set Outlook REST API client id and correct OAuth2 callback url in popup.js.
You can get Outlook REST API client id from (https://apps.dev.microsoft.com/).


