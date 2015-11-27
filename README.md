# OutlookEmailAddressExtractor
Sometimes it's handy to be able to easily extract a list of email addresses from your Outlook PST or OST file.  
The 'old fashioned' way of doing this is to export the mailbox to a CSV file, and only include email addresses as a field.
But with large mailboxes, this can be time consuming and cumbersome, and outlook will often include internal sender ID's instead of email addresses, which are not overly useful.

## What this application does

This application simply opens an Outlook profile and iterates through every message, building a collection of email addresses as it goes.

It also counts how often an email address is used, as this is a useful metric.

## Prerequisites

The application makes use of the Microsoft.Office.Interop.Outlook assembly.

You need to have Microsoft Outlook installed on your PC - otherwise the Interop assembly has nothing to talk to.

## Author

This project was created by and maintained by Matthew Proctor [@mattproctorau](https://twitter.com/mattproctorau), as an example of how to extract email addresses from Microsoft Outlook.

## Where to learn more?
Learn more about this project at [https://www.matthewproctor.com/](https://www.matthewproctor.com/)
