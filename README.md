# SaveOutlookEmails
Save and backup Outlook accounts and items (emails, appointments, attachments etc.) onto local drive.

[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?style=flat-square)](http://makeapullrequest.com)

## Purpose
In my Outlook only the last three months of emails are available offline, the rest are archived and moved into my [Online Archive - Name@Company.com](https://support.microsoft.com/en-gb/help/291626/how-to-manage-multiple-exchange-mailbox-accounts-in-outlook) account. Even when connected to the network the archived account only shows the first 200 odd characters of an email body and no attachments are available. This means that Outlook search won’t find anything from archived account.

My solution to this problem is to save all emails from all accounts onto my desktop where I can perform search in Windows Explorer: search within emails body and in attachments.

## Solution
__SaveOutlookEmails__ saves accounts from Outlook onto a desktop folder.
![ProgressBar](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/ProgressBar.jpg)
- Keep offline emails up-to-date date: autorun __SaveOutlookEmails__ when Outlook starts (at start of Outlook _Enable Macros_ when prompted with _'Microsoft Office has identified potential security concerns.'_)
- Save archived accounts: run __SaveOutlookEmails__ on selected folder (will take a while, run it at lunch time or at night, see more under _Warnings_)

Outlook's folder structure is kept the same and files are named with date-time prefix and shortened subject.
![Folder](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/Folder.jpg)

## Process
- Outlook folders are validated against `InvalidFolders` while Outlook items against `ValidItems`.
- Date and subject are checked whether the item has been saved before, if not then email validity checked in details.
- When `OLItem.MessageClass` ends with any words defined in `ArchivedArray` (in my case ends with _'EAS'_, my [email archive client](https://en.wikipedia.org/wiki/Enterprise_Archive_Solution_(EAS)), if it is different in your cases update `ArchivedArray` in _Config_) emails will be opened and then saved to get the full body and attachments.
- Size and number of recipients are limited (see _Config_).
- Outlook folder names and email subjects are cleaned for invalid characters.
- Subject dynamically shortened to fit into full path limit (255 characters on Windows).
- All successfully saved emails are added to __Log.txt__

## Features
- When auto run __SaveOutlookEmails__ items on local drive are checked using `fso.FileExists`. When the number of already saved emails reaches `OverlapResaved` and timeframe of already saved emails is over `OverlapDays` then scanning emails will stop. Autorun won’t open emails as recent items are part of the offline Outlook database, including attachments.
- When manually run on selected folders 'file exists' check is based on the _Log_ file. This check is a simple loop though the log array. After an email has been found then the next loop will start from where the previous has been found to shorten the loop time.

## Install
1. Add _Developer_ ribbon

![Add Developer Ribbon](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/1%20Add%20Developer%20ribbon.gif)

2. Check _Macro Settings_ in _Trust Canter_

![Add Developer Ribbon](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/2%20Check%20Macro%20Settings.gif)

3. Add _Microsoft Scripting Runtime_ in _VBA editor_

![Add Developer Ribbon](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/3%20Add%20Microsoft%20Scripting%20Runtime.gif)

4. Copy code files from [Code](https://github.com/licyp/SaveOutlookEmails/tree/master/Code) or [SaveOutlookEmails.zip](https://github.com/licyp/SaveOutlookEmails/blob/master/SaveOutlookEmails%20Ver1.0.zip)

![Add Developer Ribbon](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/4%20Copy%20code%20files%20from%20Code.gif)

5. Add auto run code to [ThisOutlookSession](https://github.com/licyp/SaveOutlookEmails/blob/master/Code/ThisOutlookSession.txt)

![Add Developer Ribbon](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/5%20Add%20auto%20run%20code.gif)

6. Add _Quick Access_ icon

![Add Developer Ribbon](https://github.com/licyp/SaveOutlookEmails/blob/master/Gif/6%20Add%20Quick%20Access%20icon.gif)

7. Hide _Developer_ ribbon

## Warnings
- To save archived emails with full body and attachments, they must be opened and then saved. This will happen automatically causing the ___screen to flicker___ and ___limit the use of the computer___.
- If connection is slow or IT limits access, then __SaveOutlookEmails__ will throw a run-time error:
Click on _'End'_, restart Outlook and try to run __SaveOutlookEmails__ again later.
- I am dyslexic therefore spelling mistakes are likely within the code.
- Not fully _'DRY'_, there is room for improvement.

## Configuration
- Maximum email size: `MaxItemSize = 25000000` 25MB
- Maximum number of recipients: `MaxItemTo = 250`
- Invalid characters: `* / \ : ? " % < > |`, `line feed`, `carriage return` and `horizontal tabulation`
- Overlap days: `OverlapDays = 7`
- Overlap number: `OverlapResaved = 100`
- Overlap subject: `OverlapSubject = 20` is used to left-compare email subject and file name
- Default folder on desktop: `DefultFolder = "Desktop\eMails"`

###### Current version: 1.2
###### [VBA - Visual Basic for Applications](https://docs.microsoft.com/en-us/office/vba/api/overview/outlook)
