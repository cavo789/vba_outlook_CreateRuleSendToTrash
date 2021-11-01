# Outlook - Create rule to remove emails received by spammer (email or domain)

![Banner](./banner.svg)

> Outlook macro for creating a rule that will send all emails received from a specific sender / domain to the trash.

## Description

## Table of Contents

- [Install](#install)
- [Usage](#usage)
- [Author](#author)
- [License](#license)

## Install

Get a copy of the `module.bas` VBA code and copy it into your Outlook client.

- Press `ALT-F11` in Outlook to open the `Visual Basic Editor` (aka VBE) window.
- Create a new module and copy/paste the content of the `module.bas` file that you can found in this repository
- Close the VBE
- Right-click on your Outlook ribbon to customize it so you can add a new button. Assign the `CreateRuleSendToTrash` subroutine to that button.

## Usage

1. From Outlook, select one or more spam emails
2. Run the macro below (best : add a button in your ribbon)

What does this macro ?

- Get the email address of the spammer (f.i. spam@yahoo.co.uk)
- Ask the user if emails from that domain (@yahoo.co.uk) should be removed from now or only those sender (spam@yahoo.co.uk)
- Create a rule in the "DefaultStore" of Outlook for removing emails with that emails/domain and send mails to the deleted items folder
- Iterate all accounts present in your Outlook and run rules against all accounts. So, every if the rule is created only in the "DefaultStore" (i.e. one account), the rule will be fired for all configured accounts.

## Author

Christophe Avonture

Inspired by [https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/create-a-rule-to-move-specific-e-mails-to-a-folder](https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/create-a-rule-to-move-specific-e-mails-to-a-folder) and by
[https://social.msdn.microsoft.com/Forums/office/en-US/161ca567-9b6f-4155-9833-cc371d78b66e/vba-macro-to-run-rules-for-multiple-email-accounts?forum=outlookdev](https://social.msdn.microsoft.com/Forums/office/en-US/161ca567-9b6f-4155-9833-cc371d78b66e/vba-macro-to-run-rules-for-multiple-email-accounts?forum=outlookdev)

## License

[MIT](LICENSE)
