# KeepNewOutlookOpen

A PowerShell script to keep the new Outlook open and hidden, without closing unexpectedly.

It's a disaster if you accidentally (Alt + F4, I like it) close the new Outlook, then lots of emails come in.

I could not find a helpful tool (the old COM Add-In from [SourceForge/keepoutlook](https://sourceforge.net/projects/keepoutlook/) is not working in my classical Outlook) to keep Outlook hidden when closed. Microsoft doesn't like to develop this small functionality, maybe because it's not cool.

This script checks if the new Outlook process `olk` exists; if not, open it and try to hide its Windows (with class name `Olk Host`).

You can add a Windows task to trigger this script on the user log-on and then run periodically.
