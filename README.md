# set-default-mail-app
made this kuz i was bored and spiteful.

This script sets Microsoft Outlook as the default application for mailto: links for the currently signed-in Mac user.

How it works:

It identifies the signed-in user because NinjaOne normally executes scripts as root, while default-app choices belong to each individual user.
It verifies that Microsoft Outlook is installed.
It uses Apple’s NSWorkspace framework to check the current mailto: handler.
If Outlook is already the handler, it exits successfully without changing anything.
Otherwise, it builds a temporary AppleScript application and launches it inside the user’s graphical session.
That temporary application uses Apple’s supported NSWorkspace.setDefaultApplication API to assign Outlook to the mailto URL scheme.
The temporary application stays active briefly so macOS has time to process the asynchronous request.
The script queries NSWorkspace again to confirm that Outlook actually became the handler.
Finally, it deletes the temporary application and reports success or failure.

What it affects:

mailto:user@example.com → Microsoft Outlook

This includes email links clicked in browsers, documents, and other applications.

What it does not necessarily affect:

Which application opens saved .eml email files
The default sending account inside Outlook
Browser-specific webmail handlers
Other users on the same Mac

It does not directly edit Launch Services preference files, require duti, or require Xcode/Swift to be installed. The change is made through macOS’s native application-handling API.
