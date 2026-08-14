# set-default-mail-app

________________________
default_mailto_handler_change.zsh

This mailto handler script sets Microsoft Outlook as the default application for mailto: links for the currently signed-in Mac user.

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

________________________

prompt_user_to_set_outlook_default_for_eml.zsh

This script sets Microsoft Outlook as the default application for saved .eml email files for the currently signed-in Mac user.

How it works:

Identifies the signed-in user

NinjaOne normally runs scripts as root, but file-opening preferences belong to each individual user. The script therefore finds the active console user and performs the relevant actions inside that user’s graphical session.

Confirms Outlook is installed

It checks for:

/Applications/Microsoft Outlook.app

If Outlook is missing, the script exits without making changes.

Creates a temporary .eml file

The script creates a harmless sample email containing only fictitious example.invalid addresses. This gives macOS a real file from which it can determine the correct content type.

Checks the current handler

It asks NSWorkspace which application would currently open the temporary .eml file.

If the returned bundle identifier is already:

com.microsoft.Outlook

the script exits successfully without prompting or changing anything.

Builds a temporary helper application

Apple’s default-application API is asynchronous and may need to display a user-consent window. A normal background script does not remain attached to the graphical environment reliably enough for that.

The script therefore creates a temporary AppleScript application and launches it in the signed-in user’s session.

Requests the content-type change

The helper calls:

NSWorkspace.setDefaultApplication(
    Outlook,
    for the content type of the sample .eml file
)

This is the programmatic equivalent of:

Get Info → Open with → Microsoft Outlook → Change All…

macOS requests confirmation

Because this changes the default application for every file of that content type, macOS may require the user to approve it. The script cannot suppress that confirmation using the supported API.

Verifies the result

After the helper finishes, the script asks NSWorkspace which application would open the sample .eml file again.

It reports success only if the returned bundle identifier is:

com.microsoft.Outlook

Cleans everything up

The temporary email, AppleScript source, and helper application are automatically deleted.

This affects saved .eml files only. It does not change:

mailto: links
Outlook’s sending account
.msg, .emlx, or .mbox files
Preferences belonging to other users on the Mac

________________________
