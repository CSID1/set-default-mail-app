#!/bin/bash

PATH="/usr/bin:/bin:/usr/sbin:/sbin"

TARGET_APP="/Applications/Microsoft Outlook.app"
TARGET_BUNDLE_ID="com.microsoft.Outlook"

log() {
    echo "[DefaultMail] $*"
}

console_user=$(/usr/bin/stat -f "%Su" /dev/console)

if [[ -z "$console_user" || "$console_user" == "root" || "$console_user" == "loginwindow" ]]; then
    log "No signed-in user was detected."
    exit 1
fi

console_uid=$(/usr/bin/id -u "$console_user")

if [[ ! -d "$TARGET_APP" ]]; then
    log "Outlook was not found at: $TARGET_APP"
    exit 1
fi

run_as_console_user() {
    /bin/launchctl asuser "$console_uid" \
        /usr/bin/sudo -H -u "$console_user" "$@"
}

get_mailto_handler() {
    run_as_console_user /usr/bin/osascript <<'APPLESCRIPT'
use framework "AppKit"
use framework "Foundation"

set workspace to current application's NSWorkspace's sharedWorkspace()
set mailtoURL to current application's NSURL's URLWithString:"mailto:test@example.invalid"
set applicationURL to workspace's URLForApplicationToOpenURL:mailtoURL

if applicationURL is missing value then
    return "NONE"
end if

set applicationBundle to current application's NSBundle's bundleWithURL:applicationURL

if applicationBundle is missing value then
    return applicationURL's |path|() as text
end if

set bundleIdentifier to applicationBundle's bundleIdentifier()

if bundleIdentifier is missing value then
    return applicationURL's |path|() as text
end if

return bundleIdentifier as text
APPLESCRIPT
}

current_handler=$(get_mailto_handler)

log "Signed-in user: $console_user"
log "Current mailto handler: $current_handler"

if [[ "$current_handler" == "$TARGET_BUNDLE_ID" ]]; then
    log "Outlook is already the default mailto handler."
    exit 0
fi

temporary_directory=$(/usr/bin/mktemp -d "/private/tmp/default-mail.XXXXXX")
applet_source="$temporary_directory/main.applescript"
applet_path="$temporary_directory/Default Mail Setup.app"

cleanup() {
    if [[ -n "$temporary_directory" && -d "$temporary_directory" ]]; then
        /bin/rm -rf "$temporary_directory"
    fi
}

trap cleanup EXIT

/bin/cat > "$applet_source" <<'APPLESCRIPT'
use framework "AppKit"
use framework "Foundation"
use scripting additions

on run
    set applicationObject to current application's NSApplication's sharedApplication()

    applicationObject's setActivationPolicy:(current application's NSApplicationActivationPolicyRegular)
    applicationObject's activateIgnoringOtherApps:true

    set workspace to current application's NSWorkspace's sharedWorkspace()
    set outlookURL to current application's NSURL's fileURLWithPath:"/Applications/Microsoft Outlook.app"

    workspace's setDefaultApplicationAtURL:outlookURL ¬
        toOpenURLsWithScheme:"mailto" ¬
        completionHandler:(missing value)

    -- Keep this GUI process and its Cocoa event loop alive so macOS can
    -- display and process the default-application consent request.
    set expirationDate to current application's NSDate's dateWithTimeIntervalSinceNow:45
    current application's NSRunLoop's currentRunLoop()'s runUntilDate:expirationDate
end run
APPLESCRIPT

log "Building temporary user-facing NSWorkspace helper."

/usr/bin/osacompile -o "$applet_path" "$applet_source"

compile_status=$?

if [[ "$compile_status" -ne 0 || ! -d "$applet_path" ]]; then
    log "Unable to build the temporary helper."
    exit 1
fi

/usr/sbin/chown -R "$console_uid" "$temporary_directory"
/bin/chmod -R u+rwX,go+rX "$temporary_directory"

log "Launching the NSWorkspace helper for $console_user."
log "A macOS confirmation window may now appear."

run_as_console_user /usr/bin/open -W -n "$applet_path"

launch_status=$?

if [[ "$launch_status" -ne 0 ]]; then
    log "The temporary helper returned exit code $launch_status."
fi

current_handler=$(get_mailto_handler)

log "Mailto handler after the request: $current_handler"

if [[ "$current_handler" == "$TARGET_BUNDLE_ID" ]]; then
    log "Success: Outlook is now the default mailto handler."
    exit 0
fi

log "The NSWorkspace request completed without changing the handler."
log "No Launch Services preference files were modified."
exit 2