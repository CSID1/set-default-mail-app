#!/bin/bash

# Sets Microsoft Outlook as the default application for .eml files for the
# currently signed-in macOS user. This is equivalent to choosing Outlook under
# Finder > Get Info > Open with and then selecting Change All.
#
# Intended for execution by NinjaOne as root on macOS 12 or later.

PATH="/usr/bin:/bin:/usr/sbin:/sbin"

TARGET_APP="/Applications/Microsoft Outlook.app"
TARGET_BUNDLE_ID="com.microsoft.Outlook"

log() {
    echo "[DefaultEML] $*"
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

temporary_directory=$(/usr/bin/mktemp -d "/private/tmp/default-eml.XXXXXX")

if [[ -z "$temporary_directory" || ! -d "$temporary_directory" ]]; then
    log "Unable to create the temporary working directory."
    exit 1
fi

sample_eml="$temporary_directory/sample.eml"
query_source="$temporary_directory/query-handler.applescript"
applet_source="$temporary_directory/main.applescript"
applet_path="$temporary_directory/Default EML Setup.app"

cleanup() {
    if [[ -n "$temporary_directory" && -d "$temporary_directory" ]]; then
        /bin/rm -rf "$temporary_directory"
    fi
}

trap cleanup EXIT

# A valid but entirely fictitious email message gives macOS a real .eml file
# from which to determine the applicable content type.
/bin/cat > "$sample_eml" <<'EML'
From: sender@example.invalid
To: recipient@example.invalid
Subject: Default application test
Date: Thu, 1 Jan 1970 00:00:00 +0000
MIME-Version: 1.0
Content-Type: text/plain; charset=UTF-8

This temporary message is used only to identify the macOS content type for
.eml files. It contains no real email address or user information.
EML

/bin/cat > "$query_source" <<'APPLESCRIPT'
use framework "AppKit"
use framework "Foundation"

on run arguments
    set samplePath to item 1 of arguments
    set sampleURL to current application's NSURL's fileURLWithPath:samplePath
    set workspace to current application's NSWorkspace's sharedWorkspace()
    set applicationURL to workspace's URLForApplicationToOpenURL:sampleURL

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
end run
APPLESCRIPT

/usr/sbin/chown -R "$console_uid" "$temporary_directory"
/bin/chmod -R u+rwX,go+rX "$temporary_directory"

get_eml_handler() {
    run_as_console_user /usr/bin/osascript "$query_source" "$sample_eml"
}

current_handler=$(get_eml_handler)
query_status=$?

log "Signed-in user: $console_user"

if [[ "$query_status" -ne 0 ]]; then
    log "Unable to determine the current .eml handler."
    exit 1
fi

log "Current .eml handler: $current_handler"

if [[ "$current_handler" == "$TARGET_BUNDLE_ID" ]]; then
    log "Outlook is already the default application for .eml files."
    exit 0
fi

/bin/cat > "$applet_source" <<APPLESCRIPT
use framework "AppKit"
use framework "Foundation"
use scripting additions

on run
    set applicationObject to current application's NSApplication's sharedApplication()

    applicationObject's setActivationPolicy:(current application's NSApplicationActivationPolicyRegular)
    applicationObject's activateIgnoringOtherApps:true

    set workspace to current application's NSWorkspace's sharedWorkspace()
    set outlookURL to current application's NSURL's fileURLWithPath:"$TARGET_APP"
    set emailFileURL to current application's NSURL's fileURLWithPath:"$sample_eml"

    workspace's setDefaultApplicationAtURL:outlookURL ¬
        toOpenContentTypeOfFileAtURL:emailFileURL ¬
        completionHandler:(missing value)

    -- Keep the Cocoa event loop alive while Launch Services processes the
    -- asynchronous default-application request.
    set expirationDate to current application's NSDate's dateWithTimeIntervalSinceNow:15
    current application's NSRunLoop's currentRunLoop()'s runUntilDate:expirationDate
end run
APPLESCRIPT

log "Building temporary NSWorkspace helper."

compile_output=$(/usr/bin/osacompile -o "$applet_path" "$applet_source" 2>&1)
compile_status=$?

if [[ "$compile_status" -ne 0 || ! -d "$applet_path" ]]; then
    log "Unable to build the temporary helper."
    if [[ -n "$compile_output" ]]; then
        log "$compile_output"
    fi
    exit 1
fi

/usr/sbin/chown -R "$console_uid" "$applet_path"
/bin/chmod -R u+rwX,go+rX "$applet_path"

log "Requesting Outlook as the default application for .eml files."

run_as_console_user /usr/bin/open -W -n "$applet_path"
launch_status=$?

if [[ "$launch_status" -ne 0 ]]; then
    log "The temporary helper returned exit code $launch_status."
fi

current_handler=$(get_eml_handler)
query_status=$?

if [[ "$query_status" -ne 0 ]]; then
    log "Unable to verify the .eml handler after the request."
    exit 1
fi

log ".eml handler after the request: $current_handler"

if [[ "$current_handler" == "$TARGET_BUNDLE_ID" ]]; then
    log "Success: Outlook is now the default application for .eml files."
    exit 0
fi

log "The request completed without changing the .eml handler."
log "No Launch Services preference files were modified directly."
exit 2
