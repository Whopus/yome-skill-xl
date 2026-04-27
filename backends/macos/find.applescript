-- xl find <what> [--in=<range>] — returns TSV: address\tvalue
-- See fill.applescript: bare `tab` inside an Excel tell block is Excel's
-- terminology, not the TAB character.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

set what to {{what|json}}
set inRange to {{in|json}}
tell application "Microsoft Excel"
    tell active sheet
        if inRange is "" then
            set searchRange to used range
        else
            set searchRange to range inRange
        end if
        set results to {"address" & TAB_CHAR & "value"}
        try
            set found to find searchRange what what
            if found is not missing value then
                set firstAddr to get address of found
                set end of results to firstAddr & TAB_CHAR & ((value of found) as string)
                set lastFound to found
                repeat
                    set found to find next searchRange after lastFound
                    set nextAddr to get address of found
                    if nextAddr is firstAddr then exit repeat
                    set end of results to nextAddr & TAB_CHAR & ((value of found) as string)
                    set lastFound to found
                end repeat
            end if
        end try
        set AppleScript's text item delimiters to LF_CHAR
        return results as string
    end tell
end tell
