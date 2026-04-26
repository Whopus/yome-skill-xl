-- xl find <what> [--in=<range>] — returns TSV: address\tvalue
set what to {{what|json}}
set inRange to {{in|json}}
tell application "Microsoft Excel"
    tell active sheet
        if inRange is "" then
            set searchRange to used range
        else
            set searchRange to range inRange
        end if
        set results to {"address" & tab & "value"}
        try
            set found to find searchRange what what
            if found is not missing value then
                set firstAddr to get address of found
                set end of results to firstAddr & tab & ((value of found) as string)
                set lastFound to found
                repeat
                    set found to find next searchRange after lastFound
                    set nextAddr to get address of found
                    if nextAddr is firstAddr then exit repeat
                    set end of results to nextAddr & tab & ((value of found) as string)
                    set lastFound to found
                end repeat
            end if
        end try
        set AppleScript's text item delimiters to linefeed
        return results as string
    end tell
end tell
