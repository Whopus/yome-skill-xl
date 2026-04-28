-- xl find <what> [--in=<range>] — returns TSV: address\tvalue
-- See fill.applescript: bare `tab` inside an Excel tell block is Excel's
-- terminology, not the TAB character.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)

-- NOTE: do NOT name the local var `what` — Excel's `find` command takes
-- a `what` parameter, and AppleScript will mis-parse `find … what what`.
set whatStr to {{what|json}}
set inRange to {{in|json}}
-- Collect raw addr/value pairs INSIDE the Excel tell as plain strings,
-- then build the TSV outside (text item delimiters collides with Excel's
-- terminology -> -50 if used inside the tell block).
set rawAddrs to {}
set rawVals to {}
tell application "Microsoft Excel"
    tell active sheet
        if inRange is "" then
            set searchRange to used range
        else
            set searchRange to range inRange
        end if
        try
            set found to find searchRange what whatStr
            if found is not missing value then
                set firstAddr to (get address of found) as string
                set end of rawAddrs to firstAddr
                set end of rawVals to ((value of found) as string)
                set lastFound to found
                repeat
                    set found to find next searchRange after lastFound
                    set nextAddr to (get address of found) as string
                    if nextAddr is firstAddr then exit repeat
                    set end of rawAddrs to nextAddr
                    set end of rawVals to ((value of found) as string)
                    set lastFound to found
                end repeat
            end if
        end try
    end tell
end tell
set results to {"address" & TAB_CHAR & "value"}
repeat with i from 1 to count of rawAddrs
    set end of results to (item i of rawAddrs) & TAB_CHAR & (item i of rawVals)
end repeat
-- Join OUTSIDE the Excel tell block: inside it, `text item delimiters`
-- collides with Excel's terminology and raises (-50) 参数错误.
set AppleScript's text item delimiters to LF_CHAR
set joined to results as string
set AppleScript's text item delimiters to ""
return joined
