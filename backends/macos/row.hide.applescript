-- xl row.hide <row> [--count=1] [--show=true|false]
set startRow to ({{row|json}}) as integer
set howMany to ({{count|json}}) as integer
set showFlag to {{show|json}}
set hide to true
if showFlag is "true" then set hide to false

tell application "Microsoft Excel"
    tell active sheet
        set endRow to startRow + howMany - 1
        set targetRange to range (startRow & ":" & endRow as string)
        set hidden of targetRange to hide
    end tell
end tell
if hide then
    return "hid"
else
    return "shown"
end if
