-- xl protect.book [--password=str] [--structure=true|false] [--windows=true|false]
set pwd to {{password|json}}
set sFlag to {{structure|json}}
set wFlag to {{windows|json}}
set protectStruct to true
set protectWin to false
if sFlag is "false" then set protectStruct to false
if wFlag is "true" then set protectWin to true

tell application "Microsoft Excel"
    if pwd is "" then
        protect active workbook structure protectStruct windows protectWin
    else
        protect active workbook structure protectStruct windows protectWin password pwd
    end if
end tell
return "protected"
