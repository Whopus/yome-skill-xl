-- xl sheet.protect [--name=<sheet>] [--password=str]
set sName to {{name|json}}
set pwd to {{password|json}}

tell application "Microsoft Excel"
    if sName is "" then
        set targetWS to active sheet
    else
        set targetWS to worksheet sName of active workbook
    end if
    if pwd is "" then
        protect targetWS
    else
        protect targetWS password pwd
    end if
end tell
return "protected"
