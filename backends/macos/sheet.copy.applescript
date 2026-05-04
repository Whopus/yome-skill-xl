-- xl sheet.copy <name> [--to=<existing_sheet>] [--position=before|after] [--newName=str]
set srcName to {{name|json}}
set toName to {{to|json}}
set posSpec to {{position|json}}
set newName to {{newName|json}}

tell application "Microsoft Excel"
    set wb to active workbook
    set srcWS to worksheet srcName of wb
    if toName is "" then
        copy worksheet srcWS after (worksheet (count of worksheets of wb) of wb)
    else
        set anchorWS to worksheet toName of wb
        if posSpec is "before" then
            copy worksheet srcWS before anchorWS
        else
            copy worksheet srcWS after anchorWS
        end if
    end if
    set newWS to active sheet
    if newName is not "" then
        try
            set name of newWS to newName
        end try
    end if
    return name of newWS
end tell
