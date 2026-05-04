-- xl sheet.move <name> --to=<existing_sheet> [--position=before|after]
set srcName to {{name|json}}
set toName to {{to|json}}
set posSpec to {{position|json}}

tell application "Microsoft Excel"
    set wb to active workbook
    set srcWS to worksheet srcName of wb
    set anchorWS to worksheet toName of wb
    if posSpec is "before" then
        move worksheet srcWS before anchorWS
    else
        move worksheet srcWS after anchorWS
    end if
end tell
return "moved"
