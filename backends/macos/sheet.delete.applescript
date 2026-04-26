-- xl sheet.delete <name>
tell application "Microsoft Excel"
    set display alerts to false
    delete worksheet {{name|json}} of active workbook
    set display alerts to true
end tell
return "deleted"
