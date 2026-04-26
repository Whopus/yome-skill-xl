-- xl sheet.rename <oldName> --name=<newName>
tell application "Microsoft Excel"
    set name of worksheet {{oldName|json}} of active workbook to {{name|json}}
    return {{name|json}}
end tell
