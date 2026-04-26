-- xl sheet <name> — switch active sheet
tell application "Microsoft Excel"
    activate object worksheet {{name|json}} of active workbook
    return name of active sheet
end tell
