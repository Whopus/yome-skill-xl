-- xl sheet.add --name=<name>
tell application "Microsoft Excel"
    make new worksheet at end of active workbook with properties {name:{{name|json}}}
    return name of active sheet
end tell
