-- xl export --format=pdf|csv --path=<path>
set fmt to {{format|json}}
set outPath to {{path|json}}
tell application "Microsoft Excel"
    if fmt is "pdf" then
        save active workbook in outPath as PDF file format
    else if fmt is "csv" then
        save active workbook in outPath as CSV file format
    else
        error "unsupported format: " & fmt
    end if
    return name of active workbook
end tell
