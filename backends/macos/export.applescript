-- xl export --format=pdf|csv --path=<path>
set fmt to {{format|json}}
set outPath to {{path|json}}
tell application "Microsoft Excel"
    set display alerts to false
    try
        if fmt is "pdf" then
            save active workbook in outPath as PDF file format
        else if fmt is "csv" then
            save active workbook in outPath as CSV file format
        else
            set display alerts to true
            error "unsupported format: " & fmt
        end if
        set display alerts to true
    on error errMsg number errNum
        set display alerts to true
        error errMsg number errNum
    end try
    return name of active workbook
end tell
