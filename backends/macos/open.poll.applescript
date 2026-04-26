-- Polled by the dispatcher after Launch Services opens the file.
-- Returns the active workbook name once Excel reports >=1 open book.
tell application "Microsoft Excel"
    if (count of workbooks) > 0 then
        return name of active workbook
    end if
    return ""
end tell
