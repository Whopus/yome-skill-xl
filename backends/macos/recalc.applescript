-- xl recalc [--scope=workbook|sheet|app]  (default workbook)
set scopeKind to {{scope|json}}
tell application "Microsoft Excel"
    if scopeKind is "sheet" then
        calculate active sheet
    else if scopeKind is "app" then
        calculate
    else
        calculate active workbook
    end if
end tell
return "recalc"
