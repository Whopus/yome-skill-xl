-- xl print [--copies=1] [--scope=workbook|sheet]
set copyCount to ({{copies|json}}) as integer
set scopeKind to {{scope|json}}
if copyCount < 1 then set copyCount to 1

tell application "Microsoft Excel"
    if scopeKind is "workbook" then
        print out active workbook number of copies copyCount
    else
        print out active sheet number of copies copyCount
    end if
end tell
return "printed"
