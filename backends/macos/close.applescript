-- xl close [--save=true|false]
set shouldSave to {{save|bool}}
tell application "Microsoft Excel"
    if shouldSave then
        close active workbook saving yes
    else
        close active workbook saving no
    end if
end tell
return "closed"
