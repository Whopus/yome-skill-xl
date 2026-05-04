-- xl col.hide <column> [--count=1] [--show=true|false]
set colLetter to {{column|json}}
set howMany to ({{count|json}}) as integer
set showFlag to {{show|json}}
set hide to true
if showFlag is "true" then set hide to false

tell application "Microsoft Excel"
    tell active sheet
        if howMany is 1 then
            set targetRange to range (colLetter & ":" & colLetter)
        else
            set startCol to first column index of range (colLetter & "1")
            set endColIdx to startCol + howMany - 1
            set targetRange to range ((colLetter & ":" & (address of cell 1 of column endColIdx)) as string)
        end if
        set hidden of targetRange to hide
    end tell
end tell
if hide then
    return "hid"
else
    return "shown"
end if
