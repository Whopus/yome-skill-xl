-- xl image.add --path=<image> [--cell=A1 | --left=N --top=N] [--width=N] [--height=N]
set imgPath to {{path|json}}
set cellAddr to {{cell|json}}
set leftPos to ({{left|json}}) as real
set topPos to ({{top|json}}) as real
set widthVal to ({{width|json}}) as real
set heightVal to ({{height|json}}) as real
if widthVal = 0 then set widthVal to -1
if heightVal = 0 then set heightVal to -1

tell application "Microsoft Excel"
    tell active sheet
        if cellAddr is not "" then
            set anchorCell to range cellAddr
            set leftPos to left position of anchorCell
            set topPos to top of anchorCell
        end if
        set newPic to add picture filename imgPath link to file false save with document true left position leftPos top topPos width widthVal height heightVal
        return name of newPic
    end tell
end tell
