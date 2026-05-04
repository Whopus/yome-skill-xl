-- xl cf <range> --kind=cellvalue|colorscale|databar|iconset
--   cellvalue: --op=greater|less|equal|between --v1=<n> [--v2=<n>] [--bg=color] [--color=color]
--   colorscale: [--low=color] [--mid=color] [--high=color]
--   databar:    [--color=color]
--   iconset:    [--style=3arrows|3traffic|3symbols|4arrows|5arrows]
set rawRef to {{range|json}}
set kindSpec to {{kind|json}}
set opSpec to {{op|json}}
set v1 to {{v1|json}}
set v2 to {{v2|json}}
set bgColor to {{bg|rgb}}
set fgColor to {{color|rgb}}
set lowColor to {{low|rgb}}
set midColor to {{mid|rgb}}
set highColor to {{high|rgb}}
set iconStyle to {{style|json}}

set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set targetRange to range addr
        if kindSpec is "cellvalue" then
            set opVal to greater
            if opSpec is "less" then set opVal to less
            if opSpec is "equal" then set opVal to equal
            if opSpec is "between" then set opVal to between
            if opSpec is "between" then
                set fc to make new format condition at end of (format conditions of targetRange) with properties {type:cell value, condition operator:opVal, formula1:v1, formula2:v2}
            else
                set fc to make new format condition at end of (format conditions of targetRange) with properties {type:cell value, condition operator:opVal, formula1:v1}
            end if
            try
                if bgColor is not "" then set color of interior object of fc to bgColor
            end try
            try
                if fgColor is not "" then set color of font object of fc to fgColor
            end try
        else if kindSpec is "colorscale" then
            try
                add color scale targetRange color scale type 3
            end try
        else if kindSpec is "databar" then
            try
                add databar targetRange
            end try
        else if kindSpec is "iconset" then
            try
                add icon set condition targetRange
            end try
        end if
    end tell
end tell
return "cf"
