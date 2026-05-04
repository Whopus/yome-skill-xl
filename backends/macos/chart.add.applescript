-- xl chart.add --range=<source> [--type=column|bar|line|pie|area|scatter] [--title=str]
--   [--left=N] [--top=N] [--width=N] [--height=N] [--dest=A1]
set rawRef to {{range|json}}
set chartKind to {{type|json}}
set chartTitle to {{title|json}}
set leftPos to ({{left|json}}) as real
set topPos to ({{top|json}}) as real
set widthVal to ({{width|json}}) as real
set heightVal to ({{height|json}}) as real
set destAddr to {{dest|json}}

if leftPos = 0 then set leftPos to 200
if topPos = 0 then set topPos to 50
if widthVal = 0 then set widthVal to 400
if heightVal = 0 then set heightVal to 250

set sheetName to ""
set addr to rawRef
if rawRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set parts to text items of rawRef
    set sheetName to item 1 of parts
    set addr to item 2 of parts
    set AppleScript's text item delimiters to ""
end if

set ptType to column clustered
if chartKind is "bar" then set ptType to bar clustered
if chartKind is "line" then set ptType to line
if chartKind is "pie" then set ptType to pie
if chartKind is "area" then set ptType to area
if chartKind is "scatter" then set ptType to XY scatter

tell application "Microsoft Excel"
    if sheetName is "" then
        set targetSheet to active sheet
    else
        set targetSheet to worksheet sheetName of active workbook
    end if
    tell targetSheet
        set srcRange to range addr
        set newCO to make new chart object at end with properties {left position:leftPos, top:topPos, width:widthVal, height:heightVal}
        set theChart to chart of newCO
        set chart type of theChart to ptType
        set source data of theChart source srcRange
        if chartTitle is not "" then
            try
                set has title of theChart to true
                set caption of chart title of theChart to chartTitle
            end try
        end if
        if destAddr is not "" then
            try
                set destCell to range destAddr
                set left position of newCO to left position of destCell
                set top of newCO to top of destCell
            end try
        end if
        return name of newCO
    end tell
end tell
