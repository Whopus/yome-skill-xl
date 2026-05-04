-- xl pivot.add --src=<range> --dest=<sheet@cell or cell> [--name=<pt_name>]
set srcRef to {{src|json}}
set destRef to {{dest|json}}
set ptName to {{name|json}}

set srcSheet to ""
set srcAddr to srcRef
if srcRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set sp to text items of srcRef
    set srcSheet to item 1 of sp
    set srcAddr to item 2 of sp
    set AppleScript's text item delimiters to ""
end if
set dstSheet to ""
set dstAddr to destRef
if destRef contains "@" then
    set AppleScript's text item delimiters to "@"
    set dp to text items of destRef
    set dstSheet to item 1 of dp
    set dstAddr to item 2 of dp
    set AppleScript's text item delimiters to ""
end if

tell application "Microsoft Excel"
    if srcSheet is "" then
        set sourceSheet to active sheet
    else
        set sourceSheet to worksheet srcSheet of active workbook
    end if
    if dstSheet is "" then
        set destinationSheet to active sheet
    else
        set destinationSheet to worksheet dstSheet of active workbook
    end if
    set sRange to range srcAddr of sourceSheet
    set dRange to range dstAddr of destinationSheet
    set pc to make new pivot cache of active workbook with properties {source data:sRange}
    if ptName is "" then
        set newPT to create pivot table pivot cache pc table destination dRange
    else
        set newPT to create pivot table pivot cache pc table destination dRange table name ptName
    end if
    return name of newPT
end tell
