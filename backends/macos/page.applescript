-- xl page [--orientation=portrait|landscape] [--paper=letter|legal|a4|a3] [--fit=W,H] [--header=str] [--footer=str] [--margin=N or T,B,L,R]
set ori to {{orientation|json}}
set paperKind to {{paper|json}}
set fitSpec to {{fit|json}}
set hdr to {{header|json}}
set ftr to {{footer|json}}
set marginSpec to {{margin|json}}

set fitW to 0
set fitH to 0
if fitSpec is not "" then
    set AppleScript's text item delimiters to ","
    set fitParts to text items of fitSpec
    set AppleScript's text item delimiters to ""
    if (count of fitParts) >= 1 then set fitW to ((item 1 of fitParts) as integer)
    if (count of fitParts) >= 2 then set fitH to ((item 2 of fitParts) as integer)
end if

set mT to -1
set mB to -1
set mL to -1
set mR to -1
if marginSpec is not "" then
    set AppleScript's text item delimiters to ","
    set mParts to text items of marginSpec
    set AppleScript's text item delimiters to ""
    if (count of mParts) is 1 then
        set mT to ((item 1 of mParts) as real)
        set mB to mT
        set mL to mT
        set mR to mT
    else
        if (count of mParts) >= 1 then set mT to ((item 1 of mParts) as real)
        if (count of mParts) >= 2 then set mB to ((item 2 of mParts) as real)
        if (count of mParts) >= 3 then set mL to ((item 3 of mParts) as real)
        if (count of mParts) >= 4 then set mR to ((item 4 of mParts) as real)
    end if
end if

tell application "Microsoft Excel"
    tell page setup object of active sheet
        if ori is "landscape" then set orientation to landscape
        if ori is "portrait" then set orientation to portrait
        if paperKind is "letter" then set paper size to paper letter
        if paperKind is "legal" then set paper size to paper legal
        if paperKind is "a4" then set paper size to paper A4
        if paperKind is "a3" then set paper size to paper A3
        if fitW > 0 then set fit to pages wide to fitW
        if fitH > 0 then set fit to pages tall to fitH
        if hdr is not "" then set center header to hdr
        if ftr is not "" then set center footer to ftr
        if mT >= 0 then set top margin to mT
        if mB >= 0 then set bottom margin to mB
        if mL >= 0 then set left margin to mL
        if mR >= 0 then set right margin to mR
    end tell
end tell
return "page"
