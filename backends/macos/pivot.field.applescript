-- xl pivot.field --pt=<name> --field=<colName> --area=row|column|page|data [--func=sum|count|avg|min|max]
set ptName to {{pt|json}}
set fieldName to {{field|json}}
set areaKind to {{area|json}}
set funcKind to {{func|json}}

set orient to row field
if areaKind is "column" then set orient to column field
if areaKind is "page" then set orient to page field
if areaKind is "data" then set orient to data field

set fn to consolidation sum
if funcKind is "count" then set fn to consolidation count
if funcKind is "avg" or funcKind is "average" then set fn to consolidation average
if funcKind is "min" then set fn to consolidation min
if funcKind is "max" then set fn to consolidation max

tell application "Microsoft Excel"
    tell active sheet
        set pt to pivot table ptName
        set pf to pivot field fieldName of pt
        set orientation of pf to orient
        if areaKind is "data" then
            try
                set function of pf to fn
            end try
        end if
    end tell
end tell
return "field"
