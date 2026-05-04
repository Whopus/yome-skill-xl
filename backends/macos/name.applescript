-- xl name <name> --range=<range>  (use --range="" or no flag to delete)
set theName to {{name|json}}
set rngExpr to {{range|json}}

tell application "Microsoft Excel"
    if rngExpr is "" then
        try
            delete (name theName of active workbook)
        end try
        return "deleted"
    else
        try
            delete (name theName of active workbook)
        end try
        add name names of active workbook name string theName refers to ("=" & rngExpr)
        return "named"
    end if
end tell
