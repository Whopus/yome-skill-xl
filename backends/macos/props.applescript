-- xl props [--title=str] [--author=str] [--subject=str] [--keywords=str] [--comments=str]
-- With no flags, returns the current values as TSV.
set TAB_CHAR to (ASCII character 9)
set LF_CHAR to (ASCII character 10)
set tTitle to {{title|json}}
set tAuthor to {{author|json}}
set tSubject to {{subject|json}}
set tKeywords to {{keywords|json}}
set tComments to {{comments|json}}
set anySet to (tTitle is not "") or (tAuthor is not "") or (tSubject is not "") or (tKeywords is not "") or (tComments is not "")

tell application "Microsoft Excel"
    set wb to active workbook
    if anySet then
        if tTitle is not "" then set value of built-in document property "Title" of wb to tTitle
        if tAuthor is not "" then set value of built-in document property "Author" of wb to tAuthor
        if tSubject is not "" then set value of built-in document property "Subject" of wb to tSubject
        if tKeywords is not "" then set value of built-in document property "Keywords" of wb to tKeywords
        if tComments is not "" then set value of built-in document property "Comments" of wb to tComments
        return "updated"
    else
        set rTitle to ""
        set rAuthor to ""
        set rSubject to ""
        set rKeywords to ""
        set rComments to ""
        try
            set rTitle to (value of built-in document property "Title" of wb) as string
        end try
        try
            set rAuthor to (value of built-in document property "Author" of wb) as string
        end try
        try
            set rSubject to (value of built-in document property "Subject" of wb) as string
        end try
        try
            set rKeywords to (value of built-in document property "Keywords" of wb) as string
        end try
        try
            set rComments to (value of built-in document property "Comments" of wb) as string
        end try
        return "title" & TAB_CHAR & rTitle & LF_CHAR & ¬
            "author" & TAB_CHAR & rAuthor & LF_CHAR & ¬
            "subject" & TAB_CHAR & rSubject & LF_CHAR & ¬
            "keywords" & TAB_CHAR & rKeywords & LF_CHAR & ¬
            "comments" & TAB_CHAR & rComments
    end if
end tell
