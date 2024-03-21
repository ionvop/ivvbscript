option explicit
dim objShell, objFile
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
dim directory
directory = objFile.GetParentFolderName(wscript.ScriptFullName)

sub Main()
    dim i, element, content, scope, token, result, temp, temp2, statement, funcScope, interpolate, keyword, context
    content = objFile.OpenTextFile(wscript.Arguments(0)).ReadAll()
    content = replace(content, vbcrlf, "")
    scope = "code"
    token = "start"
    result = objFile.OpenTextFile(directory & "\initial.vbs").ReadAll()
    objFile.CreateTextFile(directory & "\output.vbs", true).Write(result)
    statement = ""
    funcScope = ""
    i = 1
    
    do
        if i > len(content) then
            exit do
        end if

        element = mid(content, i, 1)

        if scope = "code" then
            select case element
                case """"
                    scope = "string"
                    statement = statement & """"
                    i = i + 1
                    element = mid(content, i, 1)
                case "`"
                    scope = "literal"
                    statement = statement & """"
                    i = i + 1
                    element = mid(content, i, 1)
                case "$"
                    if mid(content, i + 1, 1) = "^" then
                        statement = statement & "ivvbsobjshell.expandenvironmentstrings"
                        i = i + 2
                        element = mid(content, i, 1)
                    else
                        statement = statement & "ivvbsobjshell."
                        i = i + 1
                        element = mid(content, i, 1)
                    end if
                case "#"
                    statement = statement & "ivvbsobjfile."
                    i = i + 1
                    element = mid(content, i, 1)
                case "%"
                    if mid(content, i + 1, 1) = "%" then
                        statement = statement & "ivvbsobjhttp."
                        i = i + 2
                        element = mid(content, i, 1)
                    else
                        statement = statement & " mod "
                        i = i + 1
                        element = mid(content, i, 1)
                    end if
                case "^"
                    if mid(content, i + 1, 1) = "^" then
                        statement = statement & "wscript.arguments"
                        i = i + 2
                    else
                        statement = statement & "wscript."
                        i = i + 1
                    end if

                    element = mid(content, i, 1)
                case "&"
                    if mid(content, i + 1, 1) = "&" then
                        statement = statement & " and "
                        i = i + 2
                        element = mid(content, i, 1)
                    end if
                case "|"
                    if mid(content, i + 1, 1) = "|" then
                        statement = statement & " or "
                        i = i + 2
                        element = mid(content, i, 1)
                    end if
                case "["
                    if mid(content, i + 1, 1) = "]" then
                        statement = statement & "array"
                        i = i + 2
                        element = mid(content, i, 1)
                    end if
            end select
        end if

        if scope = "code" then
            if element = "{" then
                context = mid(content, i - 1, 1)
                keyword = mid(content, i + 1)
                i = i + instr(keyword, "}") + 1
                keyword = left(keyword, instr(keyword, "}") - 1)

                select case context
                    case "^"
                        select case keyword
                            case "a"
                                statement = statement & "agruments"
                            case "bv"
                                statement = statement & "buildversion"
                            case "fn"
                                statement = statement & "fullname"
                            case "i"
                                statement = statement & "interactive"
                            case "n"
                                statement = statement & "name"
                            case "p"
                                statement = statement & "path"
                            case "sfn"
                                statement = statement & "scriptfullname"
                            case "sn"
                                statement = statement & "scriptname"
                            case "se"
                                statement = statement & "stderr"
                            case "si"
                                statement = statement & "stdin"
                            case "so"
                                statement = statement & "stdout"
                            case "v"
                                statement = statement & "version"
                            case "co"
                                statement = statement & "connectobject"
                            case "cro"
                                statement = statement & "createobject"
                            case "do"
                                statement = statement & "disconnectobject"
                            case "e"
                                statement = statement & "echo"
                            case "go"
                                statement = statement & "getobject"
                            case "q"
                                statement = statement & "quit"
                            case "s"
                                statement = statement & "sleep"
                        end select
                    case "$"
                        select case keyword
                            case "cd"
                                statement = statement & "currentdirectory"
                            case "ev"
                                statement = statement & "environment"
                            case "sf"
                                statement = statement & "specialfolders"
                            case "aa"
                                statement = statement & "appactivate"
                            case "cs"
                                statement = statement & "createshortcut"
                            case "e"
                                statement = statement & "exec"
                            case "ees"
                                statement = statement & "expandenvironmentstrings"
                            case "le"
                                statement = statement & "logevent"
                            case "p"
                                statement = statement & "popup"
                            case "rd"
                                statement = statement & "regdelete"
                            case "rr"
                                statement = statement & "regread"
                            case "rw"
                                statement = statement & "regwrite"
                            case "r"
                                statement = statement & "run"
                            case "sk"
                                statement = statement & "sendkeys"
                        end select
                    case "#"
                        select case keyword
                            case "bp"
                                statement = statement & "buildpath"
                            case "cf"
                                statement = statement & "copyfile"
                            case "cfd"
                                statement = statement & "copyfolder"
                            case "crf"
                                statement = statement & "createfolder"
                            case "ctf"
                                statement = statement & "createtextfile"
                            case "df"
                                statement = statement & "deletefile"
                            case "dfd"
                                statement = statement & "deletefolder"
                            case "de"
                                statement = statement & "driveexists"
                            case "fe"
                                statement = statement & "fileexists"
                            case "fde"
                                statement = statement & "folderexists"
                            case "gapn"
                                statement = statement & "getabsolutepathname"
                            case "gbn"
                                statement = statement & "getbasename"
                            case "gd"
                                statement = statement & "getdrive"
                            case "gdn"
                                statement = statement & "getdrivename"
                            case "gen"
                                statement = statement & "getextensionname"
                            case "gf"
                                statement = statement & "getfile"
                            case "gfn"
                                statement = statement & "getfilename"
                            case "gfd"
                                statement = statement & "getfolder"
                            case "gpfn"
                                statement = statement & "getparentfoldername"
                            case "gsfn"
                                statement = statement & "getspecialfoldername"
                            case "gss"
                                statement = statement & "getstandardstream"
                            case "gtn"
                                statement = statement & "gettempname"
                            case "mf"
                                statement = statement & "movefile"
                            case "mfd"
                                statement = statement & "movefolder"
                            case "otf"
                                statement = statement & "opentextfile"
                        end select
                    case "%"
                        select case keyword
                            case "orsc"
                                statement = statement & "onreadystatechange"
                            case "rs"
                                statement = statement & "readystate"
                            case "rb"
                                statement = statement & "responsebody"
                            case "rst"
                                statement = statement & "responsestream"
                            case "rt"
                                statement = statement & "responsetext"
                            case "rx"
                                statement = statement & "responsexml"
                            case "st"
                                statement = statement & "status"
                            case "stt"
                                statement = statement & "statustext"
                            case "a"
                                statement = statement & "abort"
                            case "garh"
                                statement = statement & "getallresponseheaders"
                            case "grh"
                                statement = statement & "getresponseheader"
                            case "o"
                                statement = statement & "open"
                            case "s"
                                statement = statement & "send"
                            case "srh"
                                statement = statement & "setrequestheader"
                        end select
                    case else
                        select case keyword
                            case "#ra"
                                statement = statement & "readall"
                            case "#rl"
                                statement = statement & "readline"
                            case "#aeos"
                                statement = statement & "atendofstream"
                            case "#c"
                                statement = statement & "close"
                            case "#w"
                                statement = statement & "write"
                            case "a"
                                statement = statement & "abs"
                            case "ar"
                                statement = statement & "array"
                            case "as"
                                statement = statement & "asc"
                            case "asb"
                                statement = statement & "ascb"
                            case "asw"
                                statement = statement & "ascw"
                            case "cb"
                                statement = statement & "cbool"
                            case "cby"
                                statement = statement & "cbyte"
                            case "cc"
                                statement = statement & "ccur"
                            case "cd"
                                statement = statement & "cdate"
                            case "cdb"
                                statement = statement & "cdbl"
                            case "c"
                                statement = statement & "chr"
                            case "chb"
                                statement = statement & "chrb"
                            case "chw"
                                statement = statement & "chrw"
                            case "ci"
                                statement = statement & "cint"
                            case "cl"
                                statement = statement & "clng"
                            case "cs"
                                statement = statement & "csng"
                            case "cst"
                                statement = statement & "cstr"
                            case "d"
                                statement = statement & "date"
                            case "da"
                                statement = statement & "day"
                            case "e"
                                statement = statement & "escape"
                            case "ev"
                                statement = statement & "eval"
                            case "ex"
                                statement = statement & "exp"
                            case "f"
                                statement = statement & "filter"
                            case "fi"
                                statement = statement & "fix"
                            case "gl"
                                statement = statement & "getlocale"
                            case "h"
                                statement = statement & "hex"
                            case "ho"
                                statement = statement & "hour"
                            case "ib"
                                statement = statement & "inputbox"
                            case "is"
                                statement = statement & "instr"
                            case "isr"
                                statement = statement & "instrrev"
                            case "i"
                                statement = statement & "int"
                            case "ia"
                                statement = statement & "isarray"
                            case "id"
                                statement = statement & "isdate"
                            case "ie"
                                statement = statement & "isempty"
                            case "in"
                                statement = statement & "isnull"
                            case "inu"
                                statement = statement & "isnumeric"
                            case "io"
                                statement = statement & "isobject"
                            case "j"
                                statement = statement & "join"
                            case "lb"
                                statement = statement & "lbound"
                            case "lc"
                                statement = statement & "lcase"
                            case "l"
                                statement = statement & "left"
                            case "leb"
                                statement = statement & "leftb"
                            case "le"
                                statement = statement & "len"
                            case "leb"
                                statement = statement & "lenb"
                            case "lo"
                                statement = statement & "log"
                            case "lt"
                                statement = statement & "ltrim"
                            case "m"
                                statement = statement & "mid"
                            case "mi"
                                statement = statement & "minute"
                            case "mo"
                                statement = statement & "month"
                            case "mb"
                                statement = statement & "msgbox"
                            case "n"
                                statement = statement & "now"
                            case "ra"
                                statement = statement & "randomize"
                            case "re"
                                statement = statement & "replace"
                            case "r"
                                statement = statement & "rgb"
                            case "r"
                                statement = statement & "right"
                            case "rb"
                                statement = statement & "rightb"
                            case "rn"
                                statement = statement & "rnd"
                            case "ro"
                                statement = statement & "round"
                            case "rt"
                                statement = statement & "rtrim"
                            case "se"
                                statement = statement & "second"
                            case "sl"
                                statement = statement & "setlocale"
                            case "sg"
                                statement = statement & "sgn"
                            case "s"
                                statement = statement & "split"
                            case "sq"
                                statement = statement & "sqr"
                            case "sr"
                                statement = statement & "strreverse"
                            case "ti"
                                statement = statement & "time"
                            case "tim"
                                statement = statement & "timer"
                            case "t"
                                statement = statement & "trim"
                            case "tn"
                                statement = statement & "typename"
                            case "ub"
                                statement = statement & "ubound"
                            case "uc"
                                statement = statement & "ucase"
                            case "ue"
                                statement = statement & "unescape"
                            case "y"
                                statement = statement & "year"
                            case "+"
                                statement = statement & "true"
                            case "-"
                                statement = statement & "false"
                        end select
                end select

                element = mid(content, i, 1)
            end if
        end if
        
        select case scope
            case "code"
                select case token
                    case "start"
                        select case element
                            case " ", vbtab, vbcrlf
                            case "."
                                token = "print"
                                statement = ""
                            case ","
                                token = "input"
                                statement = ""
                            case "?"
                                if mid(content, i + 1, 1) = ";" then
                                    result = result & "end if" & vbcrlf
                                    i = i + 1
                                else
                                    token = "if"
                                end if

                                statement = ""
                            case ":"
                                if mid(content, i + 1, 1) = "?" then
                                    token = "elseif"
                                    i = i + 1
                                else
                                    result = result & "else" & vbcrlf
                                end if

                                statement = ""
                            case ">"
                                token = "dim"
                                statement = ""
                            case "<"
                                token = "return"
                                statement = ""
                            case "+", "-", "*", "/", "&"
                                token = "operate"
                                temp = element
                                statement = ""
                            case "["
                                result = result & "do" & vbcrlf
                                statement = ""
                            case "]"
                                result = result & "loop" & vbcrlf
                                statement = ""
                            case "\"
                                result = result & "exit do" & vbcrlf
                                statement = ""
                            case "!"
                                token = "exit"
                                statement = ""
                            case "~"
                                token = "sleep"
                                statement = ""
                            case "_"
                                token = "normal"
                                statement = "call "
                            case "@"
                                token = "import"
                                statement = ""
                            case else
                                token = "normal"
                                statement = statement & element
                        end select
                    case "normal"
                        select case element
                            case ";"
                                token = "start"
                                result = result & statement & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "print"
                        select case element
                            case "."
                                if statement = "" then
                                    token = "print2"
                                else
                                    statement = statement & element
                                end if
                            case ";"
                                token = "start"
                                result = result & "ivvbsprint(" & statement & ")" & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "print2"
                        select case element
                            case ";"
                                token = "start"
                                result = result & "wscript.stdout.write(" & statement & ")" & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "input"
                        select case element
                            case ":"
                                token = "input2"
                                temp = statement
                                statement = ""
                            case ";"
                                token = "start"
                                result = result & statement & " = wscript.stdin.readline()" & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "input2"
                        select case element
                            case ";"
                                token = "start"
                                result = result & "wscript.stdout.write(" & statement & ")" & vbcrlf
                                result = result & temp & " = wscript.stdin.readline()" & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "if"
                        select case element
                            case ":"
                                token = "start"
                                result = result & "if " & statement & " then" & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "elseif"
                        select case element
                            case ":"
                                token = "start"
                                result = result & "elseif " & statement & " then" & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "dim"
                        select case element
                            case ">"
                                if statement = "" then
                                    token = "redim"
                                    statement = ""
                                else
                                    statement = statement & element
                                end if
                            case "("
                                token = "function"
                                temp = statement
                                statement = ""
                            case ";"
                                token = "start"
                                result = result & "dim " & statement & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "redim"
                        select case element
                            case ";"
                                token = "start"
                                result = result & "redim preserve " & statement & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "function"
                        select case element
                            case ")"
                                token = "start"
                                result = result & "function " & temp & "(" & statement & ")" & vbcrlf
                                funcScope = temp
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "return"
                        select case element
                            case ";"
                                token = "start"
                                result = result & funcScope & " = " & statement & vbcrlf
                                result = result & "end function" & vbcrlf
                                funcScope = ""
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "operate"
                        select case element
                            case ":"
                                token = "operate2"
                                temp2 = statement
                                statement = ""
                            case ";"
                                token = "start"
                                
                                select case temp
                                    case "+", "-"
                                        result = result & statement & " = " & statement & " " & temp & " 1" & vbcrlf
                                    case "*", "/"
                                        result = result & statement & " = " & statement & " " & temp & " 2" & vbcrlf
                                    case "&"
                                        result = result & statement & " = " & statement & " " & temp & " vbcrlf" & vbcrlf
                                end select

                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "operate2"
                        select case element
                            case ";"
                                token = "start"
                                result = result & temp2 & " = " & temp2 & " " & temp & " " & statement & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "exit"
                        select case element
                            case ";"
                                token = "start"

                                if statement = "" then
                                    result = result & "wscript.quit()" & vbcrlf
                                else
                                    result = result & "ivvbsprint(" & statement & ")" & vbcrlf
                                    result = result & "wscript.quit()" & vbcrlf
                                end if

                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "sleep"
                        select case element
                            case ";"
                                token = "start"
                                result = result & "wscript.sleep(" & statement & ")" & vbcrlf
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                    case "import"
                        select case element
                            case ";"
                                token = "start"
                                content = content & objFile.OpenTextFile(objFile.GetParentFolderName(objFile.GetAbsolutePathName(wscript.arguments(0))) & "/" & statement & ".ivvbs").ReadAll()
                                statement = ""
                            case else
                                statement = statement & element
                        end select
                end select
            case "string"
                select case element
                    case """"
                        scope = "code"
                        statement = statement & """"
                    case else
                        statement = statement & element
                end select
            case "literal"
                select case element
                    case vbcrlf
                    case """"
                        statement = statement & """"""
                    case "$"
                        if mid(content, i + 1, 1) = "{" then
                            interpolate = true
                            i = i + 1
                            statement = statement & """ & "
                        else
                            statement = statement & element
                        end if
                    case "}"
                        if interpolate = true then
                            interpolate = false
                            statement = statement & " & """
                        else
                            statement = statement & element
                        end if
                    case "\"
                        i = i + 1

                        select case mid(content, i, 1)
                            case "n"
                                statement = statement & """ & vbcrlf & """
                            case "t"
                                statement = statement & """ & vbtab & """
                            case "`"
                                statement = statement & "`"
                            case "$"
                                statement = statement & "$"
                            case "\"
                                statement = statement & "\"
                            case else
                                statement = statement & element
                        end select
                    case "`"
                        scope = "code"
                        statement = statement & """"
                    case else
                        statement = statement & element
                end select
        end select

        i = i + 1
    loop

    objFile.CreateTextFile(directory & "\output.vbs", true).Write(result)
end sub

sub Breakpoint(message)
    wscript.Echo(message)
    wscript.Quit()
end sub

Main()
