option explicit
dim objShell, objFile
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
dim directory
directory = objFile.GetParentFolderName(wscript.ScriptFullName)

sub Main()
    dim i, element, content, scope, token, result, temp, temp2, statement, funcScope, interpolate, keyword
    content = objFile.OpenTextFile(wscript.Arguments(0)).ReadAll()
    content = replace(content, vbcrlf, "")
    scope = "code"
    token = "start"
    result = objFile.OpenTextFile(directory & "\initial.vbs").ReadAll()
    statement = ""
    funcScope = ""
    
    for i = 1 to len(content)
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
                    if instr("0123456789", mid(content, i + 1, 1)) <> 0 and instr("0123456789", mid(content, i + 2, 1)) <> 0 then
                        statement = statement & "wscript.arguments(" & mid(content, i + 1, 2) & ")"
                        i = i + 3
                    else
                        statement = statement & "wscript.arguments(" & mid(content, i + 1, 1) & ")"
                        i = i + 2
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
                keyword = mid(content, i + 1)
                i = i + instr(keyword, "}") + 1
                keyword = left(keyword, instr(keyword, "}") - 1)

                select case keyword
                    case "r"
                        statement = statement & "run"
                    case "e"
                        statement = statement & "exec"
                    case "sk"
                        statement = statement & "sendkeys"
                    case "bp"
                        statement = statement & "buildpath"
                    case "cf"
                        statement = statement & "copyfile"
                    case "cfd"
                        statement = statement & "copyfolder"
                    case "crfd"
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
                    case "ra"
                        statement = statement & "readall"
                    case "rl"
                        statement = statement & "readline"
                    case "aeos"
                        statement = statement & "atendofstream"
                    case "c"
                        statement = statement & "close"
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
    next

    objFile.CreateTextFile(directory & "\output.vbs", true).Write(result)
end sub

sub Breakpoint(message)
    wscript.Echo(message)
    wscript.Quit()
end sub

Main()
