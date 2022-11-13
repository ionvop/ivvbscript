option explicit
dim objShell, objFile
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
dim directory

sub Main()
    dim input, command
    directory = objFile.GetParentFolderName(wscript.ScriptFullName)
    input = inputbox("Enter the name of the file you want to run in the current directory.")

    if input = "" then
        wscript.quit
    end if

    command = "cd ""f:/Ionvop/Documents/VSCode Projects/ivvbscript/"" && javac Program.java && java Program """ & directory & "\" & input & ".ivvbs"""
    objFile.CreateTextFile(directory & "/run.bat", true).write(command)
    objShell.run("""" & directory & "/run.bat""")
    wscript.sleep(1000)
    objFile.DeleteFile(directory & "/run.bat")
end sub

sub Breakpoint(input)
    wscript.echo(input)
    wscript.quit
end sub

Main()