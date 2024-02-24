option explicit
dim ivvbsobjshell, ivvbsobjfile, ivvbsobjhttp
set ivvbsobjshell = createobject("wscript.shell")
set ivvbsobjfile = createobject("scripting.filesystemobject")
set ivvbsobjhttp = createobject("msxml2.xmlhttp.6.0")
ivvbsobjshell.currentdirectory = ivvbsobjfile.getparentfolderName(ivvbsobjfile.getabsolutepathname(wscript.arguments(0)))
sub ivvbsprint(ivvbsmessage)
dim ivvbselement
if typename(ivvbsmessage) = "Variant()" then
wscript.stdout.write("[")
for each ivvbselement in ivvbsmessage
if typename(ivvbselement) = "Variant()" then
ivvbsprint(ivvbselement)
else
wscript.stdout.write(ivvbselement & ",")
end if
next
wscript.echo("]")
exit sub
end if
wscript.echo(ivvbsmessage)
end sub
