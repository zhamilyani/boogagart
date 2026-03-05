Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d C:\cygwin64\home\sulej\projects\boogagart && http-server -p 3007 -c-1", 0, False
