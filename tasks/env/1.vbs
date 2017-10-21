Set shell = WScript.CreateObject("WScript.Shell")
Set env = shell.Environment("Process")

WScript.Echo env("USERPROFILE")
