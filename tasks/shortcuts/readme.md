---
permalink: /tasks/shortcuts/
---

# Windows Shortcuts

用 VBS 创建 Windows shortcuts(`*.lnk` 文件)

- [Working with Shortcuts](https://technet.microsoft.com/en-us/library/ee156583.aspx)

```vbs
' Create a shortcut for VSCode

Set shell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set env = shell.Environment("Process")

Sub CreateShortcut(lnkPath, targetPath, arguments)
    Set shortcut = shell.CreateShortcut(lnkPath)
    shortcut.TargetPath = targetPath
    shortcut.Arguments = arguments
    shortcut.Save
End Sub

lnkPath = "e:\Apps\Lnk\VSCode Development.lnk"
targetPath = "c:\Program Files\Microsoft VS Code Insiders\Code - Insiders.exe"
arguments = _
    "--user-data-dir=""" & fso.BuildPath(env("USERPROFILE"), ".vscode-development") & """" _
    & " --locale=en-US"
CreateShortcut lnkPath, targetPath, arguments
```
