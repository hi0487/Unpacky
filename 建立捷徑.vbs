Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
Set lnk = ws.CreateShortcut(ws.SpecialFolders("Desktop") & "\開箱寶 Unpacky.lnk")
lnk.TargetPath = base & "\開箱寶 Unpacky.exe"
lnk.WorkingDirectory = base
lnk.IconLocation = base & "\開箱寶 Unpacky.ico, 0"
lnk.Save
MsgBox "桌面捷徑已建立！", 64, "開箱寶 Unpacky"