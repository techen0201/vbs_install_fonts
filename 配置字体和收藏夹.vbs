Dim Fav_Path,Fonts_Path,Local_Path
strComputer = "."
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshell = CreateObject("Wscript.Shell")
Fav_Path = wshell.specialfolders("Favorites")
Fonts_Path = wshell.specialfolders("Fonts")
Local_Path = fso.GetFolder(".")

rem 配置字体安装参数
Set objShellApp = CreateObject("Shell.Application")
Const Fonts = &H14&
Fonts_Name = Array("方正大标宋简体","方正小标宋简体","SIMFANG","SIMKAI")


rem ----------------------------------------------------------------------------------------------------
rem 【复制收藏夹】
fso.DeleteFile Fav_Path&"\*.*",True
IF fso.FolderExists(Fav_Path&"\links") THEN
    fso.DeleteFile Fav_Path&"\links\*.*",True
ELSE
    fso.CreateFolder(Fav_Path&"\links")
    fso.CopyFile Local_Path&"\links\*.url",Fav_Path&"\links"
END IF

'Set fso_folders = fso.GetFolder(Fav_Path).SubFolders
'for Each fso_folder in fso_folders
'    fso_folder = fso_folder.name
'    fso.DeleteFolder(Fav_Path&"\"&fso_folder)
'next

rem ----------------------------------------------------------------------------------------------------
rem 【安装字体】
Set objFolder = objShellApp.Namespace(FONTS)
FOR i = 0 to UBound(Fonts_Name)
    objFolder.CopyHere Local_Path&"\fonts\"&Fonts_Name(i)&".TTF"
NEXT
msgbox "配置完成",64,"提示"