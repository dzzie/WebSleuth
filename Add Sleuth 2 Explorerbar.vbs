
set fso = createobject("Scripting.FileSystemObject")

home=fso.GetParentFolderName(wscript.scriptfullname)

if not fso.FileExists(home & "\WebSleuth.exe") then 
	msgbox "This script can only be run from web sleuth home" & _
		   "directory. It will Add Sleuth Icon to IE Toolbar" & _
		   "You must also have the 2 icons in Source directory"
    wscript.quit
end if 

if not fso.FileExists(home & "\Source\open.ico") or _
   not fso.FileExists(home & "\Source\close.ico") then 
   msgbox "Oops couldnt find icons for button exiting"
   wscript.quit
end if    


set w=createobject("Wscript.Shell")
guid="{F3D1ABFB-AB7F-4ea2-A9B2-23D41D1CDCF6}"
base="HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\" & guid

w.RegWrite base & "\Default Visible","Yes", "REG_SZ"
w.RegWrite base & "\ButtonText", "Websleuth", "REG_SZ"
w.RegWrite base & "\HotIcon", home & "\Source\open.ico", "REG_SZ"
w.RegWrite base & "\Icon", home & "\Source\close.ico", "REG_SZ"
w.RegWrite base & "\CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}", "REG_SZ"
w.RegWrite base & "\Exec", home & "\WebSleuth.exe", "REG_SZ"

msgbox "IE Button Registration Successful !"