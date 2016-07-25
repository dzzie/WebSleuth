on error resume next

set w=createobject("Wscript.Shell")
guid="{F3D1ABFB-AB7F-4ea2-A9B2-23D41D1CDCF6}"
base="HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Extensions\" & guid
foundit=empty


call getit
if foundit<>empty then 
	call dumpit	
	if  foundit = empty then msgbox "Sleuth IE Button Successfully Removed!"
else
	msgbox "Error could not find reg key"
end if 



sub GetIt()
  foundit= w.RegRead(base & "\Exec")
end sub

sub dumpit()
  foundit= w.regdelete(base & "\")
end sub

