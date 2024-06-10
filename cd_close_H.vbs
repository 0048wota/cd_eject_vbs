Dim wmp
Set wmp = CreateObject("WMPlayer.OCX")
wmp.cdromcollection.item(0).eject()
OK=MsgBox ("Close H:?", OK)
wmp.cdromcollection.item(0).eject()
