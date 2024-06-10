Dim wmp
Set wmp = CreateObject("WMPlayer.OCX")
wmp.cdromcollection.item(1).eject()
OK=MsgBox ("Close I:?", OK)
wmp.cdromcollection.item(1).eject()
