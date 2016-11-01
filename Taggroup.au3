#Include <Excel.au3>
		;WinActivate ( "Tag Group Editor" )
		;Send("^n")
		;;Sleep(1000)
		;Send("y")
		;Send("y")
Local $sFilePath1 = @ScriptDir & "\test.xlsx" ;这个文件应该已经存在
Local $oExcel = _ExcelBookOpen($sFilePath1)
;$oExcel.Activesheet.Cells(6, 1).select
;$MyVal = $oExcel.Activesheet.Cells(6, 1).offset(1,1).value
;MsgBox(0, "错误!", $MyVal)
for $i=1 to 7 step 3
$MyVal = $oExcel.Activesheet.Cells(1, $i).value
;MsgBox(0, "错误!", $MyVal)
$oExcel.Activesheet.Cells(1, $i).offset(2,1).CurrentRegion.copy
WinActivate ( "Tag Group Editor" )
Send("^n")
Send("^v")
;Sleep(1000)
send("^s")
Sleep(1000)
ControlSend("另存为","",1001,$MyVal)
;Sleep(1000)
send("!s")
;Sleep(1000)
WinActivate("test.xlsx - Microsoft Excel")
;Sleep(1000)
next