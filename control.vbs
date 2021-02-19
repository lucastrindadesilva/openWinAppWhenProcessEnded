Wscript.Sleep(1000)

Dim sQuery
Dim sComputerName

Dim sBoolTest
Dim sBoolProcess
Dim sProcessTestName
Dim sProcessOpenName 

Dim ws

'''''''''''''''''''''''''''''''''''''
sComputerName = "."

sBoolTest = True
sProcessTestName = "notepad.exe"
sProcessOpenName = "firefox.exe"
sProcessOpenPath = """C:\Program Files\Mozilla Firefox\firefox.exe"""

sQuery = "SELECT * FROM Win32_Process WHERE Name = '" & sProcessTestName & "'"
Set objWMIService = GetObject("winmgmts:\\" & sComputerName & "\root\cimv2")

Set objItems = objWMIService.ExecQuery(sQuery)

Do Until (Not sBoolTest)
sBoolTest = False
	For Each objItem In objItems
		If objItem.Name = sProcessTestName Then
			WScript.Echo "Processo [Nome:" & objItem.Name & " em execucao]"
			sBoolTest = True
			Exit For
		End If
	Next
	Wscript.Sleep(5000)	'5 segundos para prox execucao
	Set objItems = objWMIService.ExecQuery(sQuery)
Loop

WScript.Echo "Processo [Nome:" & sProcessTestName & " finalizado, iniciando app " & sProcessOpenName & "]"

Set ws = Wscript.CreateObject("Wscript.Shell")
ws.Run sProcessOpenPath, 1, True