' VBScript Source File Created for any
' NAME: whoami.vbs

Const HKCR = &H80000000
Const HKCU = &H80000001
Const HKLM = &H80000002
Const HKU = &H80000003

Set fs = CreateObject("Scripting.FileSystemObject")

'**********************************
'ComputerName aus Registry auslesen
'**********************************

s_Key = "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName"

s_Wert = "ComputerName"
Set wmireg = GetObject("winmgmts:root\default:StdRegProv")

result = wmireg.GetStringValue(HKLM, s_Key, s_Wert, s_ComputerName)

'**********************************
'UserName aus Registry auslesen
'**********************************

s_Key = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"

s_Wert = "DefaultUserName"
Set wmireg = GetObject("winmgmts:root\default:StdRegProv")

result = wmireg.GetStringValue(HKLM, s_Key, s_Wert, s_UserName)

'*****************************************
'IPConfig Daten auslesen und echo ausgeben
'*****************************************

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set IPConfigSet = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

Set colAdapters = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each IPConfig in IPConfigSet
   If Not IsNull(IPConfig.IPAddress) Then
      For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
      ausgabe = Wscript.Echo ("IP Address " & IPConfig.IPAddress(i) & vbCrLf & "Computer " & s_ComputerName & vbCrLf & "Username " & s_UserName)
      Next
   End If
Next
