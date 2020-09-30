'============================================================================
'�������������� ����������� ���������� WMI � WSH (02.07.2012)
'������ ��������� � ��������� ���� XML ��� ����������� ��������� �������� iParser
'�����:������� ������� lanc@live.ru
'������ 1.09
'��������� �������
'1 �������� hand ��� auto
'2 �������� ���� ��� ���������� ����������� (������������ ��� ������ �������, ����� ������� ���������� INV � ����� ������� �������)
'3 �������� ������ �������� ��� ������������, ����������� ",". ������������ ��� ������ �������.
'������� ����� ������� 
'cscript invent.vbs auto "\\server\folder" "192.168.10,192.168.11,192.168.12"
'cscript invent.vbs hand "c:\localfolder\"
'cscript invent.vbs hand
'============================================================================
Option Explicit	
On Error Resume Next

'============================================================================
'���������� ���������
Const TITLE = "���� ������������������ ������!" '��������� ���������� ����
Const UPDATES = False '�� ��������� ����������
Const DATA_EXT = ".xml" '���������� ����� ������
Const Data_LOG = ".log" '���������� ����� ������ ������
Const HEAD_LINE = True '�������� ��������� � ������ ������ CSV-�����
'Const CstrUser = "ITAdmin" ' ����� ���������� ��������������
'CONST CstrPwd = "P@ssw0rd1"  '������ ���������� �������������� 

'============================================================================
'���������� ����������
Dim arr_WMI_classes '������ ���������� WMI-�������
Dim DATA_DIR, DATA_LOG_DIR
Dim AUTO '��� �������������� ������� AUTO=True
Dim Subnets '������ ��������
Dim strPath '���� � ���������� ���������� �����������
Dim isDomain '��� ������� ��������� ������������� isDomain=False
Dim strComputerName, strUserName
Dim args, objFSO, objWshShell, objNetwork, objWMIService, objSID, objGroup
Dim objUser, objComputer, WClass 
Dim x,y,i, xsub, strEndScript, instance_number
Dim strInvNumber '����������� �����
Dim tf '���� ������
Dim text_log ' ���� ������ ������
'Dim compsoft, wmiosoft
'Dim nwo, wmio, 
Dim wmiosoft
Dim tempbool, substr
Dim sreg
Dim tmp, tmpctr  
Dim arrMonitorInfo(), intMonitorCount 
Dim strPrefix

'============================================================================
'����������� ���������� ��������
Set args = WScript.Arguments
Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'============================================================================
'���������� ���������� ����������
strComputerName = objNetwork.ComputerName
strUserName = objNetwork.UserName
strEndScript = "������ �������������� ��������� ������"
DATA_DIR = "Data\" '��������� ������� ���������� ������ - ������������ �� ���������, ���� �� ������� ����������
DATA_LOG_DIR = "Logs\" '��������� ������� ���������� ����
arr_WMI_classes = arr_WMI_classes_fill '
strPrefix = "M"
'============================================================================
'������ � �����������
'============================================================================
' #region �������� ����������
If args.count = 0 Then
tmp = "�� ������� ��������� �������!" & vbCrLf & "��� �������: " & vbCrLf & "1. �������� - HAND ��� AUTO" & vbCrLf & _
"2. �������� - ���� ��� ���������� ����������� (������������ ��� ������ �������, ����� ������� ���������� INV � ����� ������� �������)" & vbCrLf & _
"3. �������� - ������ �������� ��� ������������, ����������� "","". ������������ ��� �������������� �������." & vbCrLf & _
"�������:" & vbCrLf & _
"cscript invent.vbs auto ""\\server\folder"" ""192.168.10,192.168.11,192.168.12""" & vbCrLf & _
"cscript invent.vbs hand ""c:\localfolder\""" & vbCrLf & _
"cscript invent.vbs hand"
WScript.echo tmp
WScript.Quit
End If

If args.Item(0) = "hand" Then
	INVNumber_check '������ � �������� ������������ ������
	AUTO = False
	If args.Count > 1 Then DATA_DIR = args.Item(1)
	If objFSO.FolderExists(DATA_DIR) <> True Then objFSO.CreateFolder(DATA_DIR)
	If objFSO.FolderExists(DATA_LOG_DIR) <> True Then objFSO.CreateFolder(DATA_LOG_DIR)
	Set text_log = objFSO.CreateTextFile(DATA_LOG_DIR & strPrefix & strInvNumber & Data_LOG, True)
	Trace "������ ������ ��������������"
	Trace  "������ � ������ ���������: ����������� ����� " & strInvNumber
	
ElseIf args.Item(0) = "auto" Then
	AUTO = True
	If args.Count < 3 Then
		WScript.echo "������. ��� �������������� ������� ���������� ������� ��� ���������"
		WScript.echo strEndScript
		WScript.Quit
	Else 
		DATA_DIR = args.Item(1)
		Subnets=Split(args.Item(2),",")
	End If
	Trace "�������������� ������ ��������������"
Else
Trace "������ � ��������� ������/�������������� ������"
Trace strEndScript
WScript.Quit
End If

If AUTO Then
	
	'�������� ������ ��������
	For each x In Subnets
		xsub = Split(x,".")
		If ubound(xsub)<>2 Then
			Trace "������. ������ ������� ������ ���� ����: 192.168.10"
			Trace strEndScript
			WScript.Quit
		End If
		For Each y In xsub
			If IsNumeric(y)=False Then
				Trace "������. � ������ ������� �� �������� ��������. ��� ������: 192.168.10"
				Trace strEndScript
				WScript.Quit
			ElseIf y>255 Then
				Trace "������. � ������ ������� �������� ������ 255. ��� ������: 192.168.10"
				Trace strEndScript
				WScript.Quit
			End If
		Next
	Next
End If' #endregion

'============================================================================
'��� ������ �������
'����������� �������������� ������
'�������� ����������������� �������
'============================================================================
' #region ��������� ��� ������� �������
' If AUTO=False Then
' 	If objNetwork.ComputerName = objNetwork.UserDomain Then
' 		isDomain = False
' 	Else
' 		isDomain = True
' 	End If
 
 
 '��������� ��������� �� ������� ������������ � ������
' Set objSID = objWMIService.Get("Win32_SID.SID='S-1-5-32-544'")
' Set objGroup = GetObject("WinNT://./" & objSID.AccountName & ",group")
' Set objComputer = GetObject ("WinNT://" & strComputerName) 
' 	' tempbool = false
' ' 	For Each objUser in objGroup.Member
' ' 		If objUser.name = CstrUser Then tempbool = True End If
' ' 	Next
' ' 	If tempbool
' 	Set objUser = objComputer.Create("user", CstrUser) 
' 	objUser.SetPassword CstrPwd
' 	'objUser.SetInfo
' 	If Err.Number Then
' 		If Err.Number=-2147022672 Then
' 			Trace "������� �������� ��� ������������� ������������"
' 			Set objUser = GetObject("WinNT://" & strComputerName & "/" & CstrUser)
' 			objUser.SetPassword CstrPwd
' 			objUser.SetInfo
' 		ElseIf Err.Number=-2147024891 Then
'         	Trace Err.Number
'         	Trace "������. ������ ������ ���� ������� � ����������������� ������������"
'         	Trace strEndScript
' 			WScript.Quit
' 		Else
' 			Trace "������. ����������� ������ ��� �������� ���������� ������������"
' 			Trace strEndScript
' 			WScript.Quit
' 		End If
'     End If   
' End If
' #endregion


'���� ������
	Set tf = objFSO.CreateTextFile(DATA_DIR & strPrefix & strInvNumber & DATA_EXT, True)
 
		tf.WriteLine "<?xml version=""1.0"" encoding=""windows-1251""?>"
		tf.WriteLine "<!--���� ������������������ ������ ������� ������� " & strComputerName & ", ���. ����� " & strInvNumber & "  -->"
		tf.WriteLine "<!--���� ����� ���������� " & Now & "  -->"
		tf.WriteLine "<InventoryData>"
'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ��������� �������������. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear

'�������� � ������� � ��������� ���������
tf.WriteLine "  <LocalPrinterInfo>"
Printer_Data_Write
tf.WriteLine "  </LocalPrinterInfo>"

'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ��������� ��������� LocalPrinterInfo. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear

'����� ���������� �� ������ WMI �������
For Each WClass In arr_WMI_classes
tf.WriteLine "  <" & WClass & ">"
WMI_Data_Write WClass
tf.WriteLine "  </" & WClass & ">"
	'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ��������� WMI_Data_Write. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear
Next


'�������� � ������� � ��������
MonitorInfo
'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ������� Monitor_Info. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear

Monitor_Data_Write
'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ��������� Monitor_Data_Write. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear

'�������� � ������� � ������������� �����������
InventSoft
'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ��������� InventSoft. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear

tf.WriteLine "</InventoryData>"

Trace "�������������� ������� ���������!"

Sub Printer_Data_Write
Dim intPrinterCount, strPrinterModel,strPrinterSerNumber, ii
Trace "��������� LocalPrinterInfo"
intPrinterCount = PriterCount_check
If intPrinterCount > 0 Then
	ii = 0
	While ii<intPrinterCount
		strPrinterModel = InputBox ("������� ������ ��� ���������� �������� �" & ii+1,TITLE)
		strPrinterSerNumber = InputBox ("������� �������� ����� ��� ���������� �������� �" & ii+1,TITLE)
		substr = "      <N" & ii+1
		substr = substr & " Model=" & """" & RemoveSpecialChars(strPrinterModel) & """" & " SerNumber=" & """" & RemoveSpecialChars(strPrinterSerNumber) & """"
		tf.WriteLine substr & " />"
		ii=ii+1
	Wend  
End If
End Sub

Sub InventSoft

 Set wmiosoft = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\Root\default:StdRegProv")
Trace "��������� Software"

	Dim ssoft
	ssoft = ExtractSoft("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
	'If Len(ssoft) > 0 Then tfsoft.Write ssoft

	'��� 64-������ ������ ���� ��� ������ ����! (32-������ ��������� �� 64-������ �������)
	ssoft = ExtractSoft("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
'	If Len(ssoft) > 0 Then tfsoft.Write ssoft
	
End Sub

Function ExtractSoft(key)

' 	'�������� ���������
	Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
	Dim items
	wmiosoft.EnumKey HKLM, key, items
	If IsNull(items) Then
		ExtractSoft = ""
		Exit Function
	End If
 
	'�������� ������ ��������
	Dim ssoft, itemsoft, oksoft, strDisplayName, strInstallDate, xsoft, prevsoft
	Dim strDisplayVersion, strInstallLocation, strPublisher
 	'ssoft = "" '��������� ����������� � ������
 	tf.WriteLine "  <InstalledSoft>"
	instance_number = 1
 	For Each itemsoft In items
 		oksoft = True '���� �����������
 		'��������, ���������� ������ � �������������
 		prevsoft = strDisplayName
 		wmiosoft.GetStringValue HKLM, key & itemsoft, "DisplayName", strDisplayName
		If IsNull(strDisplayName) Or Len(strDisplayName) = 0 Or strDisplayName = prevsoft Then
			oksoft = False
		Else
 		End If
 				
		'�������� ��������, �� �������� ��������� ParentKeyName = "OperatingSystem"
		'If oksoft Then
		'	wmiosoft.GetStringValue HKLM, key & itemsoft, "ParentKeyName", xsoft
		'		If IsNull(xsoft) Or xsoft <> "OperatingSystem" Then oksoft = False
		'End If

		'���� ���������
		If oksoft Then
			wmiosoft.GetStringValue HKLM, key & itemsoft, "InstallDate", strInstallDate
			If IsNull(strInstallDate) Or Len(strInstallDate) < 8 Then
				strInstallDate = ""
			Else '������������� � �������� ���
				strInstallDate = Mid(strInstallDate, 7, 2) & "." & Mid(strInstallDate, 5, 2) & "." & Left(strInstallDate, 4)
			End If
			If InStr(strInstallDate,"-")<>0 or InStr(strInstallDate," ")<>0 Then strInstallDate = ""
		End If

		'�������������, ������, ���������� ���������
		If oksoft Then
			wmiosoft.GetStringValue HKLM, key & itemsoft, "Publisher", strPublisher
			If IsNull(strPublisher) Or Len(strPublisher) = 0 Then strPublisher = ""
			wmiosoft.GetStringValue HKLM, key & itemsoft, "DisplayVersion", strDisplayVersion
			If IsNull(strDisplayVersion) Or Len(strDisplayVersion) = 0 Then strDisplayVersion = ""
			wmiosoft.GetStringValue HKLM, key & itemsoft, "InstallLocation", strInstallLocation
			If IsNull(strInstallLocation) Or Len(strInstallLocation) = 0 Then strInstallLocation = ""
			strInstallLocation = Replace(strInstallLocation,"?","")
			substr ="      <N" & instance_number  & " DisplayName=" & """" & RemoveSpecialChars(strDisplayName) & """" & " InstallDate=" & """" & RemoveSpecialChars(strInstallDate) & """" & " Puslisher=" & """" & RemoveSpecialChars(strPublisher) & """" & " DisplayVersion=" & """" & RemoveSpecialChars(strDisplayVersion) & """" & " InstallLocation=" & """" & RemoveSpecialChars(strInstallLocation) & """" & " />"
			tf.WriteLine substr
		End If
	instance_number=instance_number+1
 	Next
 	tf.WriteLine "  </InstalledSoft>"
	ExtractSoft = ssoft
 End Function

'===============================================================================================
'�������� ���������� ���������� ���������
'===============================================================================================
Function PriterCount_check
Dim intPrinterCount
intPrinterCount = 0
tempbool = False
  Do Until tempbool = True 	  		
  intPrinterCount = InputBox ("������� ���������� ������������ ��������� ���������",TITLE)
		     If Len(intPrinterCount) > 0 Then 			'���������� �������� ������ ��������������
		     	If IsNumeric(intPrinterCount) Then    '�������� ����, �������� �� ��������� ���������� ������.
		            tempbool = True 
		     	Else
		     		tmp = MsgBox("������: ������� ������������ ������. ���������� ������� �����!",16,TITLE)
		        	WScript.echo "������: ������� ������������ ������. ���������� ������� �����!"
		        End If
     		 Else
     		 	intPrinterCount = 0
     		 	tempbool = True
     End If
  Loop 
PriterCount_check = CInt(intPrinterCount)
'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ������� PriterCount_check. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear
End Function

'===============================================================================================
'�������� ���������� ������������ ������ � ������ � ������
'===============================================================================================
Sub INVNumber_check ()
tempbool = False
  Do Until tempbool = True 	  		
  strInvNumber = InputBox ("������� ����������� ����� ����������." & vbCrLf & "4 �������" & vbCrLf & "������: 0138",TITLE)
		     If Len(strInvNumber)= 4 Then 			'�������� ���������� �������� ������ ��������������
		     	If IsNumeric(strInvNumber) Then    '�������� ����, �������� �� ��������� ���������� ������.
		            tempbool = True 
		     	Else
		     		tmp = MsgBox("������: ������� ������������ ������. ���������� ������� �����!",16,TITLE)
		        	WScript.echo "������: ������� ������������ ������. ���������� ������� �����!"
		        	End If
     		 ElseIf Len(strInvNumber)= 0 Then
     		 	WScript.echo strEndScript
     		    WScript.Quit  		 
     		 Else
     		 tmp = MsgBox("������: ���������� �������� ���������� ������������ ������ �� ���������!",16,TITLE)
 	         WScript.echo "������: ���������� �������� ���������� ������������ ������ �� ���������!"
     End If
  Loop 
  On Error Resume Next
  objWshShell.RegWrite "HKEY_LOCAL_MACHINE\software\Inventory", strInvNumber, "REG_SZ" '������ � ������
' If Err.Number>0 Then
' WScript.Echo "Err.Number " + Err.Number
' ' 		If Err.Number=-2147022672 Then
' End If
'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ������� INVNumber_check. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear
'sreg = objWshShell.RegRead("HKEY_LOCAL_MACHINE\software\Inventory")
End Sub
'===============================================================================================
'�����������
'===============================================================================================
Sub Trace (LogText)
text_log.WriteLine(Date & " " & Time  & ": " & LogText)
WScript.echo Date & " " & Time  & ": " & Logtext
End Sub

Sub Monitor_Data_Write
'0=VESA Mfg ID, 1=VESA Device ID, 2=MFG Date (M/YYYY),3=SerialNum (If available),4=Model Descriptor 
' '5=EDID Version 

tf.WriteLine "  <MonitorInfo>"
	For tmpctr=0 To intMonitorCount-1
		substr = "      <N" & tmpctr+1
		substr = substr & " VESA_Mfg_ID=" & """" & RemoveSpecialChars(arrMonitorInfo(tmpctr,0)) & """" & " VESA_DEVICE_ID=" & """" & RemoveSpecialChars(arrMonitorInfo(tmpctr,1)) & """" & " MFG_Date=" & """" & RemoveSpecialChars(arrMonitorInfo(tmpctr,2)) & """" & " SerialNum=" & """" & RemoveSpecialChars(arrMonitorInfo(tmpctr,3)) & """" & " Model_Descriptor=" & """" & RemoveSpecialChars(arrMonitorInfo(tmpctr,4)) & """" & " EDID_Version=" & """" & RemoveSpecialChars(arrMonitorInfo(tmpctr,5)) & """"
		tf.WriteLine substr & " />"
	Next
	tf.WriteLine "  </MonitorInfo>"

End Sub

'===============================================================================================
'������� ��� ������ ���� ��������
'===============================================================================================
Function RemoveSpecialChars(str)
Dim endstr
endstr = str
For i = 1 To Len(endstr)
	tmp = Mid(endstr,i,1)
	if tmp <> "" then
		If Asc(tmp) < 32 Then 
		endstr = ""
		End If
	end If
	i=i+1
Next
endstr = Trim(Replace(endstr, ";", "_"))
endstr = Trim(Replace(endstr, """", "&quot;"))
endstr = Trim(Replace(endstr, "&", "&amp;"))
endstr = Trim(Replace(endstr, "<", "_"))
endstr = Trim(Replace(endstr, ">", "_"))
RemoveSpecialChars = endstr
End Function
'===============================================================================================
'������ ������������������ ������ WMI � xml ����
'WMI_class - ��� ����������� ������
'� ���������� ������ ����� ��������� �������� ���� ������� ������
'===============================================================================================
Sub WMI_Data_Write(WMI_class)
	Const RETURN_IMMEDIATELY = 16
	Const FORWARD_ONLY = 32
	Dim query, classes, item, prop, value
	Trace "��������� " & WMI_class
	query = "Select * From " & WMI_class

	Set classes = objWMIService.ExecQuery(query,, RETURN_IMMEDIATELY + FORWARD_ONLY)
	
	'tf.WriteLine "  <" & WMI_class & ">"
	instance_number = 1 '����� ����������
	For Each item In classes
		substr = "      <N" & instance_number
		For Each prop In item.Properties_
			value = prop.Value
			If IsNull(value) Then 
				value = ""
			ElseIf IsArray(value) Then
				value = Join(value,",")
			ElseIf prop.CIMType = 101 Then
				value = ReadableDate(value)
			End If
			If Len(value) > 150 Then
			value = ""
			End If
			If Len(value) > 0 Then 
				 substr = substr & " " & prop.Name & "=" & """" & RemoveSpecialChars(value) & """"
			End If
		Next
		tf.WriteLine substr & " />"
		instance_number = instance_number + 1
	Next
	'tf.WriteLine "  </" & WMI_class & ">"
End Sub

Sub Log(from, sel, where, sect, param)
		'��������� WQL-������, ��������� � �������� ������ � X-����
		'������� ���������:
		'from - ����� WMI
		'sel - �������� WMI, ����� �������
		'where - ������� ������ ��� ������ ������
		'sect - ��������������� ������ ������
		'param - ��������������� ��������� ������ ������ ������, ����� �������
		'��� ����������� � ������� ��������, ����� �� ������� � �������
		
' 		Log "Win32_ComputerSystemProduct", _
' 		"UUID", "", _
' 		"���������", _
' 		"UUID"
' 
' 	Log "Win32_ComputerSystem", _
' 		"Name,Domain,PrimaryOwnerName,UserName,TotalPhysicalMemory", "", _
' 		"���������", _
' 		"������� ���,�����,��������,������� ������������,����� ������ (��)"
'sect - ������ ���� ���������
'��� ��������� � ������ � ������ ������ ���������� num 

		
	Const RETURN_IMMEDIATELY = 16
	Const FORWARD_ONLY = 32
	Dim substr
	Dim query, cls, item, prop
	query = "Select " & sel & " From " & from

	If Len(where) > 0 Then query = query & " Where " & where
	Set cls = wmio.ExecQuery(query,, RETURN_IMMEDIATELY + FORWARD_ONLY)

	Dim props, names, num, value
	props = Split(sel, ",")
	names = Split(param, ",")

		'tf.WriteLine sect & ";" & names(i) & ";" & num & ";" & value
tf.WriteLine "  <" & sect & ">"

	num = 1 '����� ����������
	For Each item In cls
	
	'tf.WriteLine "    <" & names(i) & ">"
	
		For i = 0 To UBound(props)
		If i =0 Then substr = "      <N" & num
			'����� ��������
			Set prop = item.Properties_(props(i))
			value = prop.Value

			'��� �������� �� Null ��������� ����� � �������
			If IsNull(value) Then
				value = ""

			'���� ��� ������ - ������, ������� � ������
			ElseIf IsArray(value) Then
				value = Join(value,",")

			'���� ������� ������� ������� ���������, ��������� ��������
			ElseIf Right(names(i), 4) = "(��)" Then
				value = CStr(Round(value / 1024 ^ 2))
			ElseIf Right(names(i), 4) = "(��)" Then
				value = CStr(Round(value / 1024 ^ 3))

			'���� ��� ������ - ����, ������������� � �������� ���
			ElseIf prop.CIMType = 101 Then
				value = ReadableDate(value)
			End If

			'������� � ���� �������� ��������, �������� ���������� ";"
			value = Trim(Replace(value, ";", "_"))
			If Len(value) > 0 Then 
			'Trace sect & ";" & names(i) & ";" & num & ";" & value
				'tf.WriteLine sect & ";" & names(i) & ";" & num & ";" & value
				'tf.WriteLine "<" & names(i) & ">"
				 'tf.WriteLine "<" & i & ">"
				 
				 substr = substr & " " & props(i) & "=" & """" & value & """"
			End If
		Next 'i
		'If Len(value) > 0 Then
			tf.WriteLine substr & " />"
		'End If 
		WScript.Echo "�������� substr " & substr & " />"
	    'tf.WriteLine "    </" & names(i) & ">"
		'������� � ���������� ����������
		num = num + 1
	Next 'item
	tf.WriteLine "  </" & sect & ">"
	'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ��������� Log. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear
End Sub

Function arr_WMI_classes_fill()
Dim strWMI
strWMI = "Win32_ComputerSystem " & _
"Win32_NetworkAdapter " & _
"Win32_NetworkAdapterConfiguration " & _
"Win32_BaseBoard " & _
"Win32_BIOS " & _
"Win32_CDROMDrive " & _
"Win32_DiskDrive " & _
"Win32_LogicalDisk " & _
"Win32_MemoryDevice " & _
"Win32_Processor " & _
"Win32_OperatingSystem " & _
"Win32_Service " & _
"Win32_VideoController " & _
"Win32_QuickFixEngineering"
arr_WMI_classes_fill = split(strWMI, " ")
'����������� ������		
If Err.Number <> 0 Then
   Trace "������ � ������� arr_WMI_classes_fill. ����� ������: " & Err.Number & ". �������� ������: " & Err.Description
End If
Err.Clear
End Function 

Function ReadableDate(str)
	'�������������� ���� ������� DMTF � �������� ��� (��.��.����)
	'http://msdn.microsoft.com/en-us/library/aa389802.aspx
	ReadableDate = Mid(str, 7, 2) & "." & Mid(str, 5, 2) & "." & Left(str, 4)
End Function


Function BuildVersion()
	'������ ������ (����) WMI-�������
	'������� ����� �����
	Dim cls, item
	Set cls = wmio.ExecQuery("Select BuildVersion From Win32_WMISetting")
	For Each item In cls
		BuildVersion = CInt(Left(item.BuildVersion, 4))
	Next
	Trace.WriteLine Date & ":" & Time & ("������ WMI ������")
End Function


Function Unavailable(addr)

	'��������� ����������� ���������� � ����
	'������� True, ���� ����� ����������
	Dim wmio, ping, p
	Set wmio = GetObject("WinMgmts:{impersonationLevel=impersonate}")
	Set ping = wmio.ExecQuery("SELECT StatusCode FROM Win32_PingStatus WHERE Address = '" & addr & "'")
	For Each p In ping
		If IsNull(p.StatusCode) Then
			Unavailable = True
		Else
			Unavailable = (p.StatusCode <> 0)
		End If
	Next
	text_log.WriteLine Date & ":" & Time & ("���� ��������� ������ ��� ������������ ���������� ������")
End Function


Function MonitorInfo
Dim oRegistry, sBaseKey, sBaseKey2, sBaseKey3, skey, skey2, skey3 
Dim sValue 
Dim t, iRC, iRC2, iRC3 
Dim arSubKeys, arSubKeys2, arSubKeys3, arrintEDID 
Dim strRawEDID 
Dim ByteValue, strSerFind, strMdlFind 
Dim intSerFoundAt, intMdlFoundAt, findit
Dim tmpmfgweek, tmpmfgyear, tmpmdt, tmpEDIDMajorVer, tmpEDIDRev, tmpver, tmpEDIDMfg
Dim Char1, Char2, Char3, Byte1, Byte2 
Dim tmpmfg, tmpEDIDDev1, tmpEDIDDev2, tmpdev 
Dim tmpser, tmpmdl
dim strarrRawEDID()
intMonitorCount=0
Const HKLM = &H80000002 

Trace "��������� MonitorInfo"

Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "/root/default:StdRegProv")
sBaseKey = "SYSTEM\CurrentControlSet\Enum\DISPLAY\"

iRC = oRegistry.EnumKey(HKLM, sBaseKey, arSubKeys)
For Each sKey In arSubKeys
	sBaseKey2 = sBaseKey & sKey & "\"
	iRC2 = oRegistry.EnumKey(HKLM, sBaseKey2, arSubKeys2)
	If IsArray(arSubKeys2) then
	For Each skey2 In arSubKeys2
		oRegistry.GetMultiStringValue HKLM, sBaseKey2 & skey2 & "\", "HardwareID", sValue
		For tmpctr=0 to ubound(svalue)
			If lcase(left(svalue(tmpctr),8))="monitor\" Then
				sBaseKey3 = sBaseKey2 & skey2 & "\"
				iRC3 = oRegistry.EnumKey(HKLM, sBaseKey3, arSubKeys3)
				For Each skey3 In arSubKeys3
					If skey3="Control" Then
						oRegistry.GetBinaryValue HKLM, sBaseKey3 & "Device Parameters\", "EDID", arrintEDID
						if vartype(arrintedid) <> 8204 Then 
							strRawEDID="EDID Not Available" 
						Else
							for each bytevalue In arrintEDID 
								strRawEDID=strRawEDID & chr(bytevalue)
							Next
						end If
						ReDim Preserve strarrRawEDID(intMonitorCount)
						strarrRawEDID(intMonitorCount)=strRawEDID
						intMonitorCount=intMonitorCount+1
					end If
				Next
			end If
		Next
	Next
	End if 
Next

redim arrMonitorInfo(intMonitorCount-1,5)
dim location(3)
for tmpctr=0 to intMonitorCount-1
	If strarrRawEDID(tmpctr) <> "EDID Not Available" Then
		location(0)=mid(strarrRawEDID(tmpctr),&H36+1,18)
		location(1)=mid(strarrRawEDID(tmpctr),&H48+1,18)
		location(2)=mid(strarrRawEDID(tmpctr),&H5a+1,18)
		location(3)=mid(strarrRawEDID(tmpctr),&H6c+1,18)
		strSerFind=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hff)
		strMdlFind=chr(&H00) & chr(&H00) & chr(&H00) & chr(&Hfc)
		intSerFoundAt=-1
		intMdlFoundAt=-1
		for findit = 0 to 3
			If instr(location(findit),strSerFind)>0 Then
				intSerFoundAt=findit
			end If
			If instr(location(findit),strMdlFind)>0 Then
				intMdlFoundAt=findit
			end If
		Next

		If intSerFoundAt<>-1 Then
			tmp=right(location(intSerFoundAt),14)
			If instr(tmp,chr(&H0a))>0 Then
				tmpser=trim(left(tmp,instr(tmp,chr(&H0a))-1))
			Else
				tmpser=trim(tmp)
			End If
			If left(tmpser,1)=chr(0) then tmpser=right(tmpser,len(tmpser)-1)
		Else
			tmpser="Serial Number Not Found in EDID data"
		end If

		if intMdlFoundAt<>-1 Then
			tmp=right(location(intMdlFoundAt),14)
			if instr(tmp,chr(&H0a))>0 Then
				tmpmdl=trim(left(tmp,instr(tmp,chr(&H0a))-1))
			Else
				tmpmdl=trim(tmp)
			end If
			if left(tmpmdl,1)=chr(0) then tmpmdl=right(tmpmdl,len(tmpmdl)-1)
		Else
			tmpmdl="Model Descriptor Not Found in EDID data"
		end If
		tmpmfgweek=asc(mid(strarrRawEDID(tmpctr),&H10+1,1))
		tmpmfgyear=(asc(mid(strarrRawEDID(tmpctr),&H11+1,1)))+1990
		tmpmdt=month(dateadd("ww",tmpmfgweek,datevalue("1/1/" & tmpmfgyear))) & "/" & tmpmfgyear
		tmpEDIDMajorVer=asc(mid(strarrRawEDID(tmpctr),&H12+1,1))
		tmpEDIDRev=asc(mid(strarrRawEDID(tmpctr),&H13+1,1))
		tmpver=chr(48+tmpEDIDMajorVer) & "." & chr(48+tmpEDIDRev)
		tmpEDIDMfg=mid(strarrRawEDID(tmpctr),&H08+1,2) 
		Char1=0 : Char2=0 : Char3=0 
		Byte1=asc(left(tmpEDIDMfg,1)) 
		Byte2=asc(right(tmpEDIDMfg,1)) 
		if (Byte1 and 64) > 0 then Char1=Char1+16 
		if (Byte1 and 32) > 0 then Char1=Char1+8 
		if (Byte1 and 16) > 0 then Char1=Char1+4 
		if (Byte1 and 8) > 0 then Char1=Char1+2 
		if (Byte1 and 4) > 0 then Char1=Char1+1 
		if (Byte1 and 2) > 0 then Char2=Char2+16 
		if (Byte1 and 1) > 0 then Char2=Char2+8 
		if (Byte2 and 128) > 0 then Char2=Char2+4 
		if (Byte2 and 64) > 0 then Char2=Char2+2 
		If (Byte2 and 32) > 0 then Char2=Char2+1 

		Char3=Char3+(Byte2 and 16) 
		Char3=Char3+(Byte2 and 8) 
		Char3=Char3+(Byte2 and 4) 
		Char3=Char3+(Byte2 and 2) 
		Char3=Char3+(Byte2 and 1) 
		tmpmfg=chr(Char1+64) & chr(Char2+64) & chr(Char3+64)
		tmpEDIDDev1=hex(asc(mid(strarrRawEDID(tmpctr),&H0a+1,1)))
		tmpEDIDDev2=hex(asc(mid(strarrRawEDID(tmpctr),&H0b+1,1)))
		if len(tmpEDIDDev1)=1 then tmpEDIDDev1="0" & tmpEDIDDev1
		if len(tmpEDIDDev2)=1 Then tmpEDIDDev2="0" & tmpEDIDDev2
		tmpdev=tmpEDIDDev2 & tmpEDIDDev1
		arrMonitorInfo(tmpctr,0)=tmpmfg
		arrMonitorInfo(tmpctr,1)=tmpdev
		arrMonitorInfo(tmpctr,2)=tmpmdt
		arrMonitorInfo(tmpctr,3)=tmpser
		arrMonitorInfo(tmpctr,4)=tmpmdl
		arrMonitorInfo(tmpctr,5)=tmpver
	end If
Next

MonitorInfo = arrMonitorInfo

End Function







' 
' 		'�������������
' 		If oksoft Then
' 			wmiosoft.GetStringValue HKLM, key & itemsoft, "Publisher", publsoft
' 			If IsNull(publsoft) Or Len(publsoft) = 0 Then publsoft = "-"
' 		End If
' 
' 		If oksoft Then ssoft = ssoft & namesoft & ";" & publsoft & ";" & instsoft & vbCrLf
' 
' 	Next
' 	ExtractSoft = ssoft
' 
' End Function
' 
' '��������� ����������� ���������� � ����; ������� True, ���� ����� ����������
' Function Unavailable(addr)
' 	Dim wmio, ping, p
' 	Set wmio = GetObject("WinMgmts:{impersonationLevel=impersonate}")
' 	Set ping = wmio.ExecQuery("SELECT StatusCode FROM Win32_PingStatus WHERE Address = '" & addr & "'")
' 	For Each p In ping
' 		If IsNull(p.StatusCode) Then
' 			Unavailable = True
' 		Else
' 			Unavailable = (p.StatusCode <> 0)
' 		End If
' 	Next
' End Function
' 
Function CheckAccess

         Dim strTemp, arrayOutput, arrayMatchingOutput
         Dim arrayRegistryValue

         On Error Resume Next

         'If RunTimeEnvironmentInfo.OSIdentifier >= cWindowsVistaRTM Then
          '  CheckAccess = False
         'Else
            arrayRegistryValue = ReadRegistry (objLOGFileHandle, "HKLM\SYSTEM\CurrentControlSet\Services\winmgmt\Security", _
                                               "Security", _
                                               "REG_BINARY", _
                                               False)

            If Not IsArray (arrayRegistryValue) Then
               CheckAccess = True
            Else
               CheckAccess = False
            End If
         'End If

         ' If RunTimeEnvironmentInfo.OSIdentifier >= cWindows2003RTM Then
'             boolRC = ShellExecute (objLOGFileHandle, "WHOAMI.EXE /Priv", "", arrayOutput, arrayMatchingOutput)
'             If boolRC = False And IsArray (arrayOutput) = True Then
'                WriteToLogFile objLOGFileHandle, "", cLoggingLevel3, False
'                WriteToLogFile objLOGFileHandle, UBound (arrayOutput) & " privilege(s):", cLoggingLevel3, False
'                For Each strTemp In arrayOutput 
'                    If Len (strTemp) > 0 Then
'                       WriteToLogFile objLOGFileHandle, CleanString (strTemp, 32, 255), cLoggingLevel3, False
'                    End If 
'                Next 
'             End If
'          End If

End Function
