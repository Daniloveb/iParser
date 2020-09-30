'============================================================================
'Инвентаризация компьютеров средствами WMI и WSH (02.07.2012)
'Данные выводятся в отдельный файл XML для последующей обработки сервисом iParser
'Автор:Данилов Евгений lanc@live.ru
'Версия 1.09
'Параметры запуска
'1 Параметр hand или auto
'2 Параметр путь для сохранения результатов (Необязателен при ручном запуске, будет создана директория INV в точке запуска скрипта)
'3 Параметр список подсетей для сканирования, разделитель ",". Необязателен при ручном запуске.
'Примеры строк запуска 
'cscript invent.vbs auto "\\server\folder" "192.168.10,192.168.11,192.168.12"
'cscript invent.vbs hand "c:\localfolder\"
'cscript invent.vbs hand
'============================================================================
Option Explicit	
On Error Resume Next

'============================================================================
'Глобальные константы
Const TITLE = "Сбор инвентаризационных данных!" 'заголовок диалоговых окон
Const UPDATES = False 'не учитывать обновления
Const DATA_EXT = ".xml" 'расширение файла отчета
Const Data_LOG = ".log" 'расширение файла отчета ошибок
Const HEAD_LINE = True 'выводить заголовки в первой строке CSV-файла
'Const CstrUser = "ITAdmin" ' Логин локального администратора
'CONST CstrPwd = "P@ssw0rd1"  'пароль локального администратора 

'============================================================================
'Глобальные переменные
Dim arr_WMI_classes 'Массив собираемых WMI-классов
Dim DATA_DIR, DATA_LOG_DIR
Dim AUTO 'При автоматическом запуске AUTO=True
Dim Subnets 'Массив подсетей
Dim strPath 'Путь к директории сохранения результатов
Dim isDomain 'При запуске локальным пользователем isDomain=False
Dim strComputerName, strUserName
Dim args, objFSO, objWshShell, objNetwork, objWMIService, objSID, objGroup
Dim objUser, objComputer, WClass 
Dim x,y,i, xsub, strEndScript, instance_number
Dim strInvNumber 'инвентарный номер
Dim tf 'файл отчета
Dim text_log ' файл отчета ошибок
'Dim compsoft, wmiosoft
'Dim nwo, wmio, 
Dim wmiosoft
Dim tempbool, substr
Dim sreg
Dim tmp, tmpctr  
Dim arrMonitorInfo(), intMonitorCount 
Dim strPrefix

'============================================================================
'Определение глобальных объектов
Set args = WScript.Arguments
Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objNetwork = CreateObject("WScript.Network")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

'============================================================================
'Присвоение глобальных переменных
strComputerName = objNetwork.ComputerName
strUserName = objNetwork.UserName
strEndScript = "Скрипт инвентаризации завершает работу"
DATA_DIR = "Data\" 'локальный каталог сохранения отчета - расположение по умолчанию, если не указано параметром
DATA_LOG_DIR = "Logs\" 'локальный каталог сохранения лога
arr_WMI_classes = arr_WMI_classes_fill '
strPrefix = "M"
'============================================================================
'Работа с параметрами
'============================================================================
' #region Проверка параметров
If args.count = 0 Then
tmp = "Не указаны параметры запуска!" & vbCrLf & "Вид запуска: " & vbCrLf & "1. Параметр - HAND или AUTO" & vbCrLf & _
"2. Параметр - путь для сохранения результатов (Необязателен при ручном запуске, будет создана директория INV в точке запуска скрипта)" & vbCrLf & _
"3. Параметр - список подсетей для сканирования, разделитель "","". Используется при автоматическом запуске." & vbCrLf & _
"Примеры:" & vbCrLf & _
"cscript invent.vbs auto ""\\server\folder"" ""192.168.10,192.168.11,192.168.12""" & vbCrLf & _
"cscript invent.vbs hand ""c:\localfolder\""" & vbCrLf & _
"cscript invent.vbs hand"
WScript.echo tmp
WScript.Quit
End If

If args.Item(0) = "hand" Then
	INVNumber_check 'Запрос и проверка инвентарного номера
	AUTO = False
	If args.Count > 1 Then DATA_DIR = args.Item(1)
	If objFSO.FolderExists(DATA_DIR) <> True Then objFSO.CreateFolder(DATA_DIR)
	If objFSO.FolderExists(DATA_LOG_DIR) <> True Then objFSO.CreateFolder(DATA_LOG_DIR)
	Set text_log = objFSO.CreateTextFile(DATA_LOG_DIR & strPrefix & strInvNumber & Data_LOG, True)
	Trace "Ручной запуск инвентаризации"
	Trace  "Запись в реестр добавлена: Инвентарный номер " & strInvNumber
	
ElseIf args.Item(0) = "auto" Then
	AUTO = True
	If args.Count < 3 Then
		WScript.echo "Ошибка. При автоматическом запуске необходимо указать три параметра"
		WScript.echo strEndScript
		WScript.Quit
	Else 
		DATA_DIR = args.Item(1)
		Subnets=Split(args.Item(2),",")
	End If
	Trace "Автоматический запуск инвентаризации"
Else
Trace "Ошибка в параметре Ручной/автоматический запуск"
Trace strEndScript
WScript.Quit
End If

If AUTO Then
	
	'Проверка записи подсетей
	For each x In Subnets
		xsub = Split(x,".")
		If ubound(xsub)<>2 Then
			Trace "Ошибка. Строка подсети должна быть вида: 192.168.10"
			Trace strEndScript
			WScript.Quit
		End If
		For Each y In xsub
			If IsNumeric(y)=False Then
				Trace "Ошибка. В строке подсети не числовое значение. Вид записи: 192.168.10"
				Trace strEndScript
				WScript.Quit
			ElseIf y>255 Then
				Trace "Ошибка. В строке подсети значение больше 255. Вид записи: 192.168.10"
				Trace strEndScript
				WScript.Quit
			End If
		Next
	Next
End If' #endregion

'============================================================================
'При ручном запуске
'Определение принадлежности домену
'Проверка административного доступа
'============================================================================
' #region Обработки для ручного запуска
' If AUTO=False Then
' 	If objNetwork.ComputerName = objNetwork.UserDomain Then
' 		isDomain = False
' 	Else
' 		isDomain = True
' 	End If
 
 
 'Проверяем находится ли текущий пользователь в группе
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
' 			Trace "Попытка создания уже существующего пользователя"
' 			Set objUser = GetObject("WinNT://" & strComputerName & "/" & CstrUser)
' 			objUser.SetPassword CstrPwd
' 			objUser.SetInfo
' 		ElseIf Err.Number=-2147024891 Then
'         	Trace Err.Number
'         	Trace "Ошибка. Скрипт должен быть запущен с административными полномочиями"
'         	Trace strEndScript
' 			WScript.Quit
' 		Else
' 			Trace "Ошибка. Неизвестная ошибка при создании локального пользователя"
' 			Trace strEndScript
' 			WScript.Quit
' 		End If
'     End If   
' End If
' #endregion


'файл отчета
	Set tf = objFSO.CreateTextFile(DATA_DIR & strPrefix & strInvNumber & DATA_EXT, True)
 
		tf.WriteLine "<?xml version=""1.0"" encoding=""windows-1251""?>"
		tf.WriteLine "<!--Файл инвентаризационных данных Рабочей станции " & strComputerName & ", Инв. номер " & strInvNumber & "  -->"
		tf.WriteLine "<!--Дата сбора информации " & Now & "  -->"
		tf.WriteLine "<InventoryData>"
'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в процедуре инициализации. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear

'Работаем с данными о локальных принтерах
tf.WriteLine "  <LocalPrinterInfo>"
Printer_Data_Write
tf.WriteLine "  </LocalPrinterInfo>"

'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в процедуре обработки LocalPrinterInfo. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear

'Пишем информацию по списку WMI классов
For Each WClass In arr_WMI_classes
tf.WriteLine "  <" & WClass & ">"
WMI_Data_Write WClass
tf.WriteLine "  </" & WClass & ">"
	'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в процедуре WMI_Data_Write. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear
Next


'Работаем с данными о мониторе
MonitorInfo
'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в функции Monitor_Info. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear

Monitor_Data_Write
'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в процедуре Monitor_Data_Write. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear

'Работаем с данными о установленных приложениях
InventSoft
'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в процедуре InventSoft. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear

tf.WriteLine "</InventoryData>"

Trace "Инвентаризация успешно завершена!"

Sub Printer_Data_Write
Dim intPrinterCount, strPrinterModel,strPrinterSerNumber, ii
Trace "Обработка LocalPrinterInfo"
intPrinterCount = PriterCount_check
If intPrinterCount > 0 Then
	ii = 0
	While ii<intPrinterCount
		strPrinterModel = InputBox ("Укажите модель для локального принтера №" & ii+1,TITLE)
		strPrinterSerNumber = InputBox ("Укажите серийный номер для локального принтера №" & ii+1,TITLE)
		substr = "      <N" & ii+1
		substr = substr & " Model=" & """" & RemoveSpecialChars(strPrinterModel) & """" & " SerNumber=" & """" & RemoveSpecialChars(strPrinterSerNumber) & """"
		tf.WriteLine substr & " />"
		ii=ii+1
	Wend  
End If
End Sub

Sub InventSoft

 Set wmiosoft = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\" & strComputerName & "\Root\default:StdRegProv")
Trace "Обработка Software"

	Dim ssoft
	ssoft = ExtractSoft("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\")
	'If Len(ssoft) > 0 Then tfsoft.Write ssoft

	'для 64-битных систем есть еще другой ключ! (32-битные программы на 64-битной системе)
	ssoft = ExtractSoft("HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\")
'	If Len(ssoft) > 0 Then tfsoft.Write ssoft
	
End Sub

Function ExtractSoft(key)

' 	'получить коллекцию
	Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
	Dim items
	wmiosoft.EnumKey HKLM, key, items
	If IsNull(items) Then
		ExtractSoft = ""
		Exit Function
	End If
 
	'отобрать нужные элементы
	Dim ssoft, itemsoft, oksoft, strDisplayName, strInstallDate, xsoft, prevsoft
	Dim strDisplayVersion, strInstallLocation, strPublisher
 	'ssoft = "" 'результат накапливать в строке
 	tf.WriteLine "  <InstalledSoft>"
	instance_number = 1
 	For Each itemsoft In items
 		oksoft = True 'флаг продолжения
 		'название, пропускать пустые и повторяющиеся
 		prevsoft = strDisplayName
 		wmiosoft.GetStringValue HKLM, key & itemsoft, "DisplayName", strDisplayName
		If IsNull(strDisplayName) Or Len(strDisplayName) = 0 Or strDisplayName = prevsoft Then
			oksoft = False
		Else
 		End If
 				
		'отделить заплатки, по значению параметра ParentKeyName = "OperatingSystem"
		'If oksoft Then
		'	wmiosoft.GetStringValue HKLM, key & itemsoft, "ParentKeyName", xsoft
		'		If IsNull(xsoft) Or xsoft <> "OperatingSystem" Then oksoft = False
		'End If

		'дата установки
		If oksoft Then
			wmiosoft.GetStringValue HKLM, key & itemsoft, "InstallDate", strInstallDate
			If IsNull(strInstallDate) Or Len(strInstallDate) < 8 Then
				strInstallDate = ""
			Else 'преобразовать в читаемый вид
				strInstallDate = Mid(strInstallDate, 7, 2) & "." & Mid(strInstallDate, 5, 2) & "." & Left(strInstallDate, 4)
			End If
			If InStr(strInstallDate,"-")<>0 or InStr(strInstallDate," ")<>0 Then strInstallDate = ""
		End If

		'производитель, версия, директория установки
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
'Проверка введенного количества принтеров
'===============================================================================================
Function PriterCount_check
Dim intPrinterCount
intPrinterCount = 0
tempbool = False
  Do Until tempbool = True 	  		
  intPrinterCount = InputBox ("Введите количество подключенных локальных принтеров",TITLE)
		     If Len(intPrinterCount) > 0 Then 			'количество символов номера инвентаризации
		     	If IsNumeric(intPrinterCount) Then    'Проверка того, является ли введенная информация числом.
		            tempbool = True 
		     	Else
		     		tmp = MsgBox("Ошибка: введены некорректные данные. Необходимо указать число!",16,TITLE)
		        	WScript.echo "Ошибка: введены некорректные данные. Необходимо указать число!"
		        End If
     		 Else
     		 	intPrinterCount = 0
     		 	tempbool = True
     End If
  Loop 
PriterCount_check = CInt(intPrinterCount)
'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в функции PriterCount_check. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear
End Function

'===============================================================================================
'Проверка введенного инвентарного номера и запись в реестр
'===============================================================================================
Sub INVNumber_check ()
tempbool = False
  Do Until tempbool = True 	  		
  strInvNumber = InputBox ("Введите инвентарный номер Компьютера." & vbCrLf & "4 Разряда" & vbCrLf & "Пример: 0138",TITLE)
		     If Len(strInvNumber)= 4 Then 			'параметр количество символов номера инвентаризации
		     	If IsNumeric(strInvNumber) Then    'Проверка того, является ли введенная информация числом.
		            tempbool = True 
		     	Else
		     		tmp = MsgBox("Ошибка: введены некорректные данные. Необходимо указать число!",16,TITLE)
		        	WScript.echo "Ошибка: введены некорректные данные. Необходимо указать число!"
		        	End If
     		 ElseIf Len(strInvNumber)= 0 Then
     		 	WScript.echo strEndScript
     		    WScript.Quit  		 
     		 Else
     		 tmp = MsgBox("Ошибка: количество разрядов введенного инвентарного номера не совпадает!",16,TITLE)
 	         WScript.echo "Ошибка: количество разрядов введенного инвентарного номера не совпадает!"
     End If
  Loop 
  On Error Resume Next
  objWshShell.RegWrite "HKEY_LOCAL_MACHINE\software\Inventory", strInvNumber, "REG_SZ" 'запись в реестр
' If Err.Number>0 Then
' WScript.Echo "Err.Number " + Err.Number
' ' 		If Err.Number=-2147022672 Then
' End If
'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в функции INVNumber_check. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear
'sreg = objWshShell.RegRead("HKEY_LOCAL_MACHINE\software\Inventory")
End Sub
'===============================================================================================
'Логирование
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
'Функция для замены спец символов
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
'Запись инвентаризационных данных WMI в xml файл
'WMI_class - имя собираемого класса
'в собираемые данные будут добавлены значения всех свойств класса
'===============================================================================================
Sub WMI_Data_Write(WMI_class)
	Const RETURN_IMMEDIATELY = 16
	Const FORWARD_ONLY = 32
	Dim query, classes, item, prop, value
	Trace "Обработка " & WMI_class
	query = "Select * From " & WMI_class

	Set classes = objWMIService.ExecQuery(query,, RETURN_IMMEDIATELY + FORWARD_ONLY)
	
	'tf.WriteLine "  <" & WMI_class & ">"
	instance_number = 1 'номер экземпляра
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
		'составить WQL-запрос, выполнить и записать строку в X-файл
		'входные параметры:
		'from - класс WMI
		'sel - свойства WMI, через запятую
		'where - условие отбора или пустая строка
		'sect - соответствующая секция отчета
		'param - соответствующие параметры внутри секции отчета, через запятую
		'для отображения в кратных единицах, нужно их указать в скобках
		
' 		Log "Win32_ComputerSystemProduct", _
' 		"UUID", "", _
' 		"Компьютер", _
' 		"UUID"
' 
' 	Log "Win32_ComputerSystem", _
' 		"Name,Domain,PrimaryOwnerName,UserName,TotalPhysicalMemory", "", _
' 		"Компьютер", _
' 		"Сетевое имя,Домен,Владелец,Текущий пользователь,Объем памяти (Мб)"
'sect - раздел типа компьютер
'все остальное в строку с именем номера экземпляра num 

		
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

	num = 1 'номер экземпляра
	For Each item In cls
	
	'tf.WriteLine "    <" & names(i) & ">"
	
		For i = 0 To UBound(props)
		If i =0 Then substr = "      <N" & num
			'взять значение
			Set prop = item.Properties_(props(i))
			value = prop.Value

			'без проверки на Null возможнен вылет с ошибкой
			If IsNull(value) Then
				value = ""

			'если тип данных - массив, собрать в строку
			ElseIf IsArray(value) Then
				value = Join(value,",")

			'если указана кратная единица измерения, перевести значение
			ElseIf Right(names(i), 4) = "(Мб)" Then
				value = CStr(Round(value / 1024 ^ 2))
			ElseIf Right(names(i), 4) = "(Гб)" Then
				value = CStr(Round(value / 1024 ^ 3))

			'если тип данных - дата, преобразовать в читаемый вид
			ElseIf prop.CIMType = 101 Then
				value = ReadableDate(value)
			End If

			'вывести в файл непустое значение, заменить спецсимвол ";"
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
		WScript.Echo "закрытие substr " & substr & " />"
	    'tf.WriteLine "    </" & names(i) & ">"
		'перейти к следующему экземпляру
		num = num + 1
	Next 'item
	tf.WriteLine "  </" & sect & ">"
	'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в процедуре Log. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
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
'Логирование ошибки		
If Err.Number <> 0 Then
   Trace "Ошибка в функции arr_WMI_classes_fill. Номер ошибки: " & Err.Number & ". Описание ошибки: " & Err.Description
End If
Err.Clear
End Function 

Function ReadableDate(str)
	'преобразование даты формата DMTF в читаемый вид (ДД.ММ.ГГГГ)
	'http://msdn.microsoft.com/en-us/library/aa389802.aspx
	ReadableDate = Mid(str, 7, 2) & "." & Mid(str, 5, 2) & "." & Left(str, 4)
End Function


Function BuildVersion()
	'узнать версию (билд) WMI-сервера
	'вернуть целое число
	Dim cls, item
	Set cls = wmio.ExecQuery("Select BuildVersion From Win32_WMISetting")
	For Each item In cls
		BuildVersion = CInt(Left(item.BuildVersion, 4))
	Next
	Trace.WriteLine Date & ":" & Time & ("версия WMI прошла")
End Function


Function Unavailable(addr)

	'проверить доступность компьютера в сети
	'вернуть True, если адрес недоступен
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
	text_log.WriteLine Date & ":" & Time & ("пинг локальной машины для последуюшего сохранения отчета")
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

Trace "Обработка MonitorInfo"

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
' 		'производитель
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
' 'проверить доступность компьютера в сети; вернуть True, если адрес недоступен
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
