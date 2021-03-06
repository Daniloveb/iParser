'============================================================================
'Инвентаризация компьютеров средствами WMI и WSH (23.03.2012)
'Ручное создание конфигурационных единиц
'Данные выводятся в отдельный файл XML для последующей обработки сервисом iParser
'Автор: Данилов Евгений lanc@live.ru
'Версия 1.01
'Вид запуска - cscript CreateKE.vbs
'============================================================================
Option Explicit	
'On Error Resume Next

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
'Dim arr_WMI_classes 'Массив собираемых WMI-классов
Dim DATA_DIR
'Dim AUTO 'При автоматическом запуске AUTO=True
'Dim Subnets 'Массив подсетей
Dim strPath 'Путь к директории сохранения результатов
'Dim isDomain 'При запуске локальным пользователем isDomain=False
'Dim strComputerName, strUserName
Dim args, objFSO, objWshShell, objNetwork, objWMIService, objSID, objGroup
'Dim objUser, objComputer, WClass 
'Dim x,y,i, xsub, , instance_number
Dim tmp, strEndScript
Dim strInvNumber 'инвентарный номер
Dim tf 'файл отчета
Dim text_log ' файл отчета ошибок
Dim DeviceType
Dim Prefix
Dim strMAC
Dim strIpAddress
Dim strModel, strSerialNumber

'Определение глобальных объектов
Set args = WScript.Arguments
Set objWshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'============================================================================
'Присвоение глобальных переменных
'============================================================================
strEndScript = "Скрипт инвентаризации завершает работу"
DATA_DIR = "DATA\" 'локальный каталог сохранения отчета - расположение по умолчанию, если не указано параметром
'============================================================================
'Проверка на использование cscript
'============================================================================
' If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
' 'UserInput = WScript.StdIn.ReadLine
' Else
' tmp = MsgBox("Скрипт требуется запустить с использованием cscript. Строка запуска - cscript CreateKE.vbs",16, TITLE)
' WScript.Quit
' End If
'============================================================================
'Определение вида KE
'============================================================================
tmp = "========================================================================" & vbCrLf & _
"Укажите тип Конфигурационной Единицы! " & vbCrLf & "Введите номер из списка:" & vbCrLf & "1 - Сетевой принтер" & vbCrLf & _
"2 - Преобразователь в Ethernet" & vbCrLf & _
"========================================================================"
WScript.echo tmp
DeviceType = InputBox ("Укажите тип Конфигурационной Единицы! " & vbCrLf & "Введите номер из списка:" & vbCrLf & "1 - Сетевой принтер" & vbCrLf & _
"2 - Преобразователь в Ethernet" & vbCrLf & "3 - Станочное Оборудование" & vbCrLf & "4 - Рабочая станция" & vbCrLf & "5 - Тонкий клиент" & vbCrLf & "6 - Прочее сетевое оборудование")
Select Case DeviceType
Case 1
Prefix = "R"
Case 2
Prefix = "E"
Case 3
Prefix = "S"
Case 4
Prefix = "M"
Case 5
Prefix = "T"
Case 6
Prefix = "L"
Case Else
tmp = MsgBox("Ошибка ввода! Укажите код типа оборудования.",16,TITLE)
WScript.echo strEndScript
WScript.Quit
End Select
strInvNumber = INVNumber_check
strMAC = MAC
strIpAddress = IpAddress
strModel = Model
strSerialNumber = SerNumber
'============================================================================
'Запрос модели, MAC-адреса и серийного номера
'============================================================================

'файл отчета

If objFSO.FolderExists(DATA_DIR) Then
Else
objFSO.CreateFolder(DATA_DIR)
End if

	Set tf = objFSO.CreateTextFile(DATA_DIR & Prefix & strInvNumber & DATA_EXT, True)
 
		tf.WriteLine "<?xml version=""1.0"" encoding=""windows-1251""?>"
		tf.WriteLine "<!--Файл инвентаризационных данных Конфигурационной Единицы, Инв. номер " & strInvNumber & "  -->"
		tf.WriteLine "<!--Дата сбора информации " & Now & "  -->"
		tf.WriteLine "<InventoryData>"
		tf.WriteLine "  <NetworkDeviceInfo>"
		tf.WriteLine "<N1" & " MACAdress=" & """" & RemoveSpecialChars(strMAC) & """" & " IpAddress=" & """" & RemoveSpecialChars(strIpAddress) & """" & " Model=" & """" & RemoveSpecialChars(strModel) & """" & " SerialNumber=" & """" & RemoveSpecialChars(strSerialNumber) & """" & " />"
		'substr = substr & " MACAdress=" & """" & RemoveSpecialChars(strMACAdress) & """" & " Model=" & """" & RemoveSpecialChars(strModel) & """" & " SerialNumber=" & """" & RemoveSpecialChars(strSerialNumber) & " />"
		tf.WriteLine "  </NetworkDeviceInfo>"
		tf.WriteLine "</InventoryData>"
		
		
Function Model
	Model = InputBox ("Укажите модель устройства!",TITLE)
End Function
		
Function SerNumber
	SerNumber = InputBox ("Укажите Серийный Номер устройства!",TITLE)
End Function
		
Function IpAddress
	IpAddress = InputBox ("Укажите IpAddress устройства!",TITLE)
End Function
		
Function MAC
Dim tempbool
tempbool = False
Do Until tempbool = True
	strMAC = InputBox ("Укажите MAC-Адрес, в виде 00-23-8B-FF-E5-60." & vbCrLf & "Используйте заглавные латинские буквы.",TITLE)
	If Len(strMAC) = 17 Then
	tempbool = True
	ElseIf Len(strMAC)= 0 Then
     		 	WScript.echo strEndScript
     		    WScript.Quit  		
	Else
		tmp = MsgBox("Ошибка: введены некорректные данные. Проверьте длину выражения!",16,TITLE)
		WScript.echo "Ошибка: введены некорректные данные. Проверьте длину выражения!"
	End If
Loop
MAC = strMAC
End Function

Function INVNumber_check
Dim tempbool
tempbool = False
  Do Until tempbool = True 	  		
  strInvNumber = InputBox ("Введите инвентарный номер." & vbCrLf & "4 Разряда" & vbCrLf & "Пример: 0138",TITLE)
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
  INVNumber_check = strInvNumber
End Function

Function RemoveSpecialChars(str)
Dim endstr
Dim i
endstr = str
For i = 1 To Len(endstr)
	tmp = Mid(endstr,i,1)
	If Asc(tmp) < 32 Then tmp = ""
	i=i+1
Next
endstr = Trim(Replace(endstr, ";", "_"))
endstr = Trim(Replace(endstr, """", "&quot;"))
endstr = Trim(Replace(endstr, "&", "&amp;"))
RemoveSpecialChars = endstr
End Function
