' ClearOfficeMRU v 1.0 (c) Damian Daszkiewicz
' Skrypt jest darmowy, wyrażam zgodę na dokonywanie drobnych modyfikacji, tak aby skrypt lepiej odpowiadał Twoim potrzebom
' www.OfficeBlog.pl
'
' Opis skryptu:
' https://www.officeblog.pl/automatyczne-usuwanie-listy-ostatnio-uzywanych-dokumentow-mru/

Option Explicit

Const HKEY_CURRENT_USER = &H80000001
Dim tStart, Info, tTime
Dim strComputer
Dim oReg
Dim OffVer, OffApp, OffDir, MRUKey
Dim iVer, iApp, iDir, i
Dim Custom

tStart=Timer
strComputer = "."
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

' ------------------------ FUNKCJE -------------------------------
Sub ClearMRU(MRUKey)
	Dim arrValueNames, arrValueTypes
	Dim i
	Dim itemVal
	
    oReg.EnumValues HKEY_CURRENT_USER, MRUKey, arrValueNames, arrValueTypes
    
    If IsNull(arrValueNames) Then
        'Klucz nie istnieje, nic nie rób
    Else
        For i = 0 To UBound(arrValueNames)
            If arrValueNames(i) <> "Max Display" Then
                itemVal = arrValueNames(i)
                'If Left(itemVal, 4) = "Item" Then 
                oReg.DeleteValue HKEY_CURRENT_USER, MRUKey, itemVal
            End If
        Next
    End If
End Sub


Sub OfficeKeyUserMRU(OffVersion, sApp)
	Dim strKeyPath, arrSubKeys, subkey, tmpKey
	
	strKeyPath = "Software\Microsoft\Office\" + OffVersion + "\" + sApp + "\User MRU"
	oReg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrSubKeys
	
	If IsNull(arrSubKeys) Then
		'Klucz nie istnieje, nic nie rób
	Else	    
		For Each subkey In arrSubKeys
		    tmpKey = strKeyPath + "\" + subkey + "\"
		    Call ClearMRU(tmpKey + "File MRU")
		    Call ClearMRU(tmpKey + "Place MRU")
		Next
	End If
End Sub


Sub AdobeAcrobatReaderDC_go(strKeyPath)
	Dim tmpKey, arrSubKeys, subkey
	oReg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrSubKeys
	
	If IsNull(arrSubKeys) Then
		'Klucz nie istnieje, nic nie rób
	Else	    
		For Each subkey In arrSubKeys
		    tmpKey = strKeyPath + subkey + "\"
		    oReg.DeleteKey HKEY_CURRENT_USER, tmpKey
		Next
	End If    
End Sub



Sub AdobeAcrobatReaderDC()
	Dim tmpKey
    tmpKey="Software\Adobe\Acrobat Reader\DC\AVGeneral\"
    AdobeAcrobatReaderDC_go(tmpKey + "cRecentFiles\")
    AdobeAcrobatReaderDC_go(tmpKey + "cRecentFolders\")
End Sub


Sub AddToAutostart()
	Dim strValueName, szValue, strKeyPath
	strValueName = WScript.ScriptName
	szValue = WScript.ScriptFullName
	strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Run"
	
	oReg.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, szValue
End Sub

' ------------------GŁÓWNY KOD----------------------------------


' 12.0 - Office 2007
' 13.0 - nie istnieje, w Microsofcie ludzie czczą zabobony ;-)
' 14.0 - Office 2010
' 15.0 - Office 2013
' 16.0 - Office 2016, 2019 365
OffVer = Array("12.0", "14.0", "15.0", "16.0")
OffApp = Array("Access", "Excel", "PowerPoint", "Publisher", "Word")
OffDir = Array("File MRU", "Place MRU", "Recent File List") ', "Security\Trusted Documents\TrustRecords")


' Główna pętla - tutaj w zagnieżdżonej pętli szukamy odpowiednich kluczy w rejestrze i je czyścimy
For iVer = 0 To UBound(OffVer)
    For iApp = 0 To UBound(OffApp)
        For iDir = 0 To UBound(OffDir)
            MRUKey = "Software\Microsoft\Office\" + OffVer(iVer) + "\" + OffApp(iApp) + "\" + OffDir(iDir)
            Call ClearMRU(MRUKey)
        Next
    Next
Next


'W wersji 16.0 jest dodatkowy klucz UserMRU a w nim podklucze
For i=0 to UBound(OffApp)
	Call OfficeKeyUserMRU("16.0", OffApp(i))
Next


'To do kompletu dodajmy jeszcze AdobeReadera ;-)
Call AdobeAcrobatReaderDC


' Jak chcesz na własne potrzeby dopisać na sztywno jakąś gałąź rejestru z której usuwasz elementy możesz to zrobić tutaj
Custom = Array("Software\Microsoft\Windows\CurrentVersion\Applets\Wordpad\Recent File List", _
"Software\Microsoft\Windows\CurrentVersion\Applets\Paint\Recent File List", _
"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU")

For i=0 To UBound(Custom)
	ClearMRU(Custom(i))
Next

' Jeśli chcesz, aby program był w autostarcie to odkomentuj poniższą linijkę przed pierwszym uruchomieniem
' Call AddToAutostart

tTime=FormatNumber(Timer-tStart, 2)
Info = "Wyczyściłem elementy MRU. Zajęło mi to: " + tTime + " sekund"
Info = Info + vbCrLf + vbCrLf + "Autorem programu jest Damian Daszkiewicz" + vbCrLf + "Zapraszam na www.OfficeBlog.pl"
MsgBox Info, vbInformation, "Clear Office MRU v 1.0"