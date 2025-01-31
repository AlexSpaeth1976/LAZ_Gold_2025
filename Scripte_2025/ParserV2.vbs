' Region Description
' *****************************************************************************
' *** File:Parser.vbs   				                                     ***
' *****************************************************************************
' *** Homag Holzbearbeitungssysteme AG                                      ***
' *** Author:  (c) Alexander Späth                                          ***
' *** Date:    01.07.2003 					                                ***
' *** Version: 1.00	                                                        ***
' *****************************************************************************
'******************************************************************************
' History
' Version 1.00  SPX  03.04.2008 
' Version 2.00  SPX  31.01.2025 
'******************************************************************************
' EndRegion
Option Explicit
' Titel, Scriptversion
Dim Titel, Version
Version = "V2.00"
Titel = "Parser.vbs " & Version

Dim wsh,Net,fso,objConfiguration
Dim gsScriptBase
Dim i,k

Dim key,Items 
Dim strSplittAntwort
Dim intAnzahlAntworten
Dim strLine,strAnworten
Dim strKontrolle
Dim strTempAntwortString
Dim strTXTEin
Dim strHTMLOut

gsScriptBase = left(wscript.scriptfullname,len(wscript.scriptfullname)-3)  

Set wsh = WScript.CreateObject("WScript.Shell")
Set Net = CreateObject("WScript.Network")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set objConfiguration = WScript.CreateObject("Scripting.Dictionary")


'strTXTEin =  SelectFileDialog ("c:\","Bitte Datei die eingelesen werden soll auswählen...",".txt")
'strHTMLOut = SelectFileDialog ("c:\","Bitte Ausgabedatei auswählen...",".txt")

' Datei öffnen
strTXTEin = BrowseForFile
strHTMLOut = strTXTEin + "_MOD"



ReadConfig strTXTEin


'frage[0] = new Aufgabe("Wer ist nach dem Feuerwehrgesetz Baden-W&uuml;rttemberg f&uuml;r die Aufstellung, Ausr&uuml;stung und Unterhaltung der Feuerwehr verantwortlich?",5,"Bund","Land","Kreis","Gemeinde","Kommandant");
'richtig[0] = new Kontrolle(0,0,0,1,0);


key = objConfiguration.Keys
items = objConfiguration.items

For i=0 To objConfiguration.count -1
	strLine = ""
	strAnworten = ""
	strKontrolle =""
	
	'MsgBox "Key:" & key(i) & vbNewLine & "VAlue: " &items(i) 
	strSplittAntwort = Split(items(i),vbNewLine)
	intAnzahlAntworten = UBound(strSplittAntwort)
	For k = 0 To intAnzahlAntworten	
		strTempAntwortString = Right(strSplittAntwort(k),5)
		If (InStr (1,strTempAntwortString,"!!O!!",1) > 0) Then
			If k = intAnzahlAntworten Then
				strKontrolle = strKontrolle & "0"
			Else
				strKontrolle = strKontrolle & "0,"
			End If
		ElseIf (InStr (1,strTempAntwortString,"!!X!!",1) > 0) Then
			If k = intAnzahlAntworten Then
				strKontrolle = strKontrolle & "1"
			Else
				strKontrolle = strKontrolle & "1,"
			End If
		Else
			If k = intAnzahlAntworten Then
				strKontrolle = strKontrolle & "0"
			Else
				strKontrolle = strKontrolle & "O,"
			End If
		End If
			
		If k = intAnzahlAntworten Then
			strAnworten = strAnworten & """"& Replace(Replace(strSplittAntwort(k),"!!O!!",""),"!!X!!","") &""""
		Else
			strAnworten = strAnworten & """"& Replace(Replace(strSplittAntwort(k),"!!O!!",""),"!!X!!","") &"""" & ","	
		End If
		
	Next
	strLine = "frage["&i&"] = new Aufgabe(" &"""" & key(i)&""""&","&(intAnzahlAntworten+1)&","& strAnworten & ");"
	WriteLine strLine,strHTMLOut
	
	strLine = ""
	strLine = "richtig["&i&"] = new Kontrolle(" & strKontrolle & ");"
	WriteLine strLine,strHTMLOut
Next

Function BrowseForFile()
    With CreateObject("WScript.Shell")
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
        Dim tempName : tempName = fso.GetTempName() & ".hta"
        Dim path : path = "HKCU\Volatile Environment\MsgResp"
        With tempFolder.CreateTextFile(tempName)
            .Write "<input type=file name=f>" & _
            "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
            ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
            "close();</script>"
            .Close
        End With
        .Run tempFolder & "\" & tempName, 1, True
        BrowseForFile = .RegRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder & "\" & tempName
    End With
End Function

'*************************************************************************************************************************
' Function SelectFileDialog
'*************************************************************************************************************************
Function SelectFileDialog (strPathAndFilename,strTitel,strFilter)
	Dim intPos1, intPos2, FileName,objDialog,intResult
	'*********************************************
	Set objDialog = CreateObject("UserAccounts.CommonDialog")
	objDialog.Filter = strFilter
	objDialog.FilterIndex = 1
	objDialog.FileName = strTitel
	'objDialog.Titel = strTitel
	objDialog.InitialDir = strPathAndFilename
	intResult = objDialog.ShowOpen

	If objDialog.Filename = "" Or intResult = vbFalse Then
		WScript.Quit(5)
	End If
 	strPathAndFilename = LCase (objDialog.Filename)
	strPathAndFilename = Replace (strPathAndFilename,"/","\")
'	intPos1 = (InstrRev (strPathAndFilename, "\", -1))
'	intPos2 = (Len (strPathAndFilename))
'	intPos2 = intPos2 - intPos1
'	FileName = Mid(strPathAndFilename, intPos1 + 1, intPos2) 
'	SelectDiaginfo = FileName
	SelectFileDialog = strPathAndFilename
	
End Function

'*****************************************************************************
' Function WriteLine(strFileName,strLine)													 *
' 																			 *
'*****************************************************************************
Function WriteLine(strLine,strfile)
Dim objFile, objFso
Dim FileOut
Dim Scriptpfad

' Ermittle den Pfad zum Skript
Scriptpfad = WScript.ScriptFullName
Scriptpfad = Left(ScriptPfad, InStrRev(ScriptPfad, "\"))

Set FileOut = fso.OpenTextFile(strfile, 8, True)
FileOut.WriteLine(strLine)

Set FileOut = Nothing 

End Function

'*****************************************************************************
' Function ReadConfig														 *
' 																			 *
'*****************************************************************************
Function ReadConfig (strFile)
	Dim objFile, objFso
	Dim readline,strArrayReadOnline
	Dim strArrKonfig
	Dim Scriptpfad
	Dim strText,i,objTemp
	Dim objNeu
	
	' Ermittle den Pfad zum Skript
	Scriptpfad = WScript.ScriptFullName
	Scriptpfad = Left(ScriptPfad, InStrRev(ScriptPfad, "\"))
		
	Set objFso = CreateObject("Scripting.FileSystemObject")  	
	Set objFile = objFso.OpenTextFile(strFile)  
	
	strArrayReadOnline = objFile.ReadAll   ' Datei lesen
	objFile.Close                          ' Datei schließen

	readline = Split(strArrayReadOnline,vbNewLine & vbNewLine & vbNewLine)
	For i=0 To UBound(readline)
		If (InStr(1, readline(i), "(*", 1) = 0) Then
			strArrKonfig = Split(readline(i),vbNewLine & vbNewLine)
			
			If UBound(strArrKonfig) >= 1 Then
				' Frage, Antworten
				objConfiguration.Add strArrKonfig(0),strArrKonfig(1)
			End If
		End If
	Next
End Function