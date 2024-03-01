Rem ==============================================================================================
Rem                    SCRIPT DE CAPTURA DE DADOS DA TRANSACAO SAP IW38
Rem                    DESENVOLVEDOR E SYSADMIN: IULANO SANTOS
Rem ===============================================================================================
 
Option Explicit

Dim ObjShell
Dim SapApplication
Dim SapPath
Dim FilePath
Dim LogPath
Dim FileName
Dim SapGuiAuto
Dim Connection
Dim Session
Dim SapConnectionName
Dim User
Dim SapUsr
Dim Pwd
Dim FirstDate
Dim ReceiverAddress
Dim TimeInMinutes
Dim Logged
Dim MonthTarget 
Dim MonthCount
Dim Attempt
Dim objNetwork
Dim strUserName
Dim WorkCenter
Dim WorkCenterInput
Dim homeDate
Dim homeDateInput
Dim lastDate
Dim lastDateInput
Dim userInput
Dim directoryReport
Dim  ArraySapTransactions(10)


MsgBox "AUTOMACAO CRIADA POR IULANO.SANTOS@Empresa_X.COM"
Set objNetwork = CreateObject("WScript.Network")
strUserName = objNetwork.UserName
userInput = "#################"
User  = strUserName


'userInput = InputBox("INFORME SUA S3NH@ SAP, SUA MATRICULA JA LOCALIZAMOS NO WINDOWS: " & strUserName, "SISMF_COLETA_PWD_IW38")
'WorkCenterInput = InputBox("INFORME O CENTRO DE TRABALHO: " & strUserName, "SISMF_COLETA_CENTRO_DE_TRABALHO_IW38")
'homeDateInput = InputBox("INFORME A DATA INICIAL: " & strUserName, "SISMF_COLETA_DATA_INICIAL_IW38")
'lastDateInput = InputBox("INFORME A DATA FINAL: " & strUserName, "SISMF_COLETA_DATA_FINAL_IW38")
WorkCenter = "ECEE10*"
homeDate = "01.01.2022"
lastDate = "12.12.2023"
SapUsr  = strUserName
Pwd	= userInput


SapPath = """C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"""
directoryReport = "C:\Users\"&strUserName&"\OneDrive - Empresa_X S.A\SISMF\SAP_Datasets"
ReceiverAddress = "http://IP/Site/Controller/Action"
SapConnectionName = "01 - ECC - Production - EP0"	
FileName= "export.xlsx"
LogPath = ""
TimeInMinutes = 5
Logged = False
Attempt = 4

 ArraySapTransactions(0) = "IW38"
 ArraySapTransactions(1) = "IW28"
 ArraySapTransactions(2) = "ME2N"
 ArraySapTransactions(3) = "ME5A"


Function splash_SISMF
	MsgBox "MATRICULA Empresa_X AUTENTICADA: " & strUserName, vbInformation, "SIMF AUTOMATION SAP DATA"
End Function
 
splash_SISMF

 
If Len(User) > 0  And Len(TimeInMinutes) > 0 Then
	Call Main()
Else
	MsgBox("Verifique os Dados!")
	WScript.Quit
End If
 

 
Sub OpenSap()
	'Open the sap logon screen
	On Error Resume Next
	If Not IsObject(SapApplication) Then
		Set ObjShell = CreateObject("WScript.Shell")
		ObjShell.Run SapPath 	
		WScript.sleep 4000	
		Set ObjShell = Nothing	
	End If	
	If Err.Number <> 0 Then
		Log Now() & " Erro ao abrir o SAP (" & Err.Number & " ) : " & Err.Description 
		WScript.Quit
	End If 
	If Not IsObject(SapApplication) Then	
		Set SapGuiAuto = GetObject("SAPGUI")	
		Set SapApplication = SapGuiAuto.GetScriptingEngine()
	End If		
	If Not IsObject(Connection) Then
	   Set Connection = SapApplication.OpenConnection(SapConnectionName, True)	
	End If
 
	If Not IsObject(Session) Then
	   Set Session = Connection.Sessions(0)	
	End If
 
	If IsObject(WScript) Then
	   WScript.ConnectObject Session,     "on"
	   WScript.ConnectObject SapApplication, "on"
	End If
 
	If Len(User) > 0 And Len(Pwd) > 0 Then
		FilePath = "C:\Users\" & User & "\Documents"
		LogPath = "C:\Users\" & User & "\Documents\logs"
		Session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SapUsr
		Session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = Pwd		
		Session.findById("wnd[0]").sendVKey 0	
		If(Session.ActiveWindow.Name = "wnd[1]") Then	
			Dim lvCount			
			If(InStr(session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT").text, SapUsr) > 0 And _
			InStr(session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT").text, "logged") > 0) Then
				Log "O usuario " & SapUsr & " já estava logado, rodando novamente em " & TimeInMinutes	& " minutos"			
				session.findById("wnd[1]").close()
				session.findById("wnd[0]").close()								
				lvCount = 0
				Do While (Session.ActiveWindow.Name = "wnd[1]")
					Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press			
					lvCount = lvCount + 1
					If lvCount >= Attempt Then
						Exit do
					End If
				Loop					
				Logged = False
			Else
				lvCount = 0
				On Error Resume Next
				Do While (Session.ActiveWindow.Name = "wnd[1]")
					session.findById("wnd[1]/tbar[0]/btn[0]").press					
					lvCount = lvCount + 1
					If lvCount >= Attempt Then
						Exit do
					End If
				Loop
				If lvCount > 0 Then
					Logged = True
				Else
					Logged = False 
				End If
			End If
		Else			
			Logged = True
		End If
	Else 
		Log "Error on length of User or Password!"
		MsgBox("Erro no usuario ou senha!")
		WScript.Quit
	End If
End Sub



Sub CloseSap()
	Dim lvCount
	lvCount = 0
	On Error Resume Next
	If (Logged = True) Then
		Session.findById("wnd[0]").close()	
		Do While (Session.ActiveWindow.Name = "wnd[1]")
			Session.findById("wnd[1]/usr/btnSPOP-OPTION1").press			
			lvCount = lvCount + 1
			If lvCount >= Attempt Then
				Exit do
			End If
		Loop		
	End If
	Err.Clear
	Session = Empty
	Connection = Empty
	SapApplication = Empty
	SapGuiAuto = Empty
	ObjShell = Empty
	Logged = False
End Sub




Rem ===============================================================================================
Rem                 COMANDOS DA TRANSAÇÃO IW38 PARA UPDATE DE DADOS SAP NO CRC - SISMF
Rem ===============================================================================================

Function get_Data_IW38_SISMF
	Session.findById("wnd[0]").maximize
	session.findById("wnd[0]/tbar[0]/okcd").text = "IW38"
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/usr/chkDY_MAB").selected = true
	session.findById("wnd[0]/usr/chkDY_HIS").selected = true
	session.findById("wnd[0]/usr/ctxtGEWRK-LOW").text = WorkCenter
	session.findById("wnd[0]/usr/ctxtDATUV").text = homeDate
	session.findById("wnd[0]/usr/ctxtDATUB").text = lastDate
	session.findById("wnd[0]/usr/ctxtDATUV").setFocus
	session.findById("wnd[0]/usr/ctxtDATUV").caretPosition = 10
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1,""
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = directoryReport    
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "NEW SAP.xlsx"
	session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
	session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 25
	session.findById("wnd[1]/tbar[0]/btn[11]").press
End Function



Rem ===============================================================================================
Rem                 COMANDOS DA TRANSAÇÃO ME5A PARA COLETA DE DADOS DE PEDIDOS
Rem ===============================================================================================

Function get_Data_ME5A_SISMF
	session.findById("wnd[0]").maximize
	session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00010"
	session.findById("wnd[0]/usr/ctxtP_LSTUB").text = "ALV"
	session.findById("wnd[0]/usr/ctxtP_LSTUB").setFocus
	session.findById("wnd[0]/usr/ctxtP_LSTUB").caretPosition = 3
	session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "4050"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "4064"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "4065"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1058"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "1064"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 0
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/tbar[0]/btn[8]").press
	session.findById("wnd[0]/usr/ctxtS_LFDAT-LOW").text = "01.12.2023"
	session.findById("wnd[0]/usr/ctxtS_LFDAT-HIGH").text = "31.12.2023"
	session.findById("wnd[0]/usr/ctxtS_LFDAT-HIGH").setFocus
	session.findById("wnd[0]/usr/ctxtS_LFDAT-HIGH").caretPosition = 10
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = 1
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ERNAM"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "AFNAM"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WERKS"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LFDAT"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "KNTTP"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BANFN"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BNFPO"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BADAT"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "FRGDT"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELN"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EBELP"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BEDAT"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MENGE"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MATNR"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "TXZ01"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "MEINS"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "PREIS"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "GSWRT"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "ZZJ_1BNBM"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "FNAME1"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BSART"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "LOEKZ"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EKGRP"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "EKORG"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "KONNR"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "WAERS"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "IDNLF"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "PLIFZ"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "SRM_CONTRACT_ID"
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
	session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/tbar[0]/btn[0]").press
End Function


Sub Main()
    OpenSap
	If (Logged = True) Then

		get_Data_IW38_SISMF
		get_Data_ME5A_SISMF

    End If
	CloseSap
End Sub


Rem ==============================================================================================
Rem               JAMAIS DESISTA DOS SEUS SONHOS, PERSISTA, VAI DAR CERTO!!!
Rem ==============================================================================================