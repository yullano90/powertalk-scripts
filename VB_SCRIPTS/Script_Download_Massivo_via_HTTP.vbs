MsgBox "AUTOMACAO CRIADA POR IULANO.SANTOS@Empresa_X.COM"

Set objNetwork = CreateObject("WScript.Network")
strUserName = objNetwork.UserName
User  = strUserName

Function splash_SISTEMA_X
	MsgBox "MATRICULA Empresa_X AUTENTICADA!: " & strUserName, vbInformation, "SISTEMA_X AUTOMATION DATA COLLECTION Empresa_X API"
End Function

'wscript.echo "VOCE EXECUTARÁ O DOWNLOAD MASSIVO OS ARQUIVOS DE SEU INTERESSE! PORTANTO APONTE PARA O DIRETÓRIO CORRETO!"
 
splash_SISTEMA_X

filename = "nomeDoTeuArquivo.csv"
directory = "C:\teuDiretorio\"
 
Dim oXMLHTTP
Dim oStream
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename)
Dim line,OM,LINK_API,IngnoreLine
IngnoreLine = "URL"
 
Do Until f.AtEndOfStream
    line  = f.ReadLine
    a = Split(line,",")
    OM = a(0)
    LINK_API =  a(1)
    If LINK_API <> IngnoreLine Then
        oXMLHTTP.Open "GET", LINK_API, False
        oXMLHTTP.Send
            If oXMLHTTP.Status = 200 Then
                Set oStream = CreateObject("ADODB.Stream")
                oStream.Open
                oStream.Type = 1
                oStream.Write oXMLHTTP.responseBody
                oStream.SaveToFile  directory & OM & ".png"
                oStream.Close
            End If
    End If
Loop
 
f.Close