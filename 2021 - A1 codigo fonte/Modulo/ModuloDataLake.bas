Attribute VB_Name = "ModuloDataLake"
Dim dataLakeDb      As New ADODB.Connection
Public Function cnDatabaseLake() As Boolean
    On Error GoTo TrtErroDB
    
    Dim cnString        As String
    'MySQL  |
    
    cnString = "driver={MySQL ODBC 5.1 Driver};Server=" & srv_IP & ";port=" & srv_Porta & ";uid=" & dbUsu & ";pwd=" & dbSenha & ";database=datalake"


    dataLakeDb.CursorLocation = adUseClient  'usamos um cursor do lado do cliente pois os dados 'serao acessados na maquina do cliente e nao de um servidor

    dataLakeDb.Open cnString

    
    Exit Function
TrtErroDB:
     RegLogDataBase 0, 0, "cnDatabaseLake", Err.Number & " - " & Err.Description
     Exit Function
 End Function



Public Function dataLakeInputNFe(numNFe As String, chvA As String, strData As String) As Long
    
    On Error GoTo TrataErro
    Dim sSQL       As String
    Dim sFields    As String
    Dim sValues    As String

    'Open database
    cnDatabaseLake
    

    sTabela = LCase("nfe")
    
    sFields = "DtHr,Id_Empresa,UsuID"
    sValues = "'" & Now() & "','" & ID_Empresa & "'," & ID_Usuario
    
    
    sFields = sFields & ", tipo, numNFe, chv, txtNFe"
    sValues = sValues & ",1,'" & numNFe & "','" & chvA & "' , '" & strData & "'"
    

    sSQL = "INSERT INTO " & LCase(sTabela) & _
           " (" & sFields & ") " & _
           " VALUES(" & sValues & ") "
     
     
    
    'sSQL = LCase(sSQL)
    
    Dim Rst        As ADODB.Recordset

    Set Rst = New ADODB.Recordset
    
    Rst.Open sSQL, dataLakeDb
    
    'close database
    dataLakeDb.Close
    Exit Function

TrataErro:
    MsgBox "function dataLakeInputNFe: " & Err.Description, vbInformation, Err.Number
    RegLogDataBase 0, 0, 0, "function dataLakeInputNFe: " & Err.Number & "-" & Err.Description
    
    Exit Function
End Function

