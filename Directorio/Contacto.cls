VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command

Public Sub Conectar()
cn.Open StrCon
End Sub

Public Sub Desconectar()
cn.Close
Set cn = Nothing
End Sub

Public Function LeerContactosxLetra(xLetra As String) As Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Contactos WHERE LEFT(Apel,1)='" & xLetra & "' ORDER BY 2", StrCon, adOpenStatic
Set LeerContactosxLetra = rs
Set rs = Nothing
End Function

Public Function LeerContactos() As Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Contactos ORDER BY 1", StrCon, adOpenStatic
Set LeerContactos = rs
Set rs = Nothing
End Function


Public Function BuscarContactosxCateg(xCat As Byte) As Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM Contactos WHERE IdCat=" & xCat & "", StrCon, adOpenDynamic, adLockReadOnly
Set BuscarContactosxCateg = rs
Set rs = Nothing
End Function
Public Sub GuardarContacto(ByVal xCod As String, ByVal xNomb As String, ByVal xApel As String, ByVal xDirec As String, ByVal xCodDist As Byte, ByVal xTelf As String, ByVal xCelu As String, ByVal xEmail As String, ByVal xFecNac As Date, ByVal xCodCat As Byte)
Conectar
cn.Execute "INSERT INTO Contactos VALUES('" & Trim(xCod) & "','" & Trim(xNomb) & "','" & Trim(xApel) & "','" & Trim(xDirec) & "'," & xCodDist & ",'" & Trim(xTelf) & "','" & Trim(xCelu) & "','" & Trim(xEmail) & "','" & xFecNac & "'," & xCodCat & ")"
Desconectar
End Sub

Public Sub ModificarCateg(ByVal xIdCon As String, ByVal xIdCat As Byte)
Conectar
cn.Execute "UPDATE Contactos SET IdCat='" & xIdCat & "' WHERE IdContacto='" & xIdCon & "'"
Desconectar
End Sub
