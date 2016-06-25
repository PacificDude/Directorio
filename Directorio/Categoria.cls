VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Categoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Public Sub Conectar()
cn.Open StrCon
End Sub

Public Sub Desconectar()
cn.Close
Set cn = Nothing
End Sub

Public Function LeerCategorias() As Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT IdCat,Nomb FROM Categorias ORDER BY 2", StrCon, adOpenStatic
Set LeerCategorias = rs
Set rs = Nothing
End Function


Public Sub GuardarNuevaCategoria(ByVal xNomb As String)
Dim xCod As Byte
Conectar
xCod = Val(CreaCodigo("IdCat", "Categorias"))
cn.Execute "INSERT INTO Categorias VALUES (" & xCod & ",'" & xNomb & "')"
Desconectar
End Sub

