VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Distrito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Public Function LeerDistritos() As Recordset
If rs.State = 1 Then rs.Close
rs.Open "SELECT IdDist,Nomb FROM Distritos ORDER BY 2", StrCon, adOpenStatic
Set LeerDistritos = rs
Set rs = Nothing
End Function
