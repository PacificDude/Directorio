Attribute VB_Name = "Funciones"
Dim rs As New ADODB.Recordset

Public Function CreaCodigo(ByVal xCampo As String, ByVal xTabla As String) As Variant
rs.Open "SELECT max(val(" & xCampo & ")) FROM " & xTabla & "", StrCon, adOpenStatic
If IsNull(rs(0)) Then
    CreaCodigo = "1"
Else
    CreaCodigo = rs(0) + 1
End If
End Function

