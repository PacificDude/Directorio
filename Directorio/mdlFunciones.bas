Attribute VB_Name = "mdlFunciones"

Public Sub Limpiar(ByVal frm As Form)
Dim x As Control
For Each x In frm.Controls
    If TypeOf x Is TextBox Then x.Text = Empty
    If TypeOf x Is DataCombo Then x.Text = Empty
Next
End Sub


'Public Sub CreaCodigo(ByRef xCampo As String, ByRef xtabla As String, ByRef xRS As Object)
'xRS.Open "SELECT max(val('" & xCampo & "')) FROM '" & xtabla & "'", strcon, adopenstatic
'End Sub
