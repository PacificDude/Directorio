VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOrdenarContacto 
   Caption         =   "Ordenar Contacto"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIzq 
      Caption         =   "<"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox cmbCat 
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo DtcGrupo 
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton cmdDer 
      Caption         =   ">"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   375
      Left            =   120
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Categoría:"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmOrdenarContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objCat As New Boss.Categoria
Dim objCon As New Boss.Contacto
Dim camCon As Object
Dim camCat As Object
Dim vCon(9999) As String
Dim vCat(255) As Byte
Dim i As Integer

Private Sub cmdAceptar_Click()
For x = 0 To i
    objCon.ModificarCateg vCon(x), vCat(x)
Next
Form_Load
End Sub

Private Sub cmdCancelar_Click()
If cmdCancelar.Caption = "Cancelar" Then
    DtcGrupo.Enabled = True
    cmdCancelar.Caption = "Cerrar"
Else
    Unload Me
End If
End Sub

Private Sub cmdDer_Click()
If DtcGrupo.Text = "" Then MsgBox "Elija una Categoría", vbExclamation, "Administrar": Exit Sub
If List1.Text = "" Then MsgBox "Seleccione un Contacto", vbExclamation, "Administrar": Exit Sub

For x = 0 To i
    If vCon(x) = List1.Text Then
        vCon(x) = List1.Text
        vCat(x) = DtcGrupo.BoundText
    End If
Next

vCon(i) = List1.Text
vCat(i) = DtcGrupo.BoundText
i = i + 1
List2.AddItem List1.Text

cmdAceptar.Enabled = True
cmdCancelar.Caption = "Cancelar"
End Sub

Private Sub DtcGrupo_Change()
Dim camCon2 As Object

Set camCon2 = objCon.BuscarContactosxCateg(DtcGrupo.BoundText)
List2.Clear
Do While Not camCon2.EOF
    List2.AddItem camCon2(0)
    camCon2.MoveNext
Loop
DtcGrupo.Enabled = False
cmdCancelar.Caption = "Cancelar"
End Sub

Private Sub Form_Load()
Set camCon = objCon.LeerContactos
List1.Clear
Do While Not camCon.EOF
    List1.AddItem camCon(0)
    camCon.MoveNext
Loop

Set DtcGrupo.RowSource = objCat.LeerCategorias
DtcGrupo.BoundColumn = "IdCat"
DtcGrupo.ListField = "Nomb"

Set camCat = objCat.LeerCategorias
cmbCat.Clear
Do While Not camCat.EOF
    cmbCat.AddItem camCat(1)
    camCat.MoveNext
Loop
DtcGrupo.Enabled = True
cmdAceptar.Enabled = False
cmdCancelar.Caption = "Cerrar"
End Sub
