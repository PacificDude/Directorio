VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmNuevoContacto 
   Caption         =   "Contactos"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   1920
      TabIndex        =   21
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   495
      Left            =   1080
      TabIndex        =   20
      Top             =   4320
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DtcCat 
      Height          =   315
      Left            =   1440
      TabIndex        =   19
      Top             =   3600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPFec 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19726337
      CurrentDate     =   38348
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox txtCelu 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtTelf 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DtcDist 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtDirec 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtApel 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtNomb 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Categoría:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Fec. Nac."
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "e-mail:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Celular:"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Teléfono:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Distrito:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Apellidos:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Alias:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmNuevoContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objCon As New Boss.Contacto
Dim objDis As New Boss.Distrito
Dim objCat As New Boss.Categoria

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdGuardar_Click()
If Trim(txtNomb) = "" Then MsgBox "Ingrese al menos el Nombre del contacto", vbExclamation, "Contactos": Exit Sub
If DtcCat.Text = "" Then MsgBox "Elija la categoría a la q pertenecerá el Contacto", vbExclamation, "Contactos": Exit Sub
objCon.GuardarContacto txtCod, txtNomb, txtApel, txtDirec, Val(DtcDist.BoundText), txtTelf, txtCelu, txtEmail, DTPFec.Value, Val(DtcCat.BoundText)
MsgBox "Registro guardado", vbInformation, "Contactos"
End Sub

Private Sub cmdLimpiar_Click()
Call Limpiar(Me)
End Sub

Private Sub Form_Load()
Set DtcCat.RowSource = objCat.LeerCategorias
DtcCat.BoundColumn = "IdCat"
DtcCat.ListField = "Nomb"

Set DtcDist.RowSource = objDis.LeerDistritos
DtcDist.BoundColumn = "IdDist"
DtcDist.ListField = "Nomb"

End Sub
