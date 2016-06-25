VERSION 5.00
Begin VB.Form frmNuevaCategoria 
   Caption         =   "Categoría"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtNomb 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ingrese el nombre de la nueva categoría"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmNuevaCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objCat As New Boss.Categoria


Private Sub cmdAceptar_Click()
objCat.GuardarNuevaCategoria txtNomb
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub
