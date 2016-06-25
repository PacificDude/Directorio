VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmContactos 
   Caption         =   "Contactos"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "N"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "M"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "C"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "B"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "l<"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "<"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   ">"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   ">l"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTabAlfa 
      Height          =   4770
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   8414
      _Version        =   393216
      Tabs            =   27
      TabsPerRow      =   14
      TabHeight       =   520
      TabCaption(0)   =   "A"
      TabPicture(0)   =   "frmContactos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSLista(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "B"
      TabPicture(1)   =   "frmContactos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSLista(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "C"
      TabPicture(2)   =   "frmContactos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSLista(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "D"
      TabPicture(3)   =   "frmContactos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSLista(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "E"
      TabPicture(4)   =   "frmContactos.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "F"
      TabPicture(5)   =   "frmContactos.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "G"
      TabPicture(6)   =   "frmContactos.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "H"
      TabPicture(7)   =   "frmContactos.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "I"
      TabPicture(8)   =   "frmContactos.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "J"
      TabPicture(9)   =   "frmContactos.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "K"
      TabPicture(10)  =   "frmContactos.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "L"
      TabPicture(11)  =   "frmContactos.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
      TabCaption(12)  =   "M"
      TabPicture(12)  =   "frmContactos.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      TabCaption(13)  =   "N"
      TabPicture(13)  =   "frmContactos.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).ControlCount=   0
      TabCaption(14)  =   "Ñ"
      TabPicture(14)  =   "frmContactos.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).ControlCount=   0
      TabCaption(15)  =   "O"
      TabPicture(15)  =   "frmContactos.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).ControlCount=   0
      TabCaption(16)  =   "P"
      TabPicture(16)  =   "frmContactos.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).ControlCount=   0
      TabCaption(17)  =   "Q"
      TabPicture(17)  =   "frmContactos.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).ControlCount=   0
      TabCaption(18)  =   "R"
      TabPicture(18)  =   "frmContactos.frx":01F8
      Tab(18).ControlEnabled=   0   'False
      Tab(18).ControlCount=   0
      TabCaption(19)  =   "S"
      TabPicture(19)  =   "frmContactos.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).ControlCount=   0
      TabCaption(20)  =   "T"
      TabPicture(20)  =   "frmContactos.frx":0230
      Tab(20).ControlEnabled=   0   'False
      Tab(20).ControlCount=   0
      TabCaption(21)  =   "U"
      TabPicture(21)  =   "frmContactos.frx":024C
      Tab(21).ControlEnabled=   0   'False
      Tab(21).ControlCount=   0
      TabCaption(22)  =   "V"
      TabPicture(22)  =   "frmContactos.frx":0268
      Tab(22).ControlEnabled=   0   'False
      Tab(22).ControlCount=   0
      TabCaption(23)  =   "W"
      TabPicture(23)  =   "frmContactos.frx":0284
      Tab(23).ControlEnabled=   0   'False
      Tab(23).ControlCount=   0
      TabCaption(24)  =   "X"
      TabPicture(24)  =   "frmContactos.frx":02A0
      Tab(24).ControlEnabled=   0   'False
      Tab(24).ControlCount=   0
      TabCaption(25)  =   "Y"
      TabPicture(25)  =   "frmContactos.frx":02BC
      Tab(25).ControlEnabled=   0   'False
      Tab(25).ControlCount=   0
      TabCaption(26)  =   "Z"
      TabPicture(26)  =   "frmContactos.frx":02D8
      Tab(26).ControlEnabled=   0   'False
      Tab(26).ControlCount=   0
      Begin MSFlexGridLib.MSFlexGrid MSLista 
         Height          =   3135
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   1200
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
      End
      Begin MSFlexGridLib.MSFlexGrid MSLista 
         Height          =   3135
         Index           =   1
         Left            =   -74400
         TabIndex        =   3
         Top             =   1200
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
      End
      Begin MSFlexGridLib.MSFlexGrid MSLista 
         Height          =   3135
         Index           =   2
         Left            =   -74400
         TabIndex        =   4
         Top             =   1200
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
      End
      Begin MSFlexGridLib.MSFlexGrid MSLista 
         Height          =   3135
         Index           =   3
         Left            =   -74400
         TabIndex        =   5
         Top             =   1200
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
      End
   End
   Begin VB.Menu MNUNew 
      Caption         =   "Nuevo"
      Begin VB.Menu MNUNewContact 
         Caption         =   "Contacto"
      End
      Begin VB.Menu MNUNewCategory 
         Caption         =   "Categoría"
      End
      Begin VB.Menu SA1 
         Caption         =   "-"
      End
      Begin VB.Menu MNUNewCity 
         Caption         =   "Distrito"
      End
   End
   Begin VB.Menu MNUAdministrar 
      Caption         =   "Administrar"
      Begin VB.Menu MNUOrdenarContactos 
         Caption         =   "Ordenar Contactos"
      End
   End
End
Attribute VB_Name = "frmContactos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objCon As New Boss.Contacto
Dim camCon As Object

Private Sub Form_Load()

Set camCon = objCon.LeerContactosxLetra(SSTabAlfa.TabCaption(SSTabAlfa.Tab))
For a = 1 To camCon.recordcount
    MSLista(SSTabAlfa.Tab).Rows = a + 1
    MSLista(SSTabAlfa.Tab).TextMatrix(a, 1) = camCon(1)
    MSLista(SSTabAlfa.Tab).TextMatrix(a, 2) = camCon(2)

    camCon.movenext
    
Next
End Sub


Private Sub MNUNewCategory_Click()
frmNuevaCategoria.Show 1
End Sub

Private Sub MNUNewContact_Click()
frmNuevoContacto.Show
End Sub

Private Sub MNUOrdenarContactos_Click()
frmOrdenarContacto.Show
End Sub

Private Sub SSTabAlfa_Click(PreviousTab As Integer)
Set camCon = objCon.LeerContactosxLetra(SSTabAlfa.TabCaption(SSTabAlfa.Tab))
For a = 1 To camCon.recordcount
    MSLista(SSTabAlfa.Tab).Rows = a + 1
    MSLista(SSTabAlfa.Tab).TextMatrix(a, 1) = camCon(1)
    MSLista(SSTabAlfa.Tab).TextMatrix(a, 2) = camCon(2)
    camCon.movenext
Next
End Sub

