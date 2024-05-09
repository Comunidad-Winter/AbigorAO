VERSION 5.00
Begin VB.Form Canjecazador 
   BorderStyle     =   0  'None
   Caption         =   "Objetos de Canje para Arqueros"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   Picture         =   "Canjecazador.frx":0000
   ScaleHeight     =   5070
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   3960
      ItemData        =   "Canjecazador.frx":57F2
      Left            =   120
      List            =   "Canjecazador.frx":57F4
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblPermisos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label lblPrecio 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Canjear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
End
Attribute VB_Name = "Canjecazador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call SendData("JU")
List1.AddItem "Armadura de cazador (altos)"
List1.AddItem "Armadura de cazador (Bajos)"
List1.AddItem "Casco Nordico"
List1.AddItem "Escudo de Roble"
List1.AddItem "Arco Celestial"
End Sub

Private Sub Label1_Click()
If List1.Text = "Armadura de cazador (altos)" Then Call SendData("/CANJEO T15")
If List1.Text = "Armadura de cazador (Bajos)" Then Call SendData("/CANJEO T16")
If List1.Text = "Casco Nordico" Then Call SendData("/CANJEO T17")
If List1.Text = "Escudo de Roble" Then Call SendData("/CANJEO T18")
If List1.Text = "Arco Celestial" Then Call SendData("/CANJEO T19")
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub list1_Click()
 
If List1.Text = "Armadura de cazador (altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1554.bmp")
 
    lblPrecio.Caption = "75 Canjes"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Cazador"
End If
 
If List1.Text = "Armadura de cazador (Bajos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1556.bmp")
 
    lblPrecio.Caption = "75 Canjes"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Cazador"
End If
 
If List1.Text = "Casco Nordico" Then
    Picture1.Picture = LoadPicture(DirGraficos & "2029.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Cazador"
End If

If List1.Text = "Escudo de Roble" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1920.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 2 / Max: 2"
    lblPermisos.Caption = "Cazador"
End If

If List1.Text = "Arco Celestial" Then
    Picture1.Picture = LoadPicture(DirGraficos & "9626.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 9 / Max: 15"
    lblPermisos.Caption = "Cazador"
End If

End Sub

