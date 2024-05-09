VERSION 5.00
Begin VB.Form Canjepaladin 
   BorderStyle     =   0  'None
   Caption         =   "Objetos de Canje para Arqueros"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   Picture         =   "Canjepaladin.frx":0000
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
      ItemData        =   "Canjepaladin.frx":57F2
      Left            =   120
      List            =   "Canjepaladin.frx":57F4
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
Attribute VB_Name = "Canjepaladin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call SendData("JU")
List1.AddItem "Armadura de Paladin (Altos)"
List1.AddItem "Armadura de Paladin (Bajos)"
List1.AddItem "Casco de Plumas Arcanas"
List1.AddItem "Escudo de Reflexión"
List1.AddItem "Espada del Destierro"
End Sub

Private Sub Label1_Click()
If List1.Text = "Armadura de Paladin (Altos)" Then Call SendData("/CANJEO T41")
If List1.Text = "Armadura de Paladin (Bajos)" Then Call SendData("/CANJEO T42")
If List1.Text = "Casco de Plumas Arcanas" Then Call SendData("/CANJEO T43")
If List1.Text = "Escudo de Reflexión" Then Call SendData("/CANJEO T44")
If List1.Text = "Espada del Destierro" Then Call SendData("/CANJEO T45")
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub list1_Click()
 
If List1.Text = "Armadura de Paladin (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1582.bmp")
 
    lblPrecio.Caption = "75 Canjes"
    lblStat.Caption = "Min: 50 / Max: 50"
    lblPermisos.Caption = "Paladin"
End If
 
If List1.Text = "Armadura de Paladin (Bajos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1584.bmp")
 
    lblPrecio.Caption = "75 Canjes"
    lblStat.Caption = "Min: 50 / Max: 50"
    lblPermisos.Caption = "Paladin"
End If
 
If List1.Text = "Casco de Plumas Arcanas" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1017.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 18 / Max: 21"
    lblPermisos.Caption = "Paladin"
End If

If List1.Text = "Escudo de Reflexión" Then
    Picture1.Picture = LoadPicture(DirGraficos & "3511.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 7 / Max: 9"
    lblPermisos.Caption = "Paladin"
End If

If List1.Text = "Espada del Destierro" Then
    Picture1.Picture = LoadPicture(DirGraficos & "972.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 19 / Max: 20"
    lblPermisos.Caption = "Paladin"
End If

End Sub

