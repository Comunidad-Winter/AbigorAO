VERSION 5.00
Begin VB.Form Canjepirata 
   BorderStyle     =   0  'None
   Caption         =   "Objetos de Canje para Arqueros"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   Picture         =   "Canjepirata.frx":0000
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
      ItemData        =   "Canjepirata.frx":57F2
      Left            =   120
      List            =   "Canjepirata.frx":57F4
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
Attribute VB_Name = "Canjepirata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call SendData("JU")
List1.AddItem "Armadura de Pirata (Altos)"
List1.AddItem "Armadura de Pirata (Bajos)"
List1.AddItem "Casco Endemoniado"
List1.AddItem "Escudo Oscuro"
List1.AddItem "Hacha de Oro"
End Sub

Private Sub Label1_Click()
If List1.Text = "Armadura de Pirata (Altos)" Then Call SendData("/CANJEO T46")
If List1.Text = "Armadura de Pirata (Bajos)" Then Call SendData("/CANJEO T47")
If List1.Text = "Casco Endemoniado" Then Call SendData("/CANJEO T48")
If List1.Text = "Escudo Oscuro" Then Call SendData("/CANJEO T49")
If List1.Text = "Hacha de Oro" Then Call SendData("/CANJEO T50")
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub list1_Click()
 
If List1.Text = "Armadura de Pirata (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1586.bmp")
 
    lblPrecio.Caption = "75 Canjes"
    lblStat.Caption = "Min: 50 / Max: 50"
    lblPermisos.Caption = "Pirata"
End If
 
If List1.Text = "Armadura de Pirata (Bajos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1588.bmp")
 
    lblPrecio.Caption = "75 Canjes"
    lblStat.Caption = "Min: 50 / Max: 50"
    lblPermisos.Caption = "Pirata"
End If
 
If List1.Text = "Casco Endemoniado" Then
    Picture1.Picture = LoadPicture(DirGraficos & "2021.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 7 / Max: 14"
    lblPermisos.Caption = "Pirata"
End If

If List1.Text = "Escudo Oscuro" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1922.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 2 / Max: 2"
    lblPermisos.Caption = "Pirata"
End If

If List1.Text = "Hacha de Oro" Then
    Picture1.Picture = LoadPicture(DirGraficos & "974.bmp")
   
    lblPrecio.Caption = "50 Canjes"
    lblStat.Caption = "Min: 18 / Max: 21"
    lblPermisos.Caption = "Pirata"
End If

End Sub

