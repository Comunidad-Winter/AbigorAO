VERSION 5.00
Begin VB.Form frmBaneos 
   BackColor       =   &H00000000&
   Caption         =   "Baneos By Petin."
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4650
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Banear PJ por IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   1935
   End
   Begin VB.OptionButton optED 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por tener equipo que no debe(Copas, etc)   ( IP )"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   3240
      Width           =   3735
   End
   Begin VB.OptionButton optEditGM 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por edit de un GM hacia el usuario. ( IP )"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   2640
      Width           =   3735
   End
   Begin VB.OptionButton optTS 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por querer tirar el Servidor. ( IP )"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox txtNickBaneador 
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Banear PJ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton optIGMG 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por insultos GRAVES a Game Master (IP)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   1440
      Width           =   3735
   End
   Begin VB.OptionButton optSPAM 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por Spam, dentro del servidor (10 Days)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   3615
   End
   Begin VB.OptionButton optSD 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por mal uso de Soportes/Denuncias (3 Days)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   4095
   End
   Begin VB.OptionButton optCH 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por uso de Programas Externos (30 Days)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   3855
   End
   Begin VB.OptionButton optIGM 
      BackColor       =   &H00000000&
      Caption         =   "Baneo por insultos a Game Master (10 Days)"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Baneos Por Dias."
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Baneos Por IP."
      ForeColor       =   &H000000FF&
      Height          =   2655
      Left            =   5160
      TabIndex        =   8
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Del Personaje:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Causas de Baneo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmBaneos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If optIGM.Enabled = True Then
Call SendData("/BAN Baneado por insulto a GM@" & txtNickBaneador.Text & "@10")
End If

If optCH.Enabled = True Then
Call SendData("/BAN Baneado por uso de Programa Externo@" & txtNickBaneador.Text & "@30")
End If

If optSD.Enabled = True Then
Call SendData("/BAN Baneado por mal uso de Soporte/Denuncias@" & txtNickBaneador.Text & "@3")
End If

If optSPAM.Enabled = True Then
Call SendData("/BAN Baneado por Spam en el servidor@" & txtNickBaneador.Text & "@10")
End If
End Sub

Private Sub Command2_Click()
If optIGMG.Enabled = True Then
Call SendData("/Banip " & txtNickBaneador.Text)
End If

If optTS.Enabled = True Then
Call SendData("/Banip " & txtNickBaneador.Text)
End If

If optEditGM.Enabled = True Then
Call SendData("/Banip " & txtNickBaneador.Text)
End If

If optED.Enabled = True Then
Call SendData("/Banip " & txtNickBaneador.Text)
End If
End Sub
