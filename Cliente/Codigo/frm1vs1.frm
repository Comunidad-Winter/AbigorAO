VERSION 5.00
Begin VB.Form Duelo_Torneo 
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command13 
      Caption         =   "2 - 1 A Favor"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   " 2 - 1 A Favor"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Duelean 2º Vez"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Text            =   "."
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1 A 1"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1 A 1"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Gana TORNEO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "2 - 0 A Favor"
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2 - 0 A Favor"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "1 - 0 A Favor"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "1 - 0 A Favor"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pierde"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pierde"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Duelean"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Duelo_Torneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/RMSG Torneo> Se enfrentan en esta nueva Ronda:" & " " & Text1.Text & " vs " & Text2.Text)
Call SendData("/RMSG Torneo> Suerte para Ambos..")
Call SendData("/RMSG Torneo> Esquinas y sale en..")
Call SendData("/CUENTA 5")
Call SendData("/TELEP" & " " & Text1.Text & " " & "60 88 85")
Call SendData("/TELEP" & " " & Text2.Text & " " & "60 65 72")
End Sub

Private Sub Command10_Click()
Call SendData("/RMSG Torneo> Y el Ganador del Torneo es:" & " " & Text3.Text)
Call SendData("/GANOTORNEO")
Call SendData("/TELEP" & " " & Text3.Text & " " & "1 50 69")
Call SendData("/RMSG Gracias por Participar..")
End Sub
Private Sub Command11_Click()
Call SendData("/RMSG Torneo> Se enfrentan nuevamente:" & " " & Text1.Text & " vs " & Text2.Text)
Call SendData("/RMSG Torneo> Suerte para Ambos..")
Call SendData("/RMSG Torneo> Esquinas y sale en..")
Call SendData("/TELEP" & " " & Text1.Text & " " & "60 88 85")
Call SendData("/TELEP" & " " & Text2.Text & " " & "60 65 72")
Call SendData("/CUENTA 5")
End Sub

Private Sub Command12_Click()
Call SendData("/RMSG Torneo> 2 A 1 A favor de" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command13_Click()
Call SendData("/RMSG Torneo> 2 A 1 A favor de" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

Private Sub Command2_Click()
Call SendData("/RMSG Torneo> Lo empata" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command3_Click()
Call SendData("/RMSG Torneo> Lo empata" & " " & Text2.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

Private Sub Command4_Click()
Call SendData("/RMSG Torneo> Pierde:" & " " & Text1.Text & " Y queda descalificado del Torneo.")
Call SendData("/TELEP" & " " & Text1.Text & " " & "1 50 69")
End Sub

Private Sub command5_Click()
Call SendData("/RMSG Torneo> Pierde:" & " " & Text2.Text & " Y queda descalificado del Torneo.")
Call SendData("/TELEP" & " " & Text2.Text & " " & "1 50 69")
End Sub

Private Sub Command6_Click()
Call SendData("/RMSG Torneo> 1 A 0 A favor de" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command7_Click()
Call SendData("/RMSG Torneo> 1 A 0 A favor de" & " " & Text2.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

Private Sub Command8_Click()
Call SendData("/RMSG Torneo> 2 A 0 A favor de" & " " & Text1.Text)
Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command9_Click()
Call SendData("/RMSG Torneo> 2 A 0 A favor de" & " " & Text2.Text)
Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

