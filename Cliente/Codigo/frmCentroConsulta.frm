VERSION 5.00
Begin VB.Form frmCentroConsulta 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRespuesta 
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   3735
   End
   Begin VB.TextBox txtConsulta 
      Height          =   1575
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   3735
   End
   Begin VB.ListBox lstUsuarios 
      Height          =   1620
      ItemData        =   "frmCentroConsulta.frx":0000
      Left            =   240
      List            =   "frmCentroConsulta.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ir al Panel de GM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Respuesta del GM a la consulta:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulta hecha por el usuario."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label sendResp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENVIAR RESPUESTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REINICIAR CONSULTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   7
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BORRAR CONSULTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TRAER USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IR AL USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios en espera a ser atendidos:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmCentroConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click(Index As Integer)
Select Case Index
    Case 0
        SendData ("/IRA " & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
    Case 1
        SendData ("/SUM " & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
    Case 2
        If lstUsuarios.ListIndex < 0 Then Exit Sub
        SendData ("BORRACONSULTA" & lstUsuarios.List(lstUsuarios.ListIndex))
        lstUsuarios.RemoveItem lstUsuarios.ListIndex
    Case 3
        Call SendData("/BORRAR SOS")
        lstUsuarios.Clear
End Select
 
End Sub
 
Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label6_Click()
frmPanelGm.Show
Unload Me
End Sub

Private Sub lstUsuarios_Click()
Dim ind As Integer
ind = Val(ReadField(2, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
SendData ("/VERCONSULTA " & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
End Sub
Private Sub sendResp_Click()
SendData ("/RESPONDER " & txtRespuesta.Text & "@" & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
End Sub

