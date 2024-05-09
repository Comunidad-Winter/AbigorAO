VERSION 5.00
Begin VB.Form cuenta 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   Picture         =   "cuenta.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image6 
      Height          =   615
      Left            =   840
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   1440
      Top             =   6120
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   840
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   840
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   840
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   840
      Top             =   4320
      Width           =   2175
   End
End
Attribute VB_Name = "cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
frmParty.ListaIntegrantes.Clear
LlegoParty = False
Call SendData("PARINF")
Do While Not LlegoParty
    DoEvents
Loop
frmParty.Visible = True
frmParty.SetFocus
LlegoParty = False
End Sub

Private Sub Image2_Click()
        If frmGuildLeader.Visible Then frmGuildLeader.Visible = False
        If frmGuildsNuevo.Visible Then frmGuildsNuevo.Visible = False
        If frmGuildAdm.Visible Then frmGuildAdm.Visible = False
        Call SendData("GLINFO")
End Sub

Private Sub Image3_Click()
frmOpciones.Show
End Sub

Private Sub Image4_Click()
frmEstadisticas.Show
LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
        SendData "ATRI"
        SendData "ESKI"
        SendData "FAMA"
        Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama Or Not LlegoMinist
            DoEvents
        Loop
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Show
        LlegaronAtrib = False
        LlegaronSkills = False
        LlegoFama = False
        LlegoMinist = False
End Sub

Private Sub Image5_Click()
Unload Me
End Sub

Private Sub Image6_Click()
Canjes.Show
Unload Me
End Sub
