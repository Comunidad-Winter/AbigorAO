VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPass 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3120
      Width           =   2445
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4680
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2040
      Width           =   2445
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   10680
      MouseIcon       =   "frmConnect.frx":279BD
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Image imgWeb 
      Height          =   615
      Left            =   480
      MouseIcon       =   "frmConnect.frx":27CC7
      MousePointer    =   99  'Custom
      Top             =   720
      Width           =   3375
   End
   Begin VB.Image imgGetPass 
      Height          =   435
      Left            =   7560
      MouseIcon       =   "frmConnect.frx":27FD1
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   0
      Left            =   1560
      MouseIcon       =   "frmConnect.frx":282DB
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   2
      Left            =   4560
      MouseIcon       =   "frmConnect.frx":285E5
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   2610
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Option Explicit

Private Sub Command1_Click()
Password.Left = RandomNumber(1, 9150)
Password.Top = RandomNumber(1, 7500)
Password.Show
Password.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    Call PlayWaveDS(SND_CLICK)
            
    If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
    
    If frmConnect.MousePointer = 11 Then
    frmConnect.MousePointer = 1
        Exit Sub
    End If
    
    
    UserName = txtUser.Text
    Dim aux As String
    aux = txtPass.Text
    UserPassword = aux
    If CheckUserData(False) = True Then
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        EstadoLogin = Normal
        Me.MousePointer = 11
        frmMain.Socket1.Connect
    End If
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    frmCargando.Show
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Cerrando AbigorAO.", 255, 150, 50, 1, 0, 1
    
    Call SaveGameini
    frmConnect.MousePointer = 1
    frmMain.MousePointer = 1
    prgRun = False
    
    AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
    frmCargando.Refresh
    LiberarObjetosDX
    AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, 0, 1
    AddtoRichTextBox frmCargando.Status, "��Gracias por jugar AbigorAO!!", 255, 150, 50, 1, 0, 1
    frmCargando.Refresh
    Call UnloadAllForms
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    

    
    
    


    
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    
    EngineRun = False
    
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next

 IntervaloPaso = 0.19
 IntervaloUsar = 0.14
 Picture = LoadPicture(DirGraficos & "conectar.jpg")


 
 
 
 
 
 

End Sub

Private Sub Image1_Click(Index As Integer)

CurServer = 0
Unload Password

Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0

        If Musica = 0 Then
            CurMidi = DirMidi & "7.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If

       
        EstadoLogin = dados
        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        Me.MousePointer = 11
        frmMain.Socket1.Connect
        
    Case 1
        
        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
        
        If frmConnect.MousePointer = 11 Then
        frmConnect.MousePointer = 1
            Exit Sub
        End If
        
        
        
        UserName = txtUser.Text
        Dim aux As String
        aux = txtPass.Text
        If CheckUserData(False) = True Then
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
            
            EstadoLogin = Normal
            Me.MousePointer = 11
            frmMain.Socket1.Connect
        End If
        
    Case 2
If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
 
If frmConnect.MousePointer = 11 Then
frmConnect.MousePointer = 1
Exit Sub
End If
 
frmMain.Socket1.HostName = IPdelServidor
frmMain.Socket1.RemotePort = PuertoDelServidor
EstadoLogin = BorrarPJ
Me.MousePointer = 11
frmMain.Socket1.Connect
End Select

End Sub
Private Sub Image2_Click()

MsgBox "Created By Abigor AO Staff." & vbCrLf & "Copyright � 2012. Todos los derechos reservados." & vbCrLf & vbCrLf & "Web: http://www.abigoronline.com.ar" & vbCrLf & vbCrLf & "�Gracias por Jugar nuestro Argentum Online!" & vbCrLf & "Staff Abigor AO.", vbInformation, "Versi�n 1.0"

End Sub

Private Sub imgGetPass_Click()
 
If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect
 
If frmConnect.MousePointer = 11 Then
    frmConnect.MousePointer = 1
    Exit Sub
End If

frmMain.Socket1.HostName = IPdelServidor
frmMain.Socket1.RemotePort = PuertoDelServidor
EstadoLogin = RecuperarPass
Me.MousePointer = 11
frmMain.Socket1.Connect
 
End Sub

Private Sub imgWeb_Click()

Call ShellExecute(Me.Hwnd, "open", "http://www.abigoronline.com.ar", "", "", 1)

End Sub
