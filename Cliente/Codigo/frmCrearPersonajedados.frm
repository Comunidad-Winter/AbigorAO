VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCrearPersonajedados.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   12075.47
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCorreo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   285
      Left            =   345
      TabIndex        =   30
      Top             =   2280
      Width           =   3840
   End
   Begin VB.TextBox txtPasswdCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   32
      Top             =   3720
      Width           =   3720
   End
   Begin VB.TextBox txtPasswd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   345
      PasswordChar    =   "*"
      TabIndex        =   31
      Top             =   3000
      Width           =   3720
   End
   Begin VB.TextBox txtCorreo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   330
      Left            =   360
      TabIndex        =   29
      Top             =   1440
      Width           =   3720
   End
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":611B4
      Left            =   4680
      List            =   "frmCrearPersonajedados.frx":611BE
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   7200
      Width           =   3480
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":611D1
      Left            =   4680
      List            =   "frmCrearPersonajedados.frx":611E4
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   6360
      Width           =   3480
   End
   Begin VB.ComboBox lstHogar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":61211
      Left            =   4680
      List            =   "frmCrearPersonajedados.frx":61221
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   8040
      Width           =   3480
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   300
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   9735
   End
   Begin VB.Label modCarisma 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   44
      Top             =   4680
      Width           =   330
   End
   Begin VB.Label modInteligencia 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   43
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label modConstitucion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   42
      Top             =   3360
      Width           =   330
   End
   Begin VB.Label modAgilidad 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   41
      Top             =   2640
      Width           =   330
   End
   Begin VB.Label modfuerza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   40
      Top             =   2040
      Width           =   330
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   360
      MouseIcon       =   "frmCrearPersonajedados.frx":6124A
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label lblPass2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4200
      TabIndex        =   39
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label lblMailOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4200
      TabIndex        =   37
      Top             =   1440
      Width           =   345
   End
   Begin VB.Label lblMail2OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   510
      Left            =   4200
      TabIndex        =   35
      Top             =   2160
      Width           =   345
   End
   Begin VB.Label lblPassOK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4200
      TabIndex        =   33
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   21
      Left            =   11090
      TabIndex        =   28
      Top             =   8160
      Width           =   405
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   42
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":61554
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   43
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":616A6
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   255
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   11160
      TabIndex        =   27
      Top             =   8640
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":617F8
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":6194A
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   7
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":61A9C
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   9
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":61BEE
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":61D40
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":61E92
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":61FE4
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":62136
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   19
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":62288
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   21
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":623DA
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":6252C
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":6267E
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   285
      Index           =   27
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":627D0
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":62922
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   0
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":62A74
      MousePointer    =   99  'Custom
      Top             =   1080
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":62BC6
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":62D18
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   6
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":62E6A
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   8
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":62FBC
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   10
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":6310E
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":63260
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   14
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":633B2
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   16
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":63504
      MousePointer    =   99  'Custom
      Top             =   3840
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   18
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":63656
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   20
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":637A8
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":638FA
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   24
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":63A4C
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   26
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":63B9E
      MousePointer    =   99  'Custom
      Top             =   5520
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   28
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":63CF0
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   29
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":63E42
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   30
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":63F94
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   31
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":640E6
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   32
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":64238
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   33
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":6438A
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   34
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":644DC
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   35
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":6462E
      MousePointer    =   99  'Custom
      Top             =   6840
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   36
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":64780
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":648D2
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   38
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":64A24
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":64B76
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   285
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   40
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":64CC8
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   41
      Left            =   10920
      MouseIcon       =   "frmCrearPersonajedados.frx":64E1A
      MousePointer    =   99  'Custom
      Top             =   7800
      Width           =   255
   End
   Begin VB.Image boton 
      Height          =   495
      Index           =   1
      Left            =   4440
      MouseIcon       =   "frmCrearPersonajedados.frx":64F6C
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   4005
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   450
      Index           =   0
      Left            =   240
      MouseIcon       =   "frmCrearPersonajedados.frx":650BE
      MousePointer    =   99  'Custom
      Top             =   8040
      Width           =   3960
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   20
      Left            =   11090
      TabIndex        =   26
      Top             =   7800
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   19
      Left            =   11085
      TabIndex        =   25
      Top             =   7485
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   18
      Left            =   11085
      TabIndex        =   24
      Top             =   7125
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   17
      Left            =   11090
      TabIndex        =   23
      Top             =   6795
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   16
      Left            =   11090
      TabIndex        =   22
      Top             =   6450
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   15
      Left            =   11090
      TabIndex        =   21
      Top             =   6120
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   14
      Left            =   11090
      TabIndex        =   20
      Top             =   5760
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   13
      Left            =   11090
      TabIndex        =   19
      Top             =   5400
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   12
      Left            =   11090
      TabIndex        =   18
      Top             =   5040
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   11
      Left            =   11090
      TabIndex        =   17
      Top             =   4725
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   10
      Left            =   11090
      TabIndex        =   16
      Top             =   4395
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   9
      Left            =   11090
      TabIndex        =   15
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   8
      Left            =   11090
      TabIndex        =   14
      Top             =   3750
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   7
      Left            =   11090
      TabIndex        =   13
      Top             =   3420
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   6
      Left            =   11090
      TabIndex        =   12
      Top             =   3075
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   5
      Left            =   11090
      TabIndex        =   11
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   4
      Left            =   11090
      TabIndex        =   10
      Top             =   2370
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   3
      Left            =   11090
      TabIndex        =   9
      Top             =   2010
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   2
      Left            =   11090
      TabIndex        =   8
      Top             =   1725
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   0
      Left            =   11090
      TabIndex        =   7
      Top             =   1050
      Width           =   405
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   1
      Left            =   11090
      TabIndex        =   6
      Top             =   1395
      Width           =   405
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   5
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   4
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   3
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   2
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   450
      Left            =   6120
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Public SkillPoints As Byte
Function CheckData() As Boolean

If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserHogar = 0 Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If UserSexo = -1 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True

End Function
Private Sub boton_Click(Index As Integer)
Dim i As Integer
Dim k As Object
        
Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0
        LlegoConfirmacion = False
        Confirmacion = 0

        i = 1
        
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
            UserName = Trim(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.ListIndex + 1
        UserSexo = lstGenero.ListIndex
        UserHogar = lstHogar.ListIndex + 1
        
        UserAtributos(1) = 1
        UserAtributos(2) = 1
        UserAtributos(3) = 1
        UserAtributos(4) = 1
        UserAtributos(5) = 1
        
        If CheckData() Then
            UserPassword = txtPasswd.Text
            UserEmail = txtCorreo.Text
            
            If Not CheckMailString(UserEmail) Then
                MsgBox "Direccion de mail inv�lida.", vbExclamation, "Abigor AO"
                txtCorreo.SetFocus
                Exit Sub
            End If
    
            If UserEmail <> txtCorreo2.Text Then
                MsgBox "Las direcciones de mail no coinciden.", vbExclamation, "Abigor AO"
                txtCorreo2.Text = ""
                txtCorreo2.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtPasswd)) = 0 Then
                MsgBox "Ten�s que ingresar una contrase�a.", vbExclamation, "Abigor AO"
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            If Len(Trim(txtPasswd)) < 6 Then
                MsgBox "El password debe tener al menos 6 caracteres.", vbExclamation, "Abigor AO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
            
            If Trim(txtPasswd) <> Trim(txtPasswdCheck) Then
                MsgBox "Las contrase�as no coinciden.", vbInformation, "Abigor AO"
                txtPasswd = ""
                txtPasswdCheck = ""
                txtPasswd.SetFocus
                Exit Sub
            End If
    
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
    
            Me.MousePointer = 11
            EstadoLogin = CrearNuevoPj
    
            If Not frmMain.Socket1.Connected Then
                Call MsgBox("Error: Se ha perdido la conexion con el server.")
                Unload Me
            Else
                Call Login(ValidarLoginMSG(CInt(bRK)))
            End If
            
            If Musica = 0 Then
                CurMidi = DirMidi & "2.mid"
                LoopMidi = 1
                Call CargarMIDI(CurMidi)
                Call Play_Midi
            End If
        
            frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        End If

    Case 1
        If Musica = 0 Then
            CurMidi = DirMidi & "2.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If
        
        frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        
        frmMain.Socket1.Disconnect
        frmConnect.MousePointer = 1
        Unload Me
End Select

End Sub
Private Sub Command1_Click(Index As Integer)
Call PlayWaveDS(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub
Private Sub Form_Load()

SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.Path & "\graficos\CrearPersonajeConDados.jpg")
Me.MousePointer = vbDefault

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
End Select

End Sub

Private Sub P�cture4_Click()

End Sub

Private Sub Image1_Click()
PlayWaveDS (SND_CLICK)
Call SendData("TIRDAD")
End Sub

Private Sub lstRaza_click()

Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = "+ 2"
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = ""
        modCarisma.Caption = ""
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 2"
    Case Is = "Elfo Oscuro"
        modfuerza.Caption = "+ 1"
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 1"
        modInteligencia.Caption = "+ 2"
        modCarisma.Caption = "- 3"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 3"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modCarisma.Caption = "- 3"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 4"
        modInteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 1"
End Select

End Sub
Private Sub txtCorreo_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
    Exit Sub
End If

lblMailOK = "P"
lblMailOK.ForeColor = &H80FF&

If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
Else
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
End If

End Sub
Private Sub txtCorreo_GotFocus()

MsgBox "La direcci�n de correo electr�nico DEBE SER real."

End Sub
Private Sub txtCorreo2_Change()

If Not CheckMailString(txtCorreo) Then
    lblMailOK = "O"
    lblMailOK.ForeColor = &HC0&
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
    Exit Sub
End If

lblMailOK = "P"
lblMailOK.ForeColor = &H80FF&

If (UCase$(txtCorreo.Text) = UCase$(txtCorreo2.Text)) Then
    lblMail2OK = "P"
    lblMail2OK.ForeColor = &H80FF&
Else
    lblMail2OK = "O"
    lblMail2OK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswd_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "O"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "P"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "P"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtPasswdCheck_Change()

If Len(Trim(txtPasswd)) < 6 Then
    lblPass2OK = "O"
    lblPass2OK.ForeColor = &HC0&
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
    Exit Sub
End If

lblPass2OK = "P"
lblPass2OK.ForeColor = &H80FF&

If (txtPasswdCheck = txtPasswd) Then
    lblPassOK = "P"
    lblPassOK.ForeColor = &H80FF&
Else
    lblPassOK = "O"
    lblPassOK.ForeColor = &HC0&
End If

End Sub
Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotaci�n politica los administradores borrar�n su personaje y no habr� ninguna posibilidad de recuperarlo."
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub
