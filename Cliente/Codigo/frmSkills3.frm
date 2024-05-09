VERSION 5.00
Begin VB.Form frmSkills3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSkills3.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1920
      MouseIcon       =   "frmSkills3.frx":FD82
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   855
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   43
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":1008C
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   42
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":10396
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   22
      Left            =   3720
      TabIndex        =   22
      Top             =   8280
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   21
      Top             =   240
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   20
      Top             =   600
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   19
      Top             =   960
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   18
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   17
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   16
      Top             =   2160
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   15
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   14
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   13
      Top             =   3240
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   12
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   3720
      TabIndex        =   11
      Top             =   4155
      Width           =   315
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   10
      Top             =   4470
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   0
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":106A0
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   2
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":109AA
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   3
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":10CB4
      MousePointer    =   99  'Custom
      Top             =   600
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   4
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":10FBE
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   5
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":112C8
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   6
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":115D2
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   7
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":118DC
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   8
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":11BE6
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   9
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":11EF0
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   10
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":121FA
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   11
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":12504
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   12
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":1280E
      Top             =   2520
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   13
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":12B18
      MousePointer    =   99  'Custom
      Top             =   2520
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   14
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":12E22
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   15
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":1312C
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   16
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":13436
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   17
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":13740
      MousePointer    =   99  'Custom
      Top             =   3240
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   18
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":13A4A
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   19
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":13D54
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   20
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":1405E
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   21
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":14368
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   22
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":14672
      MousePointer    =   99  'Custom
      Top             =   4440
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   23
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":1497C
      MousePointer    =   99  'Custom
      Top             =   4455
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   24
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":14C86
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   25
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":14F90
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   3720
      TabIndex        =   9
      Top             =   4920
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   26
      Left            =   4080
      MouseIcon       =   "frmSkills3.frx":1529A
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   27
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":155A4
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   3720
      TabIndex        =   8
      Top             =   5280
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   28
      Left            =   4080
      MouseIcon       =   "frmSkills3.frx":158AE
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   29
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":15BB8
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   3720
      TabIndex        =   7
      Top             =   5640
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   30
      Left            =   4080
      MouseIcon       =   "frmSkills3.frx":15EC2
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   31
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":161CC
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   16
      Left            =   3720
      TabIndex        =   6
      Top             =   6000
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   32
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":164D6
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   33
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":167E0
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   17
      Left            =   3720
      TabIndex        =   5
      Top             =   6360
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   34
      Left            =   4080
      MouseIcon       =   "frmSkills3.frx":16AEA
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   35
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":16DF4
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   18
      Left            =   3720
      TabIndex        =   4
      Top             =   6720
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   1
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":170FE
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   19
      Left            =   3720
      TabIndex        =   3
      Top             =   7200
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   36
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":17408
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   37
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":17712
      MousePointer    =   99  'Custom
      Top             =   7200
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   20
      Left            =   3720
      TabIndex        =   2
      Top             =   7560
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   38
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":17A1C
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   39
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":17D26
      MousePointer    =   99  'Custom
      Top             =   7560
      Width           =   300
   End
   Begin VB.Label text1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   21
      Left            =   3720
      TabIndex        =   1
      Top             =   7920
      Width           =   315
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   40
      Left            =   3960
      MouseIcon       =   "frmSkills3.frx":18030
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   300
      Index           =   41
      Left            =   3480
      MouseIcon       =   "frmSkills3.frx":1833A
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   300
   End
   Begin VB.Label puntos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Left            =   3750
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "frmSkills3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Option Explicit

Private Sub Command1_Click(Index As Integer)

Call PlayWaveDS(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If Alocados > 0 Then
        indice = Index \ 2 + 1
        If indice > NUMSKILLS Then indice = NUMSKILLS
        If UserSkills(indice) < MAXSKILLPOINTS And Val(Text1(indice).Caption) < 100 Then
            Text1(indice).Caption = Val(Text1(indice).Caption) + 1
            FLAGS(indice) = FLAGS(indice) + 1
            Alocados = Alocados - 1
        End If
            
    End If
Else
    If Alocados < SkillPoints Then
        
        indice = Index \ 2 + 1
        If Val(Text1(indice).Caption) > 0 And FLAGS(indice) > 0 Then
            Text1(indice).Caption = Val(Text1(indice).Caption) - 1
            FLAGS(indice) = FLAGS(indice) - 1
            Alocados = Alocados + 1
        End If
    End If
End If

Puntos.Caption = Alocados
End Sub

Private Sub Form_Deactivate()

Me.Visible = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    DX = X
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If bmoving And ((X <> DX) Or (Y <> dy)) Then Move Left + (X - DX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub

Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\AgregarPuntosSkills.jpg")




Dim i As Integer

ReDim FLAGS(1 To NUMSKILLS)






End Sub

Private Sub Image1_Click()

Dim i As Integer
Dim cad As String
For i = 1 To NUMSKILLS
    cad = cad & FLAGS(i) & ","
Next
SendData "SKSE" & cad
If Alocados = 0 Then frmMain.Label1.Visible = False
SkillPoints = Alocados
Unload Me
End Sub
