VERSION 5.00
Begin VB.Form Items 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5700
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5610
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Items.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3015
      Left            =   100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Items.frx":13705
      Top             =   2580
      Width           =   5415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      IntegralHeight  =   0   'False
      Left            =   1410
      TabIndex        =   0
      Top             =   940
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "Items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
With List1
.AddItem "Asesino"
.AddItem "Bardo"
.AddItem "Clérigo"
.AddItem "Cazador"
.AddItem "Druida"
.AddItem "Guerrero"
.AddItem "Mago"
.AddItem "Paladin"
End With
If List1.ListIndex < 0 Then
Text1.Text = "Seleccione una clase."
End If
Text1.MaxLength = 0
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub List1_Click()

With List1
 Select Case .List(.ListIndex)
 
  Case "Asesino"
    Text1.Text = "Herreria " & vbNewLine & " Armadura Clipteum (Altos)" & vbNewLine & " Armadura Homicida(Altos / Bajos) " & vbNewLine & " Escudo Mazzone " & vbNewLine & " Escudo Bactrocera " & vbNewLine & " Casco Asessino " & vbNewLine & " Sombrero de mago " & vbNewLine & " Casco Asessino Ircano " & vbNewLine & " Daga Estilete " & vbNewLine & " Empuñadura de zamak " & vbNewLine & " Anillo Obsistens + 2 " & vbNewLine & vbNewLine & _
    " Torneo: " & vbNewLine & " Manto Alado(Altos / Bajos) " & vbNewLine & " Tunica Apocalíptica(Altos / Bajos) " & vbNewLine & " Armadura de placas doradas (Altos) " & vbNewLine & " Armadura Parietis(Bajos) " & vbNewLine & " Escudo Apicibus " & vbNewLine & " Daga Ínferos " & vbNewLine & " Anillo Regonem " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Crural " & vbNewLine & " Anillo Bulgari " & vbNewLine & " Anillo Mobius " & vbNewLine & " Casco de Bronce  " & vbNewLine & " Casco Frodo " & vbNewLine & " Casco Nocturno " & vbNewLine & " Espada de Hielo " & vbNewLine & " Daga Legendaria " & vbNewLine & " Daga Serpens " & vbNewLine & " Escudo de la Realeza " & vbNewLine & " Escudo de Hielo " & vbNewLine & " Escudo de Hierro " & vbNewLine & " Armadura Templario " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Armadura Ceremonial "
    
  Case "Bardo"
    Text1.Text = "Herreria: " & vbNewLine & " Todas las túnicas (Altos/Bajos) " & vbNewLine & " Escudo Tortuges " & vbNewLine & " Sombrero de mago " & vbNewLine & " Yelmo completo " & vbNewLine & " Daga Pink " & vbNewLine & " Daga Quodar " & vbNewLine & " Anillo Obsistens + 3 " & vbNewLine & vbNewLine & _
    " Torneo: " & vbNewLine & " Manto alado(Altos / Bajos) " & vbNewLine & " Tunica Apocalíptica(Altos / Bajos) " & vbNewLine & " Escudo Rubri Deum " & vbNewLine & " Espada Magnitud " & vbNewLine & " Anillo Ringonem " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Crural " & vbNewLine & " Anillo Bulgari " & vbNewLine & " Anillo Mobius " & vbNewLine & " Jaula de Guerra " & vbNewLine & " Casco Nocturno " & vbNewLine & " Báculo Dorado " & vbNewLine & " Espada de Hielo " & vbNewLine & " Daga Letum " & vbNewLine & " Daga Serpens " & vbNewLine & " Muro de Bronce " & vbNewLine & " Escudo de Hielo " & vbNewLine & " Escudo de Hierro " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Manto de la oscuridad " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Armadura Ceremonial "

  Case "Clérigo"
    Text1.Text = "Herreria: " & vbNewLine & " Armadura Penye (Altos/Bajos) " & vbNewLine & " Armadura Leorica (Altos/Bajos) " & vbNewLine & " Escudo Asedio " & vbNewLine & " Corona Oscura " & vbNewLine & " Casco Imperial " & vbNewLine & " Anillo Obsistens + 2 " & vbNewLine & vbNewLine & _
    " Torneo: " & vbNewLine & " Manto Alado(Altos / bajos) " & vbNewLine & " Túnica apocalíptica(Altos / bajos) " & vbNewLine & " Armadura Héroes (Altos) " & vbNewLine & " Armadura Parietis (Bajos) " & vbNewLine & " Escudo Apicibus " & vbNewLine & " Casco isolubilia " & vbNewLine & " Anillo regonem " & vbNewLine & " Espada Argentum " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Crural " & vbNewLine & " Anillo Bulgari " & vbNewLine & " Anillo Mobius " & vbNewLine & " Casco Prismatico " & vbNewLine & " Casco de Bronce " & vbNewLine & " Casco Nocturno " & vbNewLine & " Espada Excalibur " & vbNewLine & " Espada de Hielo " & vbNewLine & " Escudo del Rey " & vbNewLine & " Escudo de Hielo " & vbNewLine & " Escudo de Hierro " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Piel de Dragón Verde " & vbNewLine & " Armadura Adamantina " & vbNewLine & " Armadura Ceremonial "
    
  Case "Cazador"
    Text1.Text = "Herreria: " & vbNewLine & _
    " Todas las armaduras (Bajos) " & vbNewLine & " Escudo de León " & vbNewLine & " Escudo Inuro " & vbNewLine & " Casco Sangrio " & vbNewLine & " Casco Nazgul " & vbNewLine & " Hacha maligna " & vbNewLine & " Espada Rubrum dot " & vbNewLine & " Arco brave " & vbNewLine & " Anillo Obsistens + 1 " & vbNewLine & vbNewLine & _
    " Torneo: " & vbNewLine & " Manto Alado(Altos / Bajos) " & vbNewLine & " Tunica apocalíptica(Altos / Bajos) " & vbNewLine & " Armadura Nigro Fire (Altos / Bajos) " & vbNewLine & " Escudo Contorgueo " & vbNewLine & " Casco Dilectio " & vbNewLine & " Espada Incensa " & vbNewLine & " Arco Archer " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Mobius " & vbNewLine & " Casco Prismatico " & vbNewLine & " Casco Frodo " & vbNewLine & " Casco Nocturno " & vbNewLine & " Arco Oneida " & vbNewLine & " Espada de Hielo " & vbNewLine & " Escudo de Hielo " & vbNewLine & " Escudo de Hierro " & vbNewLine & " Armadura Trueno " & vbNewLine & " Armadura Lancelot " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Armadura Ceremonial "
    
  Case "Druida"
    Text1.Text = "Herreria: " & vbNewLine & " Todas las túnicas (Altos/Bajos) " & vbNewLine & " Escudo Tegimen " & vbNewLine & " Sombrero de mago " & vbNewLine & " Anillo Obsistens + 3 " & vbNewLine & " Daga Quodar " & vbNewLine & " Espada Vikinga " & vbNewLine & " Anillo Obsistens + 3 " & vbNewLine & vbNewLine & _
    " Torneo: " & vbNewLine & " Manto Alado(Altos / Bajos) " & vbNewLine & " Tunica Apocalíptica(Altos / Bajos) " & vbNewLine & " Escudo Áurea " & vbNewLine & " Bastón Furoris " & vbNewLine & " Corona de los dioses " & vbNewLine & " Anillo Ringonem " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Crural " & vbNewLine & " Anillo Bulgari " & vbNewLine & " Anillo Mobius " & vbNewLine & " Jaula de Guerra " & vbNewLine & " Casco Nocturno " & vbNewLine & " Báculo Mágico " & vbNewLine & " Báculo Dorado " & vbNewLine & " Espada de Hielo " & vbNewLine & " Daga Serpens " & vbNewLine & " Muro de Bronce " & vbNewLine & " Escudo de Hielo " & vbNewLine & " Escudo de Hierro " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Manto de la oscuridad " & vbNewLine & " Armadura Ceremonial "
    
  Case "Guerrero"
    Text1.Text = "Herreria: " & vbNewLine & " Todas las armaduras (Bajos) " & vbNewLine & " Escudo de León " & vbNewLine & " Escudo Inuro " & vbNewLine & " Casco Sangrio " & vbNewLine & " Casco Nazgul " & vbNewLine & " Hacha maligna " & vbNewLine & " Espada Rubrum dot " & vbNewLine & " Arco brave " & vbNewLine & " Anillo Obsistens + 1 " & vbNewLine & vbNewLine & _
    " Torneo: " & vbNewLine & " Tunica Apocalíptica(Altos / Bajos) " & vbNewLine & " Armadura Nigro Fire (Altos/Bajos) " & vbNewLine & " Escudo Contorgueo " & vbNewLine & " Casco Dilectio " & vbNewLine & " Espada Incensa " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Mobius " & vbNewLine & " Casco Prismatico " & vbNewLine & " Casco Frodo " & vbNewLine & " Casco Nocturno " & vbNewLine & " Arco Oneida " & vbNewLine & " Sable de luz " & vbNewLine & " Espada de Hielo " & vbNewLine & " Escudo Guerrero " & vbNewLine & " Escudo de Hielo " & vbNewLine & " Escudo de Hierro " & vbNewLine & " Armadura Trueno " & vbNewLine & " Armadura Lancelot " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Armadura Ceremonial "
    
  Case "Mago"
    Text1.Text = "Herreria: " & vbNewLine & " Todas las túnicas (Altos/Bajos)  " & vbNewLine & " Báculo Ancestral " & vbNewLine & " Sombrero de mago " & vbNewLine & " Corona oscura " & vbNewLine & " Anillo Obsistens + 3 " & vbNewLine & vbNewLine & _
    " Torneo " & vbNewLine & " Manto Alado(Altos / Bajos) " & vbNewLine & " Tunica apocalíptica(Altos / Bajos) " & vbNewLine & " Corona de torneo " & vbNewLine & " Baculo Sagrado " & vbNewLine & " Anillo Ringonem " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Crural " & vbNewLine & " Anillo Bulgar" & vbNewLine & " Anillo Mobius" & vbNewLine & " Báculo Mágico " & vbNewLine & " Báculo Dorado " & vbNewLine & " Bastón de Mago "

  Case "Paladin"
    Text1.Text = "Herreria: " & vbNewLine & " Armadura Noctem (Altos) " & vbNewLine & " Armadura Paludem (Altos/Bajos) " & vbNewLine & " Armadura Ingens (Bajos) " & vbNewLine & " Escudo Palatin " & vbNewLine & " Escudo Cuadrato " & vbNewLine & " Casco Imperial " & vbNewLine & " Casco Vikingo " & vbNewLine & " Rompe casco " & vbNewLine & " Espada Asperum " & vbNewLine & " Anillo obsistens +1 " & vbNewLine & vbNewLine & _
    " Torneo: " & vbNewLine & " Manto Alado(Altos / bajos) " & vbNewLine & " Túnica apocalíptica(Altos / bajos) " & vbNewLine & " Armadura Nigro Fire (Altos/Bajos) " & vbNewLine & " Escudo dragón " & vbNewLine & " Casco Legendario " & vbNewLine & " Espada Incensa " & vbNewLine & vbNewLine & _
    " Ítems a la venta: " & vbNewLine & " Anillo Mobius " & vbNewLine & " Casco Prismatico " & vbNewLine & " Casco Frodo " & vbNewLine & " Casco Nocturno " & vbNewLine & " Sable de luz " & vbNewLine & " Espada de Hielo " & vbNewLine & " Escudo del Rey " & vbNewLine & " Escudo de Hielo " & vbNewLine & " Escudo de Hierro " & vbNewLine & " Armadura Trueno " & vbNewLine & " Armadura Lancelot " & vbNewLine & " Armadura Justiciero " & vbNewLine & " Armadura Ceremonial "
    
 End Select
End With

End Sub

