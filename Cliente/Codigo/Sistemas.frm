VERSION 5.00
Begin VB.Form Sistemas 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4050
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5625
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Sistemas.frx":0000
   ScaleHeight     =   4050
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   3015
      Left            =   100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Sistemas.frx":13705
      Top             =   980
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Sistemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Text1.Text = "Sistema de Guerras " & vbNewLine & "En breve descripci�n, nuestro sistema de Guerras 'automaticas', est� basado en una guerra en mapas faccionarios, cada 1 o 2 horas, cuando la guerra inicia un NPC faccionario spawnea en uno de los mapas en guerra, el el equipo que logre vencer y/o proteger el NPC, ser� recompensado con 200.000 monedas de oro y la lealt�d a su facci�n." & vbNewLine & "Para localiz�r el NPC, el minimapa de la parte inferior indicar� en color azul o rojo la posici�n del NPC ..." & vbNewLine & vbNewLine & "Sistema de Subastas" & vbNewLine & "El sistema de subastas, se lleva a cabo mediante el comando /subastar, que te aparecer� un formulario con la configuraci�n de subasta de item." & vbNewLine & _
"El usuario tendr� 3 minutos de subasta, para los usuarios del servidor, puedan /OFRECER (cantidad) y sea vendido el item, y para los usuarios que no saben que se esta subastando, tienen el comando /INFOSUB" & vbNewLine & vbNewLine & _
"AutoUpdate" & vbNewLine & "Tendremos incorporado un sistema de AutoUpdate, para descargar automaticamente, los parches nuevos de SWAO sin necesidad de tocar nada." & vbNewLine & vbNewLine & _
"Caballer�a" & vbNewLine & "Semper Widia Argentum, cuenta con un sistema de Caballer�a, el usuario Caballero, tendr� infinidades de beneficios, tanto como en torneos y oro, como en quests y puntos de torneo." & vbNewLine & _
"Hacerse caballero, consiste en conseguir los 4 items de caballer�a y 3 copas de oro (Se consiguen saliendo Primero en Torneos), y mediante el comando /CABALLERIA, lograr�s la caballer�a real, obteniendo beneficios al instante." & vbNewLine & _
"Los �tems requeridos adem�s de las 3 copas de oro, son los siguientes: " & vbNewLine & "Espada  Caballero Negro" & vbNewLine & "Corona  Caballero Negro" & vbNewLine & "Armadura  Caballero Negro" & vbNewLine & "Anillo  Caballero Negro" & vbNewLine & "�C�mo sab�r si un usuario es caballero?" & vbNewLine & _
"El usuario que sea caballero tendr� el nick en color verde claro y tendr� de Tag, < Caballero >." & vbNewLine & vbNewLine & _
"Salas de Invocaciones" & vbNewLine & "Tendremos una sala de invocaciones, llamada Silver Anguis, se encuentra en el mapa 8, es necesario 4 (o +) usuarios para pod�r invoc�r la bestia Silver Anguis Soul se tendr�n que poner uno en cada cuadro, y la bestia respawnear� en el punto medio, con 400.000 puntos de vida, Silver Anguis al morir tendr� la posiblidad de dropear uno de sus preciados items de caballero, as� podran conseguir los 4 items de caballer�a." & vbNewLine & _
"Pero tienen que tener ciudado, unicas criaturas, acechan el mapa y no son faciles de derrotar." & vbNewLine & vbNewLine & "Gemas" & vbNewLine & "Uno de los items mas preciados, las 8 gemas ..." & vbNewLine & "Naranja, Rojo, Azul, Celeste, Lila, Plateada, Verde, Violeta." & vbNewLine & "Las gemas son dropeadas por las siguientes criaturas:" & vbNewLine & "Drag�n Oscuro: Gema Lila" & vbNewLine & "Drag�n Plateado: Gema Plateada" & vbNewLine & "Drag�n Rojo: Gema Roja" & vbNewLine & "Golem: Gema Verde" & vbNewLine & "Bestia Infernal: Gema Naranja" & vbNewLine & "Yeti Polar: Gema Celeste" & vbNewLine & "Tenebrosi Magus: Gema Violeta" & vbNewLine & "Atrum Billfish: Gema Azul" & vbNewLine & "Est�s criaturas acechan en distintos dungeons." & vbNewLine & vbNewLine & _
"Creaci�n de Clan" & vbNewLine & "Para crear clan, necesitas 1 Fragmento Rojo" & vbNewLine & "Se consigue mezclado (/MEZCLAR) las 8 gemas, teniendo un 50% de posiblidad de que se cree y otro 50% de obtener 300 puntos de torneo." & vbNewLine & vbNewLine & "Switch en Inventario" & vbNewLine & "Hemos incluido en nuestra versi�n, un switch de items, para pod�r acomodar el inventario sin necesidad de tirar los items al piso ..." & vbNewLine & "Para us�r esta funcion, es necesario hac�r click derecho y arrastrar el item al slot desado, devolviendo el item del slot deseado al slot del item corrido, haciendo Switch de Items." & vbNewLine & vbNewLine & "Mapa del Mundo" & vbNewLine & "Tenemos incorporado un Mapa del Mundo, si est�s desorientado o no sab�s para que lado ir, tocando la Tecla J se abrir� un mapa donde les ilustrar� el mundo, as� pueden ir a sus lugares preferidos." & vbNewLine & vbNewLine & _
"Quests" & vbNewLine & "Est�s sin oro y sin ningun item?, necesesit�s mas de puntos de torneo?" & vbNewLine & "Tenemos un nuevo sistema de quests, que te perminte selecionar entre variados tipos de quests para completarlas y gan�r el oro y/o puntos de torneo de tal quest." & vbNewLine & "El npc de las quest, se encuentra en el mapa inicial Troyes [1], abajo de la fuente del centro de la ciudad." & vbNewLine & vbNewLine & "Sistema de Rankings" & vbNewLine & "El servidor cuenta con un sistema de rankings, donde se encontraran los usuarios que .." & vbNewLine & "Mas frags tengan." & vbNewLine & "Mas torneos ganados tengan." & vbNewLine & "Mas usuarios matados." & vbNewLine & "Mas 2v2 ganados." & vbNewLine & "Mas rondas en desafio." & vbNewLine & "Entre otros ..., para informarte del ranking, se usa el comando /RANKING" & vbNewLine & vbNewLine & "Sistema de Cuentas" & vbNewLine & _
"El servidor cuenta con un gran sistema de cuentas, en est� mismo soporta 10 personajes por cada cuenta. Solo se puede conectar un personaje por cuenta. Tambi�n podr�n eliminar personajes (menores a n�vel 45)." & vbNewLine & vbNewLine & "Sistema de Ciruj�as" & vbNewLine & "Mediante el Npc cirujano, podr�n usar el comando '/cirujia', elegir mediante un formulario la cara que deseen y obtener su nuevo rostro ! (costo m�nimo: 1.000 monedas de oro)." & vbNewLine & vbNewLine & "Sistema de Macros Configurables" & vbNewLine & "Para la comodidad de los usuarios, hemos incorporado un sistema de macros configurables, el cual podr�n configurarlo abriendo el formulario con la tecla 'F2'.." & vbNewLine & "�C�mo usarlo?, con las teclas '1;2;3;4;5;6;7;8;9;0' (n�meros �bicados arriba de las letras)." & vbNewLine & vbNewLine & _
"Dungeones" & vbNewLine & "�C�mo llegar a los dungeones?" & vbNewLine & "Dungeon Silver Anguis, tienen la entrada del mismo en el mapa 8." & vbNewLine & "Dungeon Piatra, tienen la entrada del mismo en el mapa 13." & vbNewLine & "Dungeon Dover, tienen la entrada del mismo en el mapa 39." & vbNewLine & "Dungeon Herakle�polis, tienen la entrada del mismo en el mapa 66." & vbNewLine & "Cueva polar, tienen la entrada del mismo en el mapa 62." & vbNewLine & vbNewLine & "Sistema de Monturas" & vbNewLine & "Las monturas ser�n dropeadas por las siguientes criaturas:" & vbNewLine & "'Drag�n Oscuro' ; 'Drag�n Plateado' ; 'Tenebrosi Magus' ; 'Moxostoma pugnator' ; 'Golem', pero ojo..solo ser�n dropeadas a usuarios que sean caballeros (Requiere estar convertido)." & vbNewLine & "Cuando una criatura dropee la 'montura', ser� el gr�fico de un pergamino, al cual le tienen que hacer doble click (teniendolo en nuestro inventario), y aprenderan un hechizo, y con este podr�n equiparse la montura." & vbNewLine & _
vbNewLine & "Sistema de Evento" & vbNewLine & "Cada 8 horas spawnea una criatura (Imsety) en un mapa random, con 150.000 puntos de vida, junta a esta tambien spawnear�n otras 8 criaturas (Wetyw) con 15.000 puntos de vida c/u." & vbNewLine & "Nota:" & vbNewLine & "Cuando la criatura spawnea, se da un aviso por consola general a todos los usuarios." & vbNewLine & "La criatura 'Imsety', al morir finalizar� el evento, y al usuario que logre combatirla tendr� la posibilidad de que le dropee una montura (por mas que el usuario NO sea caballero, le dropea igual), adem�s le d� 200 puntos de torneo y una gran cantidad de oro." & vbNewLine & vbNewLine & _
"Sistema de Trabajadores" & vbNewLine & "� Le�adores y Carpinteros" & vbNewLine & "Zonas de talaci�n: Todos los mapas que posean �rboles." & vbNewLine & "�tems: Para averiguar los �tems que construye el carpintero, deber�n tener un serrucho, equiparselo y hacerle doble click, a continuaci�n les aparecer� un cuadro donde podr�n ver toda la informaci�n necesaria." & vbNewLine & "� Mineros y Herreros" & vbNewLine & "Zonas de mineria: Hierro: Mapa 71 (Catacumbas, entrada en mapa 4) ; Plata: Mapa 67 (Cueva Polar, entrada en mapa 62) ; Oro: Mapa 66 (Dungeon Herakle�polis, entrada en mapa 39)" & vbNewLine & "�tems: Para averiguar los �tems que construye el herrero, deber�n tener un martillo, equiparselo, hacerle doble click y luego click a un yunque, a continuaci�n les aparecer� un cuadro donde podr�n ver toda la informaci�n necesaria." & vbNewLine & vbNewLine & _
"Para m�s info sobre el servidor, visit� nuestro foro [ www.swforos.com ]"
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
