Attribute VB_Name = "Extra"
'================================
'>>>>> WWW.FADICTOS.COM.AR <<<<<<
'================================
'Matute - matius_xd07@hotmail.com

Option Explicit
Public Function EsNewbie(UserIndex As Integer) As Boolean

EsNewbie = (UserList(UserIndex).Stats.ELV <= LimiteNewbie)

End Function
Public Sub DoTileEvents(UserIndex As Integer)
On Error GoTo errhandler
Dim Map As Integer, X As Integer, Y As Integer
Dim nPos As WorldPos, mPos As WorldPos

Map = UserList(UserIndex).POS.Map
X = UserList(UserIndex).POS.X
Y = UserList(UserIndex).POS.Y

mPos = MapData(Map, X, Y).TileExit
If Not MapaValido(mPos.Map) Or Not InMapBounds(mPos.X, mPos.Y) Then Exit Sub

If MapData(Map, X, Y).TileExit.Map = 64 Then
If Not UserList(UserIndex).flags.VIP = 1 Then
Call SendData(ToIndex, UserIndex, 0, "||Solo los donadores pueden entrar a este mapa" & FONTTYPE_FENIX)
Exit Sub
End If
End If

If MapInfo(mPos.Map).Restringir And Not EsNewbie(UserIndex) Then
    Call SendData(ToIndex, UserIndex, 0, "1J")
ElseIf UserList(UserIndex).Stats.ELV < MapInfo(mPos.Map).Nivel And Not (UserList(UserIndex).Clase = PIRATA And UserList(UserIndex).Recompensas(1) = 2) Then
    Call SendData(ToIndex, UserIndex, 0, "%/" & MapInfo(mPos.Map).Nivel)
Else
    If LegalPos(mPos.Map, mPos.X, mPos.Y, PuedeAtravesarAgua(UserIndex)) Then
        If mPos.X <> 0 And mPos.Y <> 0 Then Call WarpUserChar(UserIndex, mPos.Map, mPos.X, mPos.Y, ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)
    Else
        Call ClosestStablePos(mPos, nPos)
        If nPos.X <> 0 And nPos.Y Then Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)
    End If
    Exit Sub
End If

Call ClosestStablePos(UserList(UserIndex).POS, nPos)
If nPos.X <> 0 And nPos.Y Then Call WarpUserChar(UserIndex, nPos.Map, nPos.X, nPos.Y, ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT)

Exit Sub

errhandler:
    Call LogError("Error en DoTileEvents-" & nPos.Map & "-" & nPos.X & "-" & nPos.Y)
End Sub
Function InMapBounds(X As Integer, Y As Integer) As Boolean

InMapBounds = (X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder)

End Function
Sub ClosestStablePos(POS As WorldPos, ByRef nPos As WorldPos)
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.X, nPos.Y)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.X - LoopC To POS.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
                nPos.X = tX
                nPos.Y = tY

                tX = POS.X + LoopC
                tY = POS.Y + LoopC
  
            End If
        
        Next tX
    Next tY
    
    LoopC = LoopC + 1
    
Loop

If Notfound = True Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub
Sub ClosestLegalPos(POS As WorldPos, nPos As WorldPos, Optional AguaValida As Boolean)
Dim Notfound As Boolean
Dim LoopC As Integer
Dim tX As Integer
Dim tY As Integer

nPos.Map = POS.Map

Do While Not LegalPos(POS.Map, nPos.X, nPos.Y, AguaValida)
    If LoopC > 12 Then
        Notfound = True
        Exit Do
    End If
    
    For tY = POS.Y - LoopC To POS.Y + LoopC
        For tX = POS.X - LoopC To POS.X + LoopC
            
            If LegalPos(nPos.Map, tX, tY, AguaValida) Then
                nPos.X = tX
                nPos.Y = tY
                
                
                tX = POS.X + LoopC
                tY = POS.Y + LoopC
  
            End If
        
        Next
    Next
    
    LoopC = LoopC + 1
    
Loop

If Notfound Then
    nPos.X = 0
    nPos.Y = 0
End If

End Sub
Function ClaseIndex(ByVal Clase As String) As Integer
Dim i As Integer

For i = 1 To UBound(ListaClases)
    If UCase$(ListaClases(i)) = UCase$(Clase) Then
        ClaseIndex = i
        Exit Function
    End If
Next

End Function
Function NameIndex(ByVal Name As String) As Integer
Dim UserIndex As Integer, i As Integer

Name = Replace$(Name, "+", " ")

If Len(Name) = 0 Then
    NameIndex = 0
    Exit Function
End If
  
UserIndex = 1

If Right$(Name, 1) = "*" Then
    Name = Left$(Name, Len(Name) - 1)
    For i = 1 To LastUser
        If UCase$(UserList(i).Name) = UCase$(Name) Then
            NameIndex = i
            Exit Function
        End If
    Next
Else
    For i = 1 To LastUser
        If UCase$(Left$(UserList(i).Name, Len(Name))) = UCase$(Name) Then
            NameIndex = i
            Exit Function
        End If
    Next
End If

End Function
Function CheckForSameIP(UserIndex As Integer, ByVal UserIP As String) As Boolean
Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).flags.UserLogged Then
        If UserList(LoopC).ip = UserIP And UserIndex <> LoopC Then
            CheckForSameIP = True
            Exit Function
        End If
    End If
Next

End Function
Function CheckForSameName(UserIndex As Integer, ByVal Name As String) As Boolean
Dim LoopC As Integer

For LoopC = 1 To LastUser
    If UserList(LoopC).flags.UserLogged Then
        If UCase$(UserList(LoopC).Name) = UCase$(Name) Then
            CheckForSameName = True
            Exit Function
        End If
    End If
Next

End Function
Sub HeadtoPos(Head As Byte, POS As WorldPos)
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = POS.X
Y = POS.Y

If Head = NORTH Then
    nX = X
    nY = Y - 1
End If

If Head = SOUTH Then
    nX = X
    nY = Y + 1
End If

If Head = EAST Then
    nX = X + 1
    nY = Y
End If

If Head = WEST Then
    nX = X - 1
    nY = Y
End If

POS.X = nX
POS.Y = nY

End Sub
Function LegalPos(Map As Integer, X As Integer, Y As Integer, Optional PuedeAgua As Boolean) As Boolean

If Not MapaValido(Map) Or Not InMapBounds(X, Y) Then Exit Function

LegalPos = (MapData(Map, X, Y).Blocked = 0) And _
           (MapData(Map, X, Y).UserIndex = 0) And _
           (MapData(Map, X, Y).NpcIndex = 0) And _
           (MapData(Map, X, Y).Agua = Buleano(PuedeAgua))

End Function
Function LegalPosNPC(Map As Integer, X As Integer, Y As Integer, AguaValida As Boolean) As Boolean

If Not InMapBounds(X, Y) Then Exit Function

LegalPosNPC = (MapData(Map, X, Y).Blocked <> 1) And _
     (MapData(Map, X, Y).UserIndex = 0) And _
     (MapData(Map, X, Y).NpcIndex = 0) And _
     (MapData(Map, X, Y).trigger <> POSINVALIDA) _
     And Buleano(AguaValida) = MapData(Map, X, Y).Agua
     
End Function
Public Sub SendNPC(UserIndex As Integer, NpcIndex As Integer)
Dim Info As String
Dim CRI As Byte

Select Case UserList(UserIndex).Stats.UserSkills(Supervivencia)
    Case Is <= 20
        If Npclist(NpcIndex).Stats.MinHP = Npclist(NpcIndex).Stats.MaxHP Then
            CRI = 5
        Else: CRI = 1
        End If
    Case Is < 40
        Select Case 100 * Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP
            Case 100
                CRI = 5
            Case Is >= 50
                CRI = 2
            Case Else
                CRI = 3
        End Select
    Case Is < 60
        Select Case 100 * Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP
            Case 100
                CRI = 5
            Case Is > 66
                CRI = 2
            Case Is > 33
                CRI = 3
            Case Else
                CRI = 4
        End Select
    Case Is < 100
        CRI = 5 + Fix(10 * (1 - (Npclist(NpcIndex).Stats.MinHP / Npclist(NpcIndex).Stats.MaxHP)))
    Case Else
        Info = "||" & Npclist(NpcIndex).Name & " [" & Npclist(NpcIndex).Stats.MinHP & "/" & Npclist(NpcIndex).Stats.MaxHP & "]"
        If Npclist(NpcIndex).flags.Paralizado Then Info = Info & " - PARALIZADO"
        Call SendData(ToIndex, UserIndex, 0, Info & FONTTYPE_INFO)
        Exit Sub
End Select

Info = "9Q" & Npclist(NpcIndex).Name & "," & CRI
Call SendData(ToIndex, UserIndex, 0, Info)
                
End Sub
Public Sub Expresar(NpcIndex As Integer, UserIndex As Integer)

If Npclist(NpcIndex).NroExpresiones Then
    Dim randomi
    randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
    Call SendData(ToPCArea, UserIndex, UserList(UserIndex).POS.Map, "3Q" & vbWhite & "�" & Npclist(NpcIndex).Expresiones(randomi) & "�" & Npclist(NpcIndex).Char.CharIndex)
End If
                    
End Sub
Sub LookatTile(UserIndex As Integer, Map As Integer, X As Integer, Y As Integer)

Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim NPMUERTO As String
Dim Info As String


If InMapBounds(X, Y) Then
    UserList(UserIndex).flags.TargetMap = Map
    UserList(UserIndex).flags.TargetX = X
    UserList(UserIndex).flags.TargetY = Y
    
    If MapData(Map, X, Y).OBJInfo.OBJIndex Then
        
        If MapData(Map, X, Y).OBJInfo.Amount = 1 Then
            Call SendData(ToIndex, UserIndex, 0, "4Q" & ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Name)
        Else
            Call SendData(ToIndex, UserIndex, 0, "5Q" & ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Name & "," & MapData(Map, X, Y).OBJInfo.Amount)
        End If
        UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex
        UserList(UserIndex).flags.TargetObjMap = Map
        UserList(UserIndex).flags.TargetObjX = X
        UserList(UserIndex).flags.TargetObjY = Y
        FoundSomething = 1
    ElseIf MapData(Map, X + 1, Y).OBJInfo.OBJIndex Then
        
        If ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            Call SendData(ToIndex, UserIndex, 0, "6Q" & ObjData(MapData(Map, X + 1, Y).OBJInfo.OBJIndex).Name)
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y).OBJInfo.OBJIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex Then
        If ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            
            Call SendData(ToIndex, UserIndex, 0, "6Q" & ObjData(MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex).Name)
            UserList(UserIndex).flags.TargetObj = MapData(Map, X + 1, Y + 1).OBJInfo.OBJIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X + 1
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    ElseIf MapData(Map, X, Y + 1).OBJInfo.OBJIndex Then
        If ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).ObjType = OBJTYPE_PUERTAS Then
            
            Call SendData(ToIndex, UserIndex, 0, "6Q" & ObjData(MapData(Map, X, Y + 1).OBJInfo.OBJIndex).Name)
            UserList(UserIndex).flags.TargetObj = MapData(Map, X, Y).OBJInfo.OBJIndex
            UserList(UserIndex).flags.TargetObjMap = Map
            UserList(UserIndex).flags.TargetObjX = X
            UserList(UserIndex).flags.TargetObjY = Y + 1
            FoundSomething = 1
        End If
    End If
    
    If Y + 1 <= YMaxMapSize Then
        If MapData(Map, X, Y + 1).UserIndex Then
            TempCharIndex = MapData(Map, X, Y + 1).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y + 1).NpcIndex Then
            TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
            FoundChar = 2
        End If
    End If
    
    If FoundChar = 0 Then
        If MapData(Map, X, Y).UserIndex Then
            TempCharIndex = MapData(Map, X, Y).UserIndex
            FoundChar = 1
        End If
        If MapData(Map, X, Y).NpcIndex Then
            TempCharIndex = MapData(Map, X, Y).NpcIndex
            FoundChar = 2
        End If
    End If
    
    
    
    If FoundChar = 1 Then
            
        If UserList(TempCharIndex).flags.AdminInvisible Then Exit Sub
        
        If UserList(TempCharIndex).Faccion.Bando Then
            If UserList(TempCharIndex).Faccion.BandoOriginal <> UserList(TempCharIndex).Faccion.Bando Then
                Stat = Stat & " <" & ListaBandos(UserList(TempCharIndex).Faccion.Bando) & "> <Mercenario>"
            ElseIf UserList(TempCharIndex).Faccion.Jerarquia Then
                Stat = Stat & " <" & ListaBandos(UserList(TempCharIndex).Faccion.Bando) & "> <" & Titulo(TempCharIndex) & ">"
            Else
                Stat = Stat & " <" & Titulo(TempCharIndex) & ">"
            End If
        End If
        
        If UserList(TempCharIndex).flags.Privilegios < 1 Then 'SI NO ES GM
        Stat = Stat & " <Clase: " & ListaClases(UserList(TempCharIndex).Clase) & ">" 'DICE LA CLASE
        End If
        
        If UserList(TempCharIndex).flags.VIP = 1 Then
                Stat = Stat & " <Donador> ~255~255~0~1~0"
        End If
        
        If Len(UserList(TempCharIndex).GuildInfo.GuildName) > 0 Then
            Stat = Stat & " <" & UserList(TempCharIndex).GuildInfo.GuildName & ">"
        End If
        
        If Len(UserList(TempCharIndex).Desc) > 0 Then
            Stat = UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
        Else
            Stat = UserList(TempCharIndex).Name & Stat
        End If
        
            If UserList(TempCharIndex).flags.Canje > 0 Then
                Stat = Stat & " <Canjes: " & UserList(TempCharIndex).flags.Canje & ">"
            End If
        
If UserList(TempCharIndex).flags.Privilegios Then
            Stat = "Ves a " & Stat
        Else
            If UserList(TempCharIndex).flags.Muerto Then
                Stat = "Ves a " & UserList(TempCharIndex).Name
            ElseIf UserList(TempCharIndex).Faccion.Bando = Real Then
                Stat = "Ves a " & Stat
            ElseIf UserList(TempCharIndex).Faccion.Bando = Caos Then
                Stat = "Ves a " & Stat
            ElseIf EsNewbie(TempCharIndex) Then
                Stat = "Ves a " & Stat
            ElseIf UserList(TempCharIndex).flags.VIP = 1 Then
                Stat = "Ves a " & Stat
            Else
                Stat = "Ves a " & Stat
            End If
        End If
        
        If UserList(TempCharIndex).flags.Privilegios = 3 Then
            Stat = Stat & " <Dios> ~255~128~64~1~0"
        ElseIf UserList(TempCharIndex).flags.Privilegios = 2 Then
            Stat = Stat & " <Semi-Dios> ~200~88~209~1~0"
        ElseIf UserList(TempCharIndex).flags.Privilegios = 1 Then
            Stat = Stat & " <Consejero> ~0~255~255~1~0"
        Else
            If UserList(TempCharIndex).flags.Muerto Then
                Stat = Stat & " <Muerto> ~255~255~255~1~0"
            ElseIf UserList(TempCharIndex).Faccion.Bando = Caos Then
                Stat = Stat & " ~255~0~0~1~0"
            ElseIf UserList(TempCharIndex).Faccion.Bando = Real Then
                Stat = Stat & " ~0~128~255~1~0"
            ElseIf EsNewbie(TempCharIndex) Then
                Stat = Stat & " <Newbie> ~0~240~0~1~0"
            ElseIf UserList(TempCharIndex).flags.VIP = 1 Then
                Stat = Stat & " <Donador> ~255~255~0~1~0 "
            Else
                Stat = Stat & " <Neutral> ~190~190~190~1~0"
        End If
        End If
        Call SendData(ToIndex, UserIndex, 0, "||" & Stat)
        
            
        
        FoundSomething = 1
        UserList(UserIndex).flags.TargetUser = TempCharIndex
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
       
       
    ElseIf FoundChar = 2 Then
            
            Dim wPos As WorldPos
            wPos.Map = Map
            wPos.X = X
            wPos.Y = Y
            If Distancia(Npclist(TempCharIndex).POS, wPos) > 1 Then
                MapData(Map, X, Y).NpcIndex = 0
                Exit Sub
            End If
                
            If Npclist(TempCharIndex).flags.TiendaUser Then
                If UserIndex = Npclist(TempCharIndex).flags.TiendaUser Then
                    If UserList(UserIndex).Tienda.Gold Then
                        Call SendData(ToIndex, UserIndex, 0, "/O" & UserList(UserIndex).Tienda.Gold & "," & Npclist(TempCharIndex).Char.CharIndex)
                    Else
                        Call SendData(ToIndex, UserIndex, 0, "/P" & Npclist(TempCharIndex).Char.CharIndex)
                    End If
                Else
                    Call SendData(ToIndex, UserIndex, 0, "/Q" & UserList(Npclist(TempCharIndex).flags.TiendaUser).Name & "," & Npclist(TempCharIndex).Char.CharIndex)
                End If
            ElseIf Len(Npclist(TempCharIndex).Desc) > 1 Then
                Call SendData(ToIndex, UserIndex, 0, "3Q" & vbWhite & "�" & Npclist(TempCharIndex).Desc & "�" & Npclist(TempCharIndex).Char.CharIndex)
            ElseIf Npclist(TempCharIndex).MaestroUser Then
                Call SendData(ToIndex, UserIndex, 0, "7Q" & Npclist(TempCharIndex).Name & "," & UserList(Npclist(TempCharIndex).MaestroUser).Name)
            ElseIf Npclist(TempCharIndex).AutoCurar = 1 Then
                Call SendData(ToIndex, UserIndex, 0, "8Q" & Npclist(TempCharIndex).Name)
            Else
                Call SendNPC(UserIndex, TempCharIndex)
            End If
            FoundSomething = 1
            UserList(UserIndex).flags.TargetNpcTipo = Npclist(TempCharIndex).NPCtype
            UserList(UserIndex).flags.TargetNpc = TempCharIndex
            UserList(UserIndex).flags.TargetUser = 0
            UserList(UserIndex).flags.TargetObj = 0
    End If
    
    If FoundChar = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
    End If
    
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If

Else
    If FoundSomething = 0 Then
        UserList(UserIndex).flags.TargetNpc = 0
        UserList(UserIndex).flags.TargetNpcTipo = 0
        UserList(UserIndex).flags.TargetUser = 0
        UserList(UserIndex).flags.TargetObj = 0
        UserList(UserIndex).flags.TargetObjMap = 0
        UserList(UserIndex).flags.TargetObjX = 0
        UserList(UserIndex).flags.TargetObjY = 0
    End If
End If

End Sub
Function FindDirection(POS As WorldPos, Target As WorldPos) As Byte
Dim X As Integer, Y As Integer

X = POS.X - Target.X
Y = POS.Y - Target.Y

If Sgn(X) = -1 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = 1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = -1 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = -1 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = -1 Then
    FindDirection = SOUTH
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = 1 Then
    FindDirection = NORTH
    Exit Function
End If

If Sgn(X) = 1 And Sgn(Y) = 0 Then
    FindDirection = WEST
    Exit Function
End If

If Sgn(X) = -1 And Sgn(Y) = 0 Then
    FindDirection = EAST
    Exit Function
End If

If Sgn(X) = 0 And Sgn(Y) = 0 Then
    FindDirection = 0
    Exit Function
End If

End Function
Public Function ItemEsDeMapa(ByVal Map As Integer, X As Integer, Y As Integer) As Boolean

ItemEsDeMapa = ObjData(MapData(Map, X, Y).OBJInfo.OBJIndex).Agarrable Or MapData(Map, X, Y).Blocked

End Function

Sub PasarSegundo()
Dim mapa As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Integer
 
'listo, fijate si asi anda...
 
 
       
For i = 1 To LastUser
   mapa = UserList(i).flags.DondeTiroMap
X = UserList(i).flags.DondeTiroX
Y = UserList(i).flags.DondeTiroY
    If UserList(i).Counters.CreoTeleport = True Then  'si el usuario creo un teleport....
        UserList(i).Counters.TimeTeleport = UserList(i).Counters.TimeTeleport + 1 'sumamos 1 cont
 
        If UserList(i).Counters.TimeTeleport = 3 Then 'cuando llega a 3
            Call EraseObj(ToMap, 0, UserList(i).flags.DondeTiroMap, MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).OBJInfo.Amount, UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY) 'verificamos que destruye el objeto anterior.
            Dim ET As Obj
            ET.Amount = 1
            ET.OBJIndex = Teleport 'Ac� se puede cambiar por su telep personalizado
                       
            Call MakeObj(ToMap, 0, UserList(i).POS.Map, ET, UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY)
            MapData(UserList(i).POS.Map, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).TileExit.Map = 61 'llevams a ulla
            MapData(UserList(i).POS.Map, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).TileExit.X = 50
            MapData(UserList(i).POS.Map, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).TileExit.Y = 50
        ElseIf UserList(i).Counters.TimeTeleport >= 10 Then
            UserList(i).flags.TiroPortalL = 0
            UserList(i).Counters.TimeTeleport = 0
            UserList(i).Counters.CreoTeleport = False
            Call EraseObj(ToMap, 0, UserList(i).flags.DondeTiroMap, MapData(UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY).OBJInfo.Amount, UserList(i).flags.DondeTiroMap, UserList(i).flags.DondeTiroX, UserList(i).flags.DondeTiroY) 'verificamos que destruye el objeto anterior.
            MapData(mapa, X, Y).TileExit.Map = 0
            MapData(mapa, X, Y).TileExit.X = 0
            MapData(mapa, X, Y).TileExit.Y = 0
            UserList(i).flags.DondeTiroMap = ""
            UserList(i).flags.DondeTiroX = ""
            UserList(i).flags.DondeTiroY = ""
        End If
    End If
 
Next i
 
End Sub
