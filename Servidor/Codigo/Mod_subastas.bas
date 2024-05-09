Attribute VB_Name = "Mod_subastas"
Option Explicit
 
Type S
    HaySubastas         As Boolean
    Vendedor            As Integer
    comprador           As Integer
    oferta              As Long
    ItemEnVenta         As Integer
    CantidadVenta       As Integer
    VendedorQuisoSalir  As Byte
    CompradorQuisoSalir As Byte
End Type
 
Public Const NPCSubastas As Byte = 55
Public Const Duracion_Subasta As Byte = 2
Public Subastas As S
 
Public SegundosSubasta As Byte
Public MinutosSubasta As Byte
 
 
Sub PasarMinutoSubasta()
 
 
If Subastas.HaySubastas = True Then
    MinutosSubasta = MinutosSubasta + 1
        If MinutosSubasta = (Duracion_Subasta) Then
            TerminarSubasta
            MinutosSubasta = 0
        Else
            If Subastas.comprador = 0 Then
                SendData ToAll, 0, 0, "||El usuario " & UserList(Subastas.Vendedor).Name & " esta subastando " & Subastas.CantidadVenta & " " & _
                ObjData(Subastas.ItemEnVenta).Name & _
                " a un precio inicial de " & Subastas.oferta & "." & FONTTYPE_FENIX
            Else
                SendData ToAll, 0, 0, "||El usuario " & UserList(Subastas.Vendedor).Name & " esta subastando " & Subastas.CantidadVenta & " " & _
                ObjData(Subastas.ItemEnVenta).Name & _
                " a un precio de " & Subastas.oferta & "." & FONTTYPE_FENIX
            End If
            SendData ToAll, 0, 0, "||La Subasta terminara en " & (Duracion_Subasta - MinutosSubasta) & " Minutos." & FONTTYPE_FENIX
        End If
End If
 
 
End Sub
 
Sub TerminarSubasta()
 
Dim ob As Obj
 
With Subastas
 
    ob.Amount = .CantidadVenta
    ob.OBJIndex = .ItemEnVenta
   
If .comprador = 0 Then
 '   Call SendData(ToIndex, .Vendedor, 0, "||Nadie respondio a tu subasta" & FONTTYPE_INFO)
        If Not MeterItemEnInventario(.Vendedor, ob) Then Call TirarItemAlPiso(UserList(.Vendedor).POS, ob)
            SendData ToAll, 0, 0, "||Subasta finalizada." & FONTTYPE_FENIX
            Else
            UserList(.Vendedor).Stats.GLD = UserList(.Vendedor).Stats.GLD + .oferta
            Call SendUserORO(.Vendedor)
            If Not MeterItemEnInventario(.comprador, ob) Then Call TirarItemAlPiso(UserList(.comprador).POS, ob)
                 SendData ToAll, 0, 0, "||La subasta termino, el ganador fue " & UserList(.comprador).Name & " a un precio de " & _
                .oferta & " monedas de oro." & FONTTYPE_FENIX
   
End If
 
.HaySubastas = False
.Vendedor = 0
.comprador = 0
.oferta = 0
.ItemEnVenta = 0
.CantidadVenta = 0
.CompradorQuisoSalir = 0
.VendedorQuisoSalir = 0
 
 If Subastas.CompradorQuisoSalir = 1 Then CloseSocket Subastas.comprador
 If Subastas.VendedorQuisoSalir = 1 Then CloseSocket Subastas.Vendedor
 
End With
 
End Sub
 
Sub Subastar(Userindex As Integer, Precioinicial As Long)
 
Dim npc As Integer
 
npc = UserList(Userindex).flags.TargetNpc
 
 If npc = 0 Then Exit Sub
 
If Npclist(npc).NPCtype <> NPCSubastas Then
Exit Sub
End If
 
Dim Itemsubastar As Integer
Dim CantidadItemSubastar As Integer
 
Itemsubastar = MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).OBJInfo.OBJIndex
CantidadItemSubastar = MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).OBJInfo.Amount
 
If Distancia(Npclist(npc).POS, UserList(Userindex).POS) > 1 Then
Call SendData(ToIndex, Userindex, 0, "||Estas Muy Lejos" & FONTTYPE_INFO)
Exit Sub
End If
 
If CantidadItemSubastar = 0 Then
SendData ToIndex, Userindex, 0, "||Tira el item q deseas subastar" & FONTTYPE_INFO
Exit Sub
End If
 
'ObjData (MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).OBJInfo.OBJIndex)
Dim ob As Obj
 
EraseObj ToMap, Userindex, UserList(Userindex).POS.Map, _
MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).OBJInfo.Amount, _
UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y
 
 
 
Subastas.Vendedor = Userindex
Subastas.ItemEnVenta = Itemsubastar
Subastas.CantidadVenta = CantidadItemSubastar
Subastas.oferta = Precioinicial
Subastas.comprador = 0
Subastas.oferta = Precioinicial
SegundosSubasta = 0
 
 
SendData ToAll, 0, 0, "||El usuario " & UserList(Userindex).Name & " esta subastando " & CantidadItemSubastar & " " & _
ObjData(Itemsubastar).Name & _
" a un precio inicial de " & Precioinicial & "." & FONTTYPE_FENIX
 
 
Subastas.HaySubastas = True
 
End Sub
 
Sub Ofertar(Userindex As Integer, oferta As Long)
 
 
With Subastas
 
If .HaySubastas = False Then
SendData ToIndex, Userindex, 0, "||No hay ninguna subasta en este momento" & FONTTYPE_INFO
Exit Sub
End If
 
 
If UserList(Userindex).Stats.GLD < oferta Then
SendData ToIndex, Userindex, 0, "||No tienes esa cantidad" & FONTTYPE_INFO
Exit Sub
End If
 
 
    If oferta <= .oferta Then
    SendData ToIndex, Userindex, 0, "||Tu Oferta deve ser mayor a " & .oferta & " monedas de oro." & FONTTYPE_INFO
    Exit Sub
    End If
 
 
If .comprador <> 0 Then
UserList(.comprador).Stats.GLD = UserList(.comprador).Stats.GLD - .oferta
Call SendUserORO(.comprador)
End If
 
Call SendData(ToAll, 0, 0, "||El usuario " & UserList(Userindex).Name & " ha ofertado " & oferta & " monedas de oro." & FONTTYPE_FENIX)
.oferta = oferta
.comprador = Userindex
 
End With
 
UserList(Userindex).Stats.GLD = UserList(Userindex).Stats.GLD - oferta
Call SendUserORO(Userindex)
 
End Sub
 

