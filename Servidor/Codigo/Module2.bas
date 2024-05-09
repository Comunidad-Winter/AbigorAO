Attribute VB_Name = "Module2"
Option Explicit
'sistema de subastas =D
Type S
HaySubastas As Boolean
Vendedor As Integer
comprador As Integer
oferta As Integer
ItemEnVenta As Integer
CantidadVenta As Integer
End Type
Public NPCSubastas As Byte
Public Subastas As S
 
Sub TerminarSubasta()
 
Dim ob As Obj
 
With Subastas
 
    ob.Amount = .CantidadVenta
    ob.OBJIndex = .ItemEnVenta
   
If .comprador = 0 Then
    Call SendData(ToIndex, .Vendedor, 0, "|| Nadie respondio a tu subasta" & FONTTYPE_INFO)
        If Not MeterItemEnInventario(.Vendedor, ob) Then Call TirarItemAlPiso(UserList(.Vendedor).POS, ob)
            UserList(.Vendedor).Stats.GLD = UserList(.Vendedor).Stats.GLD + .oferta
        Call SendUserORO(.Vendedor)
    Exit Sub
End If
 
If Not MeterItemEnInventario(.comprador, ob) Then Call TirarItemAlPiso(UserList(.comprador).POS, ob)
 
Call SendData(ToAll, 0, 0, "|| La subasta termino, el ganador fue " & UserList(.comprador).Name & " a un precio de " & _
.oferta & "." & FONTTYPE_INFO)
 
.HaySubastas = False
.Vendedor = 0
.comprador = 0
.oferta = 0
.ItemEnVenta = 0
.CantidadVenta = 0
 
End With
 
End Sub
 
Sub Subastar(UserIndex As Integer, Precioinicial As Long)
 
Dim npc As Integer
 
npc = UserList(UserIndex).flags.TargetNpc
 
If Npclist(npc).NPCtype <> NPCSubastas Then
Exit Sub
End If
 
Dim Itemsubastar As Integer
Dim CantidadItemSubastar As Integer
 
Itemsubastar = MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).OBJInfo.OBJIndex
CantidadItemSubastar = MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).OBJInfo.Amount
 
If Distancia(Npclist(npc).POS, UserList(UserIndex).POS) < 1 Then
Call SendData(ToIndex, UserIndex, 0, "|| Estas Muy Lejos" & FONTTYPE_INFO)
Exit Sub
End If
 
If Itemsubastar = 0 Then
SendData ToIndex, UserIndex, 0, "|| Tira el item q deseas subastar" & FONTTYPE_INFO
Exit Sub
End If
 
'ObjData (MapData(UserList(Userindex).POS.Map, UserList(Userindex).POS.X, UserList(Userindex).POS.Y).OBJInfo.OBJIndex)
 
Subastas.Vendedor = UserIndex
Subastas.ItemEnVenta = Itemsubastar
Subastas.CantidadVenta = CantidadItemSubastar
Subastas.oferta = Precioinicial
 
SendData ToAll, 0, 0, "|| el usuario " & UserList(UserIndex).Name & " esta subasttando " & _
ObjData(MapData(UserList(UserIndex).POS.Map, UserList(UserIndex).POS.X, UserList(UserIndex).POS.Y).OBJInfo.OBJIndex).Name & _
" a un precio inicial de " & Precioinicial & "." & FONTTYPE_INFO
 
Subastas.HaySubastas = True
 
End Sub

