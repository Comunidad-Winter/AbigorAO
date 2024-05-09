Attribute VB_Name = "Mod_TileEngine"
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

Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, _
ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, _
ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer
 
Public Declare Function vbDABLcolorblend16555 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
 
Public Declare Function vbDABLcolorblend16565 Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
 
Public Declare Function vbDABLcolorblend16555ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long
 
Public Declare Function vbDABLcolorblend16565ck Lib "vbDABL" (ByRef sPtr As Any, ByRef dPtr As Any, ByVal alpha_val%, _
ByVal Width%, ByVal Height%, ByVal sPitch%, ByVal dPitch%, ByVal rVal%, ByVal gVal%, ByVal bVal%) As Long

Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521


Public Const SRCCOPY = &HCC0020









Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type


Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type


Public Type Position
    X As Integer
    Y As Integer
End Type


Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type



Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
    Active As Boolean
    MiniMap_color As Long
End Type


Public Type Grh
    GrhIndex As Integer
    FrameCounter As Byte
    SpeedCounter As Byte
    Started As Byte
End Type


Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type


Public Type HeadData
    Head(1 To 4) As Grh
End Type


Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
End Type


Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type



Public Type FxData
    FX As Grh
    OffsetX As Long
    OffsetY As Long
End Type


Public Type Char
    Aura_Index As Integer
    Aura As Grh
    Active As Byte
    Heading As Byte
    POS As Position

    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Navegando As Byte
    
    Nombre As String
    GM As Integer
    
    haciendoataque As Byte
    Moving As Byte
    MoveOffset As Position
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    
End Type


Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type


Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh

    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type


Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String



Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte


Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long


Public WindowTileWidth As Integer
Public WindowTileHeight As Integer


Public MainViewTop As Integer
Public MainViewLeft As Integer




Public TileBufferSize As Integer


Public DisplayFormhWnd As Long


Public TilePixelHeight As Integer
Public TilePixelWidth As Integer



Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer



Public lastTime As Long



Public MainDestRect   As RECT

Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer





Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh



Public MapData() As MapBlock
Public MapInfo As MapInfo



Public CharList(1 To 10000) As Char




Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long






Public bRain        As Boolean
Public bRainST      As Boolean
Public bTecho       As Boolean
Public brstTick     As Long

Private RLluvia(7)  As RECT
Private iFrameIndex As Byte
Private llTick      As Long
Private LTLluvia(4) As Integer
            

Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum


    Public Enum PlayLoop
        plNone = 0
        plLluviain = 1
        plLluviaout = 2
        plFogata = 3
    End Enum





Sub CargarCabezas()
On Error Resume Next
Dim n As Integer, i As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Numheads


ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #n, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCascos()
On Error Resume Next
Dim n As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCascos


ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #n, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

n = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCuerpos


ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #n, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
Next i

Close #n

End Sub
Sub CargarFxs()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

n = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumFxs


ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #n, , MisFxs(i)
    Call InitGrh(FxData(i).FX, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #n

End Sub
Sub CargarArrayLluvia()
On Error Resume Next
Dim n As Integer, i As Integer
Dim Nu As Integer

n = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Nu


ReDim bLluvia(1 To Nu) As Byte

For i = 1 To Nu
    Get #n, , bLluvia(i)
Next i

Close #n

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)




Dim HWindowX As Integer
Dim HWindowY As Integer

cx = cx - StartPixelLeft
cy = cy - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)


cx = (cx \ TilePixelWidth)
cy = (cy \ TilePixelHeight)

If cx > HWindowX Then
    cx = (cx - HWindowX)
Else
    If cx < HWindowX Then
        cx = (0 - (HWindowX - cx))
    Else
        cx = 0
    End If
End If

If cy > HWindowY Then
    cy = cy + 0 - HWindowY
Else
    If cy < HWindowY Then
        cy = (cy - HWindowY)
    Else
        cy = 0
    End If
End If

tX = UserPos.X + cx
tY = UserPos.Y + cy

End Sub
Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer)
On Error Resume Next


If CharIndex > LastChar Then LastChar = CharIndex

NumChars = NumChars + 1

If arma = 0 Then arma = 2
If escudo = 0 Then escudo = 2
If casco = 0 Then casco = 2

CharList(CharIndex).Head = HeadData(Head)

CharList(CharIndex).Body = BodyData(Body)

If Body > 83 And Body < 88 Then
    CharList(CharIndex).Navegando = 1
Else: CharList(CharIndex).Navegando = 0
End If

CharList(CharIndex).arma = WeaponAnimData(arma)
    
CharList(CharIndex).escudo = ShieldAnimData(escudo)
CharList(CharIndex).casco = CascoAnimData(casco)

CharList(CharIndex).Heading = Heading


CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0


CharList(CharIndex).POS.X = X
CharList(CharIndex).POS.Y = Y


CharList(CharIndex).Active = 1


MapData(X, Y).CharIndex = CharIndex

End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)

CharList(CharIndex).Active = 0
CharList(CharIndex).Criminal = 0
CharList(CharIndex).FX = 0
CharList(CharIndex).FxLoopTimes = 0
CharList(CharIndex).invisible = False
CharList(CharIndex).Moving = 0
CharList(CharIndex).muerto = False
CharList(CharIndex).Nombre = ""
CharList(CharIndex).pie = False
CharList(CharIndex).POS.X = 0
CharList(CharIndex).POS.Y = 0
CharList(CharIndex).UsandoArma = False

End Sub

Sub EraseChar(ByVal CharIndex As Integer)
On Error Resume Next





CharList(CharIndex).Active = 0


If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).CharIndex = 0

Call ResetCharInfo(CharIndex)


NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)



If GrhIndex = 0 Then Exit Sub
Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1







If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed



End Sub

Sub MoveCharByHead(CharIndex As Integer, nheading As Byte)



Dim addX As Integer
Dim addY As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y


Select Case nheading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = X + addX
nY = Y + addY

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading

If UserEstado <> 1 Then Call DoPasosFx(CharIndex)


End Sub


Public Sub DoFogataFx()
If FX = 0 Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then frmMain.StopSound
    Else
        bFogata = HayFogata()
        If bFogata Then frmMain.Play "fuego.wav", True
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim X As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For X = UserPos.X - MinXBorder + 1 To UserPos.X + MinXBorder - 1
            
            If MapData(X, Y).CharIndex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next X
Next Y

EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
Static TickCount(200) As Integer
If Cont = 999 Then Exit Function
TickCount(Cual) = TickCount(Cual) + 1
If TickCount(Cual) < Cont Then
    TickON = False
Else
    TickCount(Cual) = 0
    TickON = True
End If
End Function
Sub DoPasosFx(ByVal CharIndex As Integer)
Static pie As Boolean

If CharList(CharIndex).Navegando = 0 Then
    If UserMontando And EstaPCarea(CharIndex) And CharIndex = UserCharIndex Then
        If TickON(0, 4) Then Call PlayWaveDS(SND_MONTANDO)
    Else
        If CharList(CharIndex).Criminal = 1 Then Exit Sub
        If Not CharList(CharIndex).muerto And EstaPCarea(CharIndex) Then
            CharList(CharIndex).pie = Not CharList(CharIndex).pie
            If CharList(CharIndex).pie Then
                Call PlayWaveDS(SND_PASOS1)
            Else
                Call PlayWaveDS(SND_PASOS2)
            End If
        End If
    End If
Else: Call PlayWaveDS(SND_NAVEGANDO)
End If

End Sub
Sub MoveCharByPosAndHead(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)

On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer



X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y




MapData(nX, nY).CharIndex = CharIndex


CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading




End Sub
Sub MoveCharByPos(CharIndex As Integer, nX As Integer, nY As Integer)
On Error Resume Next

Dim X As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nheading As Byte

X = CharList(CharIndex).POS.X
Y = CharList(CharIndex).POS.Y

MapData(X, Y).CharIndex = 0

addX = nX - X
addY = nY - Y


If Sgn(addX) = 1 Then nheading = EAST
If Sgn(addX) = -1 Then nheading = WEST
If Sgn(addY) = -1 Then nheading = NORTH
If Sgn(addY) = 1 Then nheading = SOUTH

MapData(nX, nY).CharIndex = CharIndex

CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addX)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nheading

End Sub
Sub MoveCharByPosConHeading(CharIndex As Integer, nX As Integer, nY As Integer, nheading As Byte)
On Error Resume Next

If InMapBounds(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y) Then MapData(CharList(CharIndex).POS.X, CharList(CharIndex).POS.Y).CharIndex = 0

MapData(nX, nY).CharIndex = CharIndex

CharList(CharIndex).POS.X = nX
CharList(CharIndex).POS.Y = nY

CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0

CharList(CharIndex).Heading = nheading

End Sub

Sub MoveScreen(Heading As Byte)



Dim X As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer


Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        X = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        X = -1
        
End Select


tX = UserPos.X + X
tY = UserPos.Y + Y


If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    
    AddtoUserPos.X = X
    UserPos.X = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.X - 8 To UserPos.X + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function

Function NextOpenChar() As Integer
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Sub LoadGrhData()
On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer


ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

Get #1, , Grh

Do Until Grh <= 0
   GrhData(Grh).Active = True
    
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If
    
    
    Get #1, , Grh

Loop


Close #1

Dim Count As Long
 
Open IniPath & "minimap.dat" For Binary As #1
    Seek #1, 1
    For Count = 1 To 15000
        If GrhData(Count).Active And (GrhData(Count).pixelHeight = 32 And GrhData(Count).pixelWidth = 32) Then
            Get #1, , GrhData(Count).MiniMap_color
        End If
    Next Count
Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    
    If MapData(X, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    
    If MapData(X, Y).CharIndex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function

Function LegalPosMuerto(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPosMuerto = False
    Exit Function
End If

    
    If MapData(X, Y).Blocked = 1 Then
        LegalPosMuerto = False
        Exit Function
    End If
    
    
    If MapData(X, Y).CharIndex > 0 Then
    If CharList(MapData(X, Y).CharIndex).muerto = True Then
        LegalPosMuerto = False
        Exit Function
    End If
    End If
   
    If Not UserNavegando Then
        If HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    Else
        If Not HayAgua(X, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    End If
    
LegalPosMuerto = True

End Function




Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean





If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(X As Integer, Y As Integer) As Boolean




If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
Surface.BltFast X, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(Surface As DirectDrawSurface7, Grh As Integer, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = X
    .Top = Y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

Surface.GetSurfaceDesc SurfaceDesc


If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    Surface.BltFast destRect.Left, destRect.Top, SurfaceDB.GetBMP(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub



    Sub DDrawTransGrhtoSurface(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)











Dim iGrhIndex As Integer

Dim SourceRect As RECT

Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub


iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)


If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

Surface.BltFast X, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub
Sub DrawBackBufferSurface()

Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)

End Sub
Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function
Sub DrawGrhtoHdc(Hwnd As Long, hDC As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)

If Grh <= 0 Then Exit Sub

SecundaryClipper.SetHWnd Hwnd
SurfaceDB.GetBMP(GrhData(Grh).FileNum).BltToDC hDC, SourceRect, destRect

End Sub
Sub PlayWaveAPI(File As String)
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub
Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Integer, PixelOffsetY As Integer)
On Error Resume Next
If UserCiego Then Exit Sub
Dim Y        As Integer
Dim X        As Integer
Dim minY     As Integer
Dim maxY     As Integer
Dim minX     As Integer
Dim maxX     As Integer
Dim ScreenX  As Integer
Dim ScreenY  As Integer
Dim Moved    As Byte
Dim Grh      As Grh
Dim tempChar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer
Dim iPPy     As Integer
Dim rSourceRect      As RECT
Dim iGrhIndex        As Integer
Dim PixelOffsetXTemp As Integer
Dim PixelOffsetYTemp As Integer
Dim Tiempo As Double
 
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)
 
Tiempo = GetTickCount
ScreenY = 8 + RenderMod.iImageSize
For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 8) - RenderMod.iImageSize
    ScreenX = 8 + RenderMod.iImageSize
    For X = (minX + 8) + RenderMod.iImageSize To (maxX - 8) - RenderMod.iImageSize
        If X > 100 Or Y < 1 Then Exit For
       
        With MapData(X, Y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then .FrameCounter = 1
                    End If
                End If
            End If
 
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With
 
        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight
 
        Call BackBufferSurface.BltFast(((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), rSourceRect, DDBLTFAST_WAIT)
       
        If Not RenderMod.bNoCostas And MapData(X, Y).Graphic(2).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y
 
ScreenY = 8 + RenderMod.iImageSize
For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
    ScreenX = 5 + RenderMod.iImageSize
    For X = (minX + 5) + RenderMod.iImageSize To (maxX - 5) - RenderMod.iImageSize
        If Not (X > 100 Or X < 1) Then
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
   
            Dim tX As Integer
Dim tY As Integer
 
                Call ConvertCPtoTP(frmMain.MainViewShp.Left, frmMain.MainViewShp.Top, frmMain.MouseX, frmMain.MouseY, tX, tY)
   
        '    If MapData(x, y).ObjGrh.GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, y).ObjGrh, iPPx, iPPy, 1, 1)
           
If Abs(tX - X) < 1 And (Abs(tY - Y)) < 1 Then
    If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
        Call SurfaceConColor(BackBufferSurface, MapData(X, Y).ObjGrh, iPPx, iPPy, 1, 1, 250, 250, 250)
     '   Dialogos.DrawText frmMain.MouseX , frmMain.MouseY, MapData(X, Y).NombreItem, vbWhite
    End If
Else
    If MapData(X, Y).ObjGrh.GrhIndex <> 0 Then
        Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).ObjGrh, iPPx, iPPy, 1, 1)
    End If
End If
           
            If MapData(X, Y).CharIndex > 0 Then
                tempChar = CharList(MapData(X, Y).CharIndex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0
   
            If tempChar.MoveOffset.X <> 0 Then
                tempChar.Body.Walk(tempChar.Heading).Started = 1
                tempChar.arma.WeaponWalk(tempChar.Heading).Started = 1
                tempChar.escudo.ShieldWalk(tempChar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + tempChar.MoveOffset.X
                tempChar.MoveOffset.X = tempChar.MoveOffset.X - (8 * Sgn(tempChar.MoveOffset.X))
                Moved = 1
            End If
 
            If tempChar.MoveOffset.Y <> 0 Then
                tempChar.Body.Walk(tempChar.Heading).Started = 1
                tempChar.arma.WeaponWalk(tempChar.Heading).Started = 1
                tempChar.escudo.ShieldWalk(tempChar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + tempChar.MoveOffset.Y
                tempChar.MoveOffset.Y = tempChar.MoveOffset.Y - (8 * Sgn(tempChar.MoveOffset.Y))
                Moved = 1
            End If
 
            If Moved = 0 And tempChar.Moving = 1 Then
                tempChar.Moving = 0
                tempChar.Body.Walk(tempChar.Heading).FrameCounter = 1
                tempChar.Body.Walk(tempChar.Heading).Started = 0
                tempChar.arma.WeaponWalk(tempChar.Heading).FrameCounter = 1
                tempChar.arma.WeaponWalk(tempChar.Heading).Started = 0
                tempChar.escudo.ShieldWalk(tempChar.Heading).FrameCounter = 1
                tempChar.escudo.ShieldWalk(tempChar.Heading).Started = 0
                tempChar.haciendoataque = 0
            End If
           
            If tempChar.haciendoataque = 0 And tempChar.MoveOffset.X = 0 And tempChar.MoveOffset.Y = 0 Then
                tempChar.arma.WeaponWalk(tempChar.Heading).Started = 0
                tempChar.arma.WeaponWalk(tempChar.Heading).FrameCounter = 1
                End If
            If tempChar.haciendoataque = 1 Then
                tempChar.arma.WeaponWalk(tempChar.Heading).Started = 1
                tempChar.haciendoataque = 0
            End If
           
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
               
                If Len(tempChar.Nombre) = 0 Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), iPPx, iPPy, 1, 1)
                    If tempChar.Head.Head(tempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.Head.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                Else
                If tempChar.Navegando = 1 Then
                        Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), iPPx, iPPy, 1, 1)
                        ElseIf (CharList(MapData(X, Y).CharIndex).invisible And UserCharIndex = MapData(X, Y).CharIndex) Then
                        Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), iPPx, iPPy, 1, 1)
                        If tempChar.Head.Head(tempChar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.Head.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                        If tempChar.casco.Head(tempChar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.casco.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                        If tempChar.arma.WeaponWalk(tempChar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.arma.WeaponWalk(tempChar.Heading), iPPx, iPPy, 1, 1)
                       If tempChar.escudo.ShieldWalk(tempChar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.escudo.ShieldWalk(tempChar.Heading), iPPx, iPPy, 1, 1)
 
                    ElseIf Not CharList(MapData(X, Y).CharIndex).invisible And tempChar.Head.Head(tempChar.Heading).GrhIndex > 0 Then
                    If tempChar.Aura_Index > 0 Then
                        Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.Aura, iPPx, iPPy + 40, 1, 1)
                        End If
                   
  If CharList(MapData(X, Y).CharIndex).muerto = True And Not CharList(MapData(X, Y).CharIndex).muerto = False Then
                Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), iPPx, iPPy, 1, 1)
                       If tempChar.Head.Head(tempChar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempChar.Head.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                Else
                 
                        If UserParalizado = True Then 'DDrawTransGrhtoSurface
                    Call SurfaceTonalidad(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), iPPx, iPPy, 1, 1)
                        If tempChar.Head.Head(tempChar.Heading).GrhIndex > 0 Then Call SurfaceTonalidad(BackBufferSurface, tempChar.Head.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                        If tempChar.casco.Head(tempChar.Heading).GrhIndex > 0 Then Call SurfaceTonalidad(BackBufferSurface, tempChar.casco.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                        If tempChar.arma.WeaponWalk(tempChar.Heading).GrhIndex > 0 Then Call SurfaceTonalidad(BackBufferSurface, tempChar.arma.WeaponWalk(tempChar.Heading), iPPx, iPPy, 1, 1)
                        If tempChar.escudo.ShieldWalk(tempChar.Heading).GrhIndex > 0 Then Call SurfaceTonalidad(BackBufferSurface, tempChar.escudo.ShieldWalk(tempChar.Heading), iPPx, iPPy, 1, 1)
                    Else
                    
                    Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.Body.Walk(tempChar.Heading), iPPx, iPPy, 1, 1)
                        If tempChar.Head.Head(tempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.Head.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                        If tempChar.casco.Head(tempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.casco.Head(tempChar.Heading), iPPx + tempChar.Body.HeadOffset.X, iPPy + tempChar.Body.HeadOffset.Y, 1, 0)
                        If tempChar.arma.WeaponWalk(tempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.arma.WeaponWalk(tempChar.Heading), iPPx, iPPy, 1, 1)
                        If tempChar.escudo.ShieldWalk(tempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempChar.escudo.ShieldWalk(tempChar.Heading), iPPx, iPPy, 1, 1)
                    End If
                    End If
                    End If
                    
                    If Nombres Then
                        If (CharList(MapData(X, Y).CharIndex).invisible And UserCharIndex = MapData(X, Y).CharIndex) Then    'matute!!
                            Dim lCenter2 As Long
                            If InStr(tempChar.Nombre, "<") > 0 And InStr(tempChar.Nombre, ">") > 0 Then
                                Dim sClan2 As String
                                lCenter2 = (frmMain.TextWidth(Left$(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1)) / 2) - 16
                                sClan2 = Mid$(tempChar.Nombre, InStr(tempChar.Nombre, "<"))
                                Call Dialogos.DrawText(iPPx - lCenter2, iPPy + 30, Left$(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1), RGB(255, 255, 1))
                                lCenter2 = (frmMain.TextWidth(sClan2) / 2) - 16
                                Call Dialogos.DrawText(iPPx - lCenter2, iPPy + 45, sClan2, RGB(255, 255, 0))
                            Else
                                lCenter2 = (frmMain.TextWidth(tempChar.Nombre) / 2) - 16
                                Call Dialogos.DrawText(iPPx - lCenter2, iPPy + 30, tempChar.Nombre, RGB(255, 255, 0))
                            End If
                     
                        End If
 
                        If Not (tempChar.invisible Or tempChar.Navegando = 1) Then
                       
                            Dim lCenter As Long
                            If InStr(tempChar.Nombre, "<") > 0 And InStr(tempChar.Nombre, ">") > 0 Then
                                Dim sClan As String
                                lCenter = (frmMain.TextWidth(Left$(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1)) / 2) - 16
                                sClan = Mid$(tempChar.Nombre, InStr(tempChar.Nombre, "<"))
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left$(tempChar.Nombre, InStr(tempChar.Nombre, "<") - 1), RGB(RG(tempChar.Criminal, 1), RG(tempChar.Criminal, 2), RG(tempChar.Criminal, 3)))
                                lCenter = (frmMain.TextWidth(sClan) / 2) - 16
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(RG(tempChar.Criminal, 1), RG(tempChar.Criminal, 2), RG(tempChar.Criminal, 3)))
                            Else
                                lCenter = (frmMain.TextWidth(tempChar.Nombre) / 2) - 16
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempChar.Nombre, RGB(RG(tempChar.Criminal, 1), RG(tempChar.Criminal, 2), RG(tempChar.Criminal, 3)))
                            End If
                     
                        End If
                       
                    End If
                End If
   
                If Dialogos.CantidadDialogos > 0 Then Call Dialogos.Update_Dialog_Pos((iPPx + tempChar.Body.HeadOffset.X), (iPPy + tempChar.Body.HeadOffset.Y), MapData(X, Y).CharIndex)
               
                CharList(MapData(X, Y).CharIndex) = tempChar
   
                If CharList(MapData(X, Y).CharIndex).FX <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, FxData(tempChar.FX).FX, iPPx + FxData(tempChar.FX).OffsetX, iPPy + FxData(tempChar.FX).OffsetY, 1, 1, MapData(X, Y).CharIndex)
                'If CharList(MapData(X, Y).CharIndex).FX <> 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, FxData(tempchar.FX).FX, iPPx + FxData(tempchar.FX).OffsetX, iPPy + FxData(tempchar.FX).OffsetY, 1, 1, MapData(X, Y).CharIndex) CLIENTE AB
            End If
 
        End If
        If MapData(X, Y).Graphic(3).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
        ScreenX = ScreenX + 1
    Next X
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y
 
If Not bTecho Then
    ScreenY = 5 + RenderMod.iImageSize
    For Y = (minY + 5) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
        ScreenX = 5 + RenderMod.iImageSize
        For X = (minX + 5) + RenderMod.iImageSize To (maxX - 0) - RenderMod.iImageSize
           
            If X < 101 And X > 0 And Y < 101 And Y > 0 Then
                If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(4), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
            'If MapData(X, Y).Graphic(4).GrhIndex <> 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, MapData(X, Y).Graphic(4), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1) CLIENTE AB
            End If
           
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
End If
 
If bLluvia(UserMap) = 1 Then
    If bRain Or bRainST Then
        If llTick < DirectX.TickCount - 50 Then
            iFrameIndex = iFrameIndex + 1
            If iFrameIndex > 7 Then iFrameIndex = 0
            llTick = DirectX.TickCount
        End If
 
        For Y = 0 To 4
            For X = 0 To 4
                Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(X), SurfaceDB.GetBMP(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
            Next X
        Next Y
    End If
End If
 
 If EsNoche = False Then
Call Dialogos.DrawText(260, 260, "Día", &H80C0FF)
Else
EfectoNoche BackBufferSurface
Call Dialogos.DrawText(260, 260, "Noche", &H808080)
End If

If EsNiebla = True Then
EfectoNiebla 225, 225, 225, BackBufferSurface, 100
Call Dialogos.DrawText(260, 260, "Niebla", &H808080)
End If

End Sub

Public Function RenderSounds()

    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> plLluviain Then
                    Call frmMain.StopSound
                    Call frmMain.Play("lluviain.wav", True)
                    frmMain.IsPlaying = plLluviain
                End If
                
                
            Else
                If frmMain.IsPlaying <> plLluviaout Then
                    Call frmMain.StopSound
                    Call frmMain.Play("lluviaout.wav", True)
                    frmMain.IsPlaying = plLluviaout
                End If
                
                
            End If
        End If
    End If

End Function


Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            CharList(UserCharIndex).POS.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).POS.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).POS.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And CharList(UserCharIndex).POS.Y <= Y
        
End If

End Function



Function PixelPos(X As Integer) As Integer




PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function


Sub LoadGraphics()
        Dim loopc As Integer
        Dim SurfaceDesc As DDSURFACEDESC2
        Dim ddck As DDCOLORKEY
        Dim ddsd As DDSURFACEDESC2
        Dim iLoopUpdate As Integer

        SurfaceDB.MaxEntries = 150
        SurfaceDB.lpDirectDraw7 = DirectDraw
        SurfaceDB.Path = DirGraficos
        Call SurfaceDB.Init(mododinamico)

        
        Call GetBitmapDimensions(DirGraficos & "5556.bmp", ddsd.lWidth, ddsd.lHeight)
              
        RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
        RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
        RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
        RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    
        RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
        RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
        RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
        RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
        AddtoRichTextBox frmCargando.Status, "Hecho.", 255, 150, 50, 1, , False
End Sub



Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean





Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"


UserPos.X = MinXBorder
UserPos.Y = MinYBorder



DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

If Musica = 0 Or FX = 0 Then
    DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
End If



With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmMain.Hwnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0 + 32 * RenderMod.iImageSize
    .Top = 0 + 32 * RenderMod.iImageSize
    .Right = (TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
    .Bottom = (TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
End With
With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If RenderMod.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck
Call InitBlend(BackBufferSurface)

Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

AddtoRichTextBox frmCargando.Status, "Cargando Gráficos....", 255, 150, 50, , , True
Call LoadGraphics

InitTileEngine = True
End Function





Sub ShowNextFrame()












    Static OffsetCounterX As Integer
    Static OffsetCounterY As Integer

    If EngineRun Then
        
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = (OffsetCounterX - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.X)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = 0
                End If
            

            
            
            ElseIf AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (IIf(UserMontando, (32 / 3), 8) * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If

            
            Call RenderScreen(UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            If ModoTrabajo Then Call Dialogos.DrawText(260, 260, "MODO TRABAJO", vbRed)
            
            Call DibujarCartel
            
            Call DrawBackBufferSurface
            
            
            FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1)
Call CrearGrh(534, 2)
End Sub

Function ControlVelocidad(ByVal lastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - lastTime > 20)
End Function

Public Sub DibujarMiniMapa()
   
Dim map_x As Long, map_y As Long
 
    For map_y = 1 To 100
        For map_x = 1 To 100
            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                SetPixel frmMain.Minimap.hDC, map_x, map_y, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color
            End If
        Next map_x
    Next map_y
   
    SetPixel frmMain.Minimap.hDC, UserPos.X, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hDC, UserPos.X + 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hDC, UserPos.X - 1, UserPos.Y, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hDC, UserPos.X, UserPos.Y - 1, RGB(255, 0, 0)
    SetPixel frmMain.Minimap.hDC, UserPos.X, UserPos.Y + 1, RGB(255, 0, 0)
 
    frmMain.Minimap.Refresh
 
End Sub

Sub SurfaceTonalidad(Surface As DirectDrawSurface7, Grh As Grh, ByVal X As Integer _
, ByVal Y As Integer, center As Byte, Animate As Byte, Optional Cantidad As Byte = 150, Optional ByVal KillAnim As Integer = 0)
 
 On Error GoTo errorcito
 
Dim iGrhIndex As Integer
Dim SourceRect As RECT
Dim QuitarAnimacion As Boolean
 
 
If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + 1
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
 
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then 'Matamos la anim del fx )
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
 
                        End If
                    End If
               End If
            End If
        End If
    End If
End If
 
If Grh.GrhIndex = 0 Then Exit Sub
 
iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
 
If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16 'hard coded for speed
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32 'hard coded for speed
    End If
End If
 
With SourceRect
    .Left = GrhData(iGrhIndex).sX + IIf(X < 0, Abs(X), 0)
    .Top = GrhData(iGrhIndex).sY + IIf(Y < 0, Abs(Y), 0)
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With
 
Dim Src As DirectDrawSurface7
Dim rDest As RECT
Dim dArray() As Byte, sArray() As Byte
Dim ddsdSrc As DDSURFACEDESC2, ddsdDest As DDSURFACEDESC2
Dim modo As Long
 
Set Src = SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum)
 
Src.GetSurfaceDesc ddsdSrc
Surface.GetSurfaceDesc ddsdDest
With rDest
    .Left = X
    .Top = Y
    .Right = X + GrhData(iGrhIndex).pixelWidth
    .Bottom = Y + GrhData(iGrhIndex).pixelHeight
   
    If .Right > ddsdDest.lWidth Then
        .Right = ddsdDest.lWidth
    End If
    If .Bottom > ddsdDest.lHeight Then
        .Bottom = ddsdDest.lHeight
    End If
End With
 
Dim SrcLock As Boolean, DstLock As Boolean
SrcLock = False: DstLock = False
 
On Local Error GoTo errorcito
 
Src.Lock SourceRect, ddsdSrc, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
Surface.Lock rDest, ddsdDest, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
 
Surface.GetLockedArray dArray()
Src.GetLockedArray sArray()
       
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
  modo = 555
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
  modo = 565
Else
  MsgBox "Modo de vídeo no esta en 555 o 565 o algo falló."
  End
End If
 
    Call vbDABLcolorblend16565ck(ByVal VarPtr(sArray(SourceRect.Left * 2, SourceRect.Top)), ByVal VarPtr(dArray(X + X, Y)), 125, rDest.Right - rDest.Left, rDest.Bottom - rDest.Top, ddsdSrc.lPitch, ddsdDest.lPitch, 1, 240, 1)
 
 
Surface.Unlock rDest
Src.Unlock SourceRect
Exit Sub
 
errorcito:
Src.Unlock SourceRect
Surface.Unlock rDest
End Sub

Public Sub EfectoNoche(ByRef Surface As DirectDrawSurface7)
If UseAlphaBlending = True Then
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim modo As Long
Dim rRect As RECT
 
Surface.GetSurfaceDesc ddsdDest
 
With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With
 
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
modo = 1
Else
modo = 2
End If
 
Dim DstLock As Boolean
DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True
 
Surface.GetLockedArray dArray()
 
Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), 130, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, 0, 0, 0)
 
HayErrorAlpha:
 
If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If
 End If
End Sub

Public Sub EfectoNiebla(ByVal R As Integer, ByVal g As Integer, ByVal b As Integer, ByRef Surface As DirectDrawSurface7, ByVal a As Integer)
Dim dArray() As Byte, sArray() As Byte
Dim ddsdDest As DDSURFACEDESC2
Dim modo As Long
Dim rRect As RECT
 
Surface.GetSurfaceDesc ddsdDest
 
With rRect
.Left = 0
.Top = 0
.Right = ddsdDest.lWidth
.Bottom = ddsdDest.lHeight
End With
 
If ddsdDest.ddpfPixelFormat.lGBitMask = &H3E0 Then
modo = 0
ElseIf ddsdDest.ddpfPixelFormat.lGBitMask = &H7E0 Then
modo = 1
Else
modo = 2
End If
 
Dim DstLock As Boolean
DstLock = False
 
On Local Error GoTo HayErrorAlpha
 
Surface.Lock rRect, ddsdDest, DDLOCK_WAIT, 0
DstLock = True
 
Surface.GetLockedArray dArray()
 
 
Call vbDABLcolorblend16565ck(ByVal VarPtr(dArray(0, 0)), ByVal VarPtr(dArray(0, 0)), a, rRect.Right - rRect.Left, rRect.Bottom - rRect.Top, ddsdDest.lPitch, ddsdDest.lPitch, R, g, b)
 
HayErrorAlpha:
 
If DstLock = True Then
Surface.Unlock rRect
DstLock = False
End If
 
End Sub
