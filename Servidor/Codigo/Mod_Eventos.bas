Attribute VB_Name = "Mod_Eventos"
Option Explicit
Dim i As Integer
Public Sub ResetUser(UserIndex As Integer)
For i = 1 To 14
    If UserIndex = SistemaEventos.UsuarioA(i) Then
        UserList(UserIndex).flags.entorneo.Contrincante = 0
        UserList(UserIndex).flags.entorneo.Esperando = False
        UserList(UserIndex).flags.entorneo.entorneo = False
        UserList(UserIndex).flags.entorneo.enDuelo = False
        Call WarpUserChar(UserIndex, 1, 66, 73, True)
        SistemaEventos.UsuarioA(i) = 0
    End If
Next
End Sub
Public Sub ResetUserSocket(UserIndex As Integer)
For i = 1 To 14
    If UserIndex = SistemaEventos.UsuarioA(i) Then
        UserList(UserIndex).flags.entorneo.Contrincante = 0
        UserList(UserIndex).flags.entorneo.Esperando = False
        UserList(UserIndex).flags.entorneo.entorneo = False
        UserList(UserIndex).flags.entorneo.enDuelo = False
        Call WarpUserChar(UserIndex, 1, 66, 73, True)
        SistemaEventos.UsuarioA(i) = 9099
    End If
Next
End Sub
'Inicializacion de Modo de Torneo Automatico
Public Sub ModoTorneo(Tipo As Integer)
    Select Case Tipo
        Case "1"
            SistemaEventos.Modo = 1
            SistemaEventos.Contadores = 0
            SistemaEventos.Inicializacion = True
            SistemaEventos.General = False
            Call SendData(ToAll, 1, 0, "||Torneo 1 vs 1 activado. Cupo: 8" & FONTTYPE_FENIX)
            Call SendData(ToAll, 1, 0, "||Para ingresar, tipeea /ACCEDER" & FONTTYPE_INFO)
            Exit Sub
        Case "2"
            SistemaEventos.Contadores = 0
            SistemaEventos.Inicializacion = True
            SistemaEventos.General = False
            Call SendData(ToAll, 1, 0, "||Torneo 2 vs 2 activado. Cupo: 8 Parejas" & FONTTYPE_FENIX)
            Call SendData(ToAll, 1, 0, "||Para ingresar, tipeea /ACCEDER" & FONTTYPE_INFO)
            Exit Sub
        Case "3"
            SistemaEventos.Contadores = 0
            SistemaEventos.Inicializacion = True
            SistemaEventos.General = False
            Call SendData(ToAll, 1, 0, "||Torneo 3 vs 3 activado. Cupo: 8 Parejas" & FONTTYPE_FENIX)
            Call SendData(ToAll, 1, 0, "||Para ingresar, tipeea /ACCEDER" & FONTTYPE_INFO)
            Exit Sub
        Case "4"
            SistemaEventos.Contadores = 0
            SistemaEventos.Inicializacion = True
            SistemaEventos.General = False
            Call SendData(ToAll, 1, 0, "||Llego la hora de una DEATHMATCH" & FONTTYPE_FENIX)
            Call SendData(ToAll, 1, 0, "||Para ingresar, tipeea /ACCEDER" & FONTTYPE_INFO)
            Exit Sub
    End Select
End Sub

Public Sub InicioBatalla()

'Sistema de Eventos. Inicializacion Modo 1 vs 1
If SistemaEventos.Modo = 1 Then
    Call SendData(ToAll, SistemaEventos.UsuarioA(1), 0, "||1 vs 1> Se enfrentan " & UserList(SistemaEventos.UsuarioA(1)).Name & " contra " & UserList(SistemaEventos.UsuarioA(2)).Name & FONTTYPE_INFO)
    Call WarpUserChar(SistemaEventos.UsuarioA(1), 190, 66, 72, True)
    UserList(SistemaEventos.UsuarioA(1)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(2)
    UserList(SistemaEventos.UsuarioA(1)).flags.entorneo.Esperando = True
    UserList(SistemaEventos.UsuarioA(1)).flags.entorneo.enDuelo = True
    Call WarpUserChar(SistemaEventos.UsuarioA(2), 190, 84, 83, True)
    UserList(SistemaEventos.UsuarioA(2)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(1)
    UserList(SistemaEventos.UsuarioA(2)).flags.entorneo.Esperando = True
    UserList(SistemaEventos.UsuarioA(2)).flags.entorneo.enDuelo = True
    SistemaEventos.Indice = 2
    Exit Sub
End If
End Sub

Public Sub DueloSimple(Ganador As Integer)
If SistemaEventos.Modo = 1 Then
'Octavos
Select Case SistemaEventos.Indice
    Case 2
        SistemaEventos.UsuarioA(9) = Ganador
        If Ganador <> 9099 Then
            Call WarpUserChar(Ganador, 1, 25, 47, True)
            UserList(Ganador).flags.entorneo.enDuelo = False
        End If
        'Next Fight in Octavos
        Call SendData(ToAll, 1, 0, "||1 vs 1> Siguiente Batalla ... " & FONTTYPE_VENENO)
        '####### SEGURIDAD DE DESLOGEO DE USUARIO ##########
        If SistemaEventos.UsuarioA(3) = 9099 And SistemaEventos.UsuarioA(4) = 9099 Then
            SistemaEventos.Indice = 4
            Call DueloSimple(9099)
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(3) = 9099 Then
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a cuartos " & UserList(SistemaEventos.UsuarioA(4)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 4
            Call DueloSimple(SistemaEventos.UsuarioA(4))
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(4) = 9099 Then
            'Llamo al siguiente duelo de cuartos
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a cuartos " & UserList(SistemaEventos.UsuarioA(3)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 4
            Call DueloSimple(SistemaEventos.UsuarioA(3))
            Exit Sub
        End If
        '################ DUELO 1 vs 1 entre 3 y 4 ####################
        Call SendData(ToAll, 1, 0, "||1 vs 1> Se enfrentan " & UserList(SistemaEventos.UsuarioA(3)).Name & " contra " & UserList(SistemaEventos.UsuarioA(4)).Name & FONTTYPE_INFO)
        Call WarpUserChar(SistemaEventos.UsuarioA(3), 190, 66, 72, True)
        UserList(SistemaEventos.UsuarioA(3)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(4)
        UserList(SistemaEventos.UsuarioA(3)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(3)).flags.entorneo.enDuelo = True
        'Asignacion para el 2º Usuario
        Call WarpUserChar(SistemaEventos.UsuarioA(4), 190, 84, 83, True)
        UserList(SistemaEventos.UsuarioA(4)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(3)
        UserList(SistemaEventos.UsuarioA(4)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(4)).flags.entorneo.enDuelo = True
        SistemaEventos.Indice = 4
        Exit Sub
    Case 4
        SistemaEventos.UsuarioA(10) = Ganador
        If Ganador <> 9099 Then
            Call WarpUserChar(Ganador, 1, 25, 47, True)
            UserList(Ganador).flags.entorneo.enDuelo = False
        End If
        'Next Fight in Octavos
        Call SendData(ToAll, 1, 0, "||1 vs 1> Siguiente Batalla ... " & FONTTYPE_VENENO)
        '####### SEGURIDAD DE DESLOGEO DE USUARIO ##########
        If SistemaEventos.UsuarioA(5) = 9099 And SistemaEventos.UsuarioA(6) = 9099 Then
            SistemaEventos.Indice = 6
            Call DueloSimple(9099)
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(5) = 9099 Then
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a cuartos " & UserList(SistemaEventos.UsuarioA(6)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 6
            Call DueloSimple(SistemaEventos.UsuarioA(6))
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(6) = 9099 Then
            'Llamo al siguiente duelo de cuartos
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a cuartos " & UserList(SistemaEventos.UsuarioA(5)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 6
            Call DueloSimple(SistemaEventos.UsuarioA(5))
            Exit Sub
        End If
        '################ DUELO 1 vs 1 entre 5 y 6 ####################
        Call SendData(ToAll, 1, 0, "||1 vs 1> Se enfrentan " & UserList(SistemaEventos.UsuarioA(5)).Name & " contra " & UserList(SistemaEventos.UsuarioA(6)).Name & FONTTYPE_INFO)
        Call WarpUserChar(SistemaEventos.UsuarioA(5), 190, 66, 72, True)
        UserList(SistemaEventos.UsuarioA(5)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(6)
        UserList(SistemaEventos.UsuarioA(5)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(5)).flags.entorneo.enDuelo = True
        'Asignacion para el 2º Usuario
        Call WarpUserChar(SistemaEventos.UsuarioA(6), 190, 84, 83, True)
        UserList(SistemaEventos.UsuarioA(6)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(5)
        UserList(SistemaEventos.UsuarioA(6)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(6)).flags.entorneo.enDuelo = True
        SistemaEventos.Indice = 6
        Exit Sub
    Case 6
        SistemaEventos.UsuarioA(11) = Ganador
        If Ganador <> 9099 Then
            Call WarpUserChar(Ganador, 1, 25, 47, True)
            UserList(Ganador).flags.entorneo.enDuelo = False
        End If
        'Next Fight in Octavos
        Call SendData(ToAll, 1, 0, "||1 vs 1> Siguiente Batalla ... " & FONTTYPE_VENENO)
        '####### SEGURIDAD DE DESLOGEO DE USUARIO ##########
        If SistemaEventos.UsuarioA(7) = 9099 And SistemaEventos.UsuarioA(8) = 9099 Then
            SistemaEventos.Indice = 8
            Call DueloSimple(9099)
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(7) = 9099 Then
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a cuartos " & UserList(SistemaEventos.UsuarioA(8)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 8
            Call DueloSimple(SistemaEventos.UsuarioA(8))
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(8) = 9099 Then
            'Llamo al siguiente duelo de cuartos
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a cuartos " & UserList(SistemaEventos.UsuarioA(7)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 8
            Call DueloSimple(SistemaEventos.UsuarioA(7))
            Exit Sub
        End If
        '################ DUELO 1 vs 1 entre 7 y 8 ####################
        Call SendData(ToAll, 1, 0, "||1 vs 1> Se enfrentan " & UserList(SistemaEventos.UsuarioA(7)).Name & " contra " & UserList(SistemaEventos.UsuarioA(8)).Name & FONTTYPE_INFO)
        Call WarpUserChar(SistemaEventos.UsuarioA(7), 190, 66, 72, True)
        UserList(SistemaEventos.UsuarioA(7)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(8)
        UserList(SistemaEventos.UsuarioA(7)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(7)).flags.entorneo.enDuelo = True
        'Asignacion para el 2º Usuario
        Call WarpUserChar(SistemaEventos.UsuarioA(8), 190, 84, 83, True)
        UserList(SistemaEventos.UsuarioA(8)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(7)
        UserList(SistemaEventos.UsuarioA(8)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(8)).flags.entorneo.enDuelo = True
        SistemaEventos.Indice = 8
        Exit Sub
    Case 8
        SistemaEventos.UsuarioA(12) = Ganador
        If Ganador <> 9099 Then
            Call WarpUserChar(Ganador, 1, 25, 47, True)
            UserList(Ganador).flags.entorneo.enDuelo = False
        End If
        'Next Fight in Octavos
        Call SendData(ToAll, 1, 0, "||1 vs 1> Comenzamos con Cuartos de Final " & FONTTYPE_VENENO)
        '####### SEGURIDAD DE DESLOGEO DE USUARIO ##########
        If SistemaEventos.UsuarioA(9) = 9099 And SistemaEventos.UsuarioA(10) = 9099 Then
            SistemaEventos.Indice = 10
            Call DueloSimple(9099)
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(9) = 9099 Then
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a la Final " & UserList(SistemaEventos.UsuarioA(10)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 10
            Call DueloSimple(SistemaEventos.UsuarioA(10))
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(10) = 9099 Then
            'Llamo al siguiente duelo de cuartos
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a la Final " & UserList(SistemaEventos.UsuarioA(9)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 10
            Call DueloSimple(SistemaEventos.UsuarioA(9))
            Exit Sub
        End If
        '################ DUELO 1 vs 1 entre 9 y 10 ####################
        Call SendData(ToAll, 1, 0, "||1 vs 1> Se enfrentan " & UserList(SistemaEventos.UsuarioA(9)).Name & " contra " & UserList(SistemaEventos.UsuarioA(10)).Name & FONTTYPE_INFO)
        Call WarpUserChar(SistemaEventos.UsuarioA(9), 190, 66, 72, True)
        UserList(SistemaEventos.UsuarioA(9)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(10)
        UserList(SistemaEventos.UsuarioA(9)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(9)).flags.entorneo.enDuelo = True
        'Asignacion para el 2º Usuario
        Call WarpUserChar(SistemaEventos.UsuarioA(10), 190, 84, 83, True)
        UserList(SistemaEventos.UsuarioA(10)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(9)
        UserList(SistemaEventos.UsuarioA(10)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(10)).flags.entorneo.enDuelo = True
        SistemaEventos.Indice = 10
        Exit Sub
    Case 10
        SistemaEventos.UsuarioA(13) = Ganador
        If Ganador <> 9099 Then
            Call WarpUserChar(Ganador, 1, 25, 47, True)
            UserList(Ganador).flags.entorneo.enDuelo = False
        End If
        'Next Fight in Octavos
        Call SendData(ToAll, 1, 0, "||1 vs 1> Siguiente Batalla " & FONTTYPE_VENENO)
        '####### SEGURIDAD DE DESLOGEO DE USUARIO ##########
        If SistemaEventos.UsuarioA(11) = 9099 And SistemaEventos.UsuarioA(12) = 9099 Then
            SistemaEventos.Indice = 12
            Call DueloSimple(9099)
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(11) = 9099 Then
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a la Final " & UserList(SistemaEventos.UsuarioA(12)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 12
            Call DueloSimple(SistemaEventos.UsuarioA(12))
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(12) = 9099 Then
            'Llamo al siguiente duelo de cuartos
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a la Final " & UserList(SistemaEventos.UsuarioA(11)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 12
            Call DueloSimple(SistemaEventos.UsuarioA(10))
            Exit Sub
        End If
        '################ DUELO 1 vs 1 entre 11 y 12 ####################
        Call SendData(ToAll, 1, 0, "||1 vs 1> Se enfrentan " & UserList(SistemaEventos.UsuarioA(11)).Name & " contra " & UserList(SistemaEventos.UsuarioA(12)).Name & FONTTYPE_INFO)
        Call WarpUserChar(SistemaEventos.UsuarioA(11), 190, 66, 72, True)
        UserList(SistemaEventos.UsuarioA(11)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(12)
        UserList(SistemaEventos.UsuarioA(11)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(11)).flags.entorneo.enDuelo = True
        'Asignacion para el 2º Usuario
        Call WarpUserChar(SistemaEventos.UsuarioA(12), 190, 84, 83, True)
        UserList(SistemaEventos.UsuarioA(12)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(11)
        UserList(SistemaEventos.UsuarioA(12)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(12)).flags.entorneo.enDuelo = True
        SistemaEventos.Indice = 12
        Exit Sub
    Case 12
        SistemaEventos.UsuarioA(14) = Ganador
        If Ganador <> 9099 Then
            Call WarpUserChar(Ganador, 1, 25, 47, True)
            UserList(Ganador).flags.entorneo.enDuelo = False
        End If
        'Next Fight in Octavos
        Call SendData(ToAll, 1, 0, "||1 vs 1> Final del Torneo 1 vs 1 Powered by Localstrike " & FONTTYPE_VENENO)
        '####### SEGURIDAD DE DESLOGEO DE USUARIO ##########
        If SistemaEventos.UsuarioA(13) = 9099 And SistemaEventos.UsuarioA(14) = 9099 Then
            SistemaEventos.Indice = 14
            Call DueloSimple(9099)
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(13) = 9099 Then
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a la Final " & UserList(SistemaEventos.UsuarioA(14)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 14
            Call DueloSimple(SistemaEventos.UsuarioA(14))
            Exit Sub
        End If
        If SistemaEventos.UsuarioA(14) = 9099 Then
            'Llamo al siguiente duelo de cuartos
            Call SendData(ToAll, 1, 0, "||1 vs 1> Automaticamente pasa a la Final " & UserList(SistemaEventos.UsuarioA(13)).Name & FONTTYPE_INFO)
            SistemaEventos.Indice = 14
            Call DueloSimple(SistemaEventos.UsuarioA(13))
            Exit Sub
        End If
        '################ DUELO 1 vs 1 entre 13 y 14 ####################
        Call SendData(ToAll, 1, 0, "||1 vs 1> Se enfrentan " & UserList(SistemaEventos.UsuarioA(13)).Name & " contra " & UserList(SistemaEventos.UsuarioA(14)).Name & FONTTYPE_INFO)
        Call WarpUserChar(SistemaEventos.UsuarioA(13), 190, 66, 72, True)
        UserList(SistemaEventos.UsuarioA(13)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(14)
        UserList(SistemaEventos.UsuarioA(13)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(13)).flags.entorneo.enDuelo = True
        'Asignacion para el 2º Usuario
        Call WarpUserChar(SistemaEventos.UsuarioA(14), 190, 84, 83, True)
        UserList(SistemaEventos.UsuarioA(14)).flags.entorneo.Contrincante = SistemaEventos.UsuarioA(13)
        UserList(SistemaEventos.UsuarioA(14)).flags.entorneo.Esperando = True
        UserList(SistemaEventos.UsuarioA(14)).flags.entorneo.enDuelo = True
        SistemaEventos.Indice = 14
        Exit Sub
    Case 14
        If Ganador = 9099 Then
            Call SendData(ToAll, 1, 0, "||1 vs 1> No hay nadie en la Final ya que los finalistan no se encuentran INGAME. Cerramos el Torneo" & FONTTYPE_INFO)
            SistemaEventos.Contadores = 0
            SistemaEventos.Inicializacion = False
            SistemaEventos.Finalizacion = False
            SistemaEventos.General = False
            Exit Sub
        End If
        Call SendData(ToAll, 1, 0, "||1 vs 1> El gran ganador de este TORNEO es: " & UserList(Ganador).Name & FONTTYPE_INFO)
        Call ResetUser(Ganador)
        'ENTREGA DE PREMIOS
        SistemaEventos.Contadores = 0
        SistemaEventos.Inicializacion = False
        SistemaEventos.Finalizacion = False
        SistemaEventos.General = False
        Exit Sub
    End Select
End If
End Sub
