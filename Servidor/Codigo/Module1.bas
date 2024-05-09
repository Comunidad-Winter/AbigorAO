Attribute VB_Name = "Module1"
Option Explicit
 
Dim oMail As clsCDOmail
 
Public Function EnviarCorreo(ByVal UserNick As String, ByVal UserMail As String) As Boolean
 
Set oMail = New clsCDOmail
 
    With oMail
        .Servidor = "smtp.gmail.com"
        .Puerto = 465
        .UseAuntentificacion = True
        .SSL = True
        .Usuario = "RECUPERADORDEPASS"
        .PassWord = "estereolove010.-"
        .Asunto = "Recuperación de claves del personaje " & UserNick
        .De = "Maniac-Ao"
        .Para = UserMail
        .Mensaje = "La contraseña de tu personaje es " & ObtenerPassword(UserNick)
        If .Enviar_Backup Then
            EnviarCorreo = True
        Else
            EnviarCorreo = False
        End If
    End With
 
    Set oMail = Nothing
 
End Function
