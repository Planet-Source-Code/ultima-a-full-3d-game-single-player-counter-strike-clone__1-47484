Attribute VB_Name = "GameEngine"
Public mEnemy As New cEnemy
Public PlayerHealth As Single
Public GODI As Integer

Public Function GameInit()
    mEnemy.Init App.Path + "\Ene1.x"
    PlayerHealth = 10
End Function

Public Function RenderGame()
    mEnemy.Render
End Function

Public Function HitPlayer()
    PlayerHealth = PlayerHealth - 1
    If PlayerHealth <= 0 Then
        GODI = 1
    End If
End Function
