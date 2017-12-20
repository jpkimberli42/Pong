Sub MoveBall()

Dim HorDir As Integer
Dim VertDir As Integer
Dim rep_count As Integer

Dim paddleA_Top As Integer
Dim paddleA_Bot As Integer
Dim paddleA_Edge As Integer

Dim paddleB_Top As Integer
Dim paddleB_Bot As Integer
Dim paddleB_Edge As Integer

Dim ball_Top As Integer
Dim ball_Bot As Integer
Dim ball_Right As Integer
Dim ball_Left As Integer

Dim table_Top As Integer
Dim table_Bot As Integer
Dim table_Right As Integer
Dim table_Left As Integer

HorDir = 1
VertDir = -1

rep_count = 0

Do
rep_count = rep_count + 1
DoEvents

    With ActiveSheet.Shapes.Range(Array("Oval 1"))
        .IncrementLeft HorDir
        .IncrementTop VertDir
        ball_Left = .Left
        ball_Right = .Left + .Width
        ball_Top = .Top
        ball_Bot = .Top + .Height
    End With
    With ActiveSheet.Shapes.Range(Array("Rectangle 2"))
        table_Left = .Left
        table_Right = .Left + .Width
        table_Top = .Top
        table_Bot = .Top + .Height
    End With
    With ActiveSheet.Shapes.Range(Array("Rectangle 3"))
        paddleA_Edge = .Left + .Width
        paddleA_Top = .Top
        paddleA_Bot = .Top + .Height
    End With
    With ActiveSheet.Shapes.Range(Array("Rectangle 5"))
        paddleB_Edge = .Left
        paddleB_Top = .Top
        paddleB_Bot = .Top + .Height
    End With
    
    
    'ActiveSheet.Shapes.Range(Array("Rectangle 5")).IncrementTop VertDir
    
    If ball_Right >= paddleB_Edge And ball_Bot >= paddleB_Top And ball_Top <= paddleB_Bot Then
        HorDir = -1
    End If
    
    If ball_Left <= paddleA_Edge And ball_Bot >= paddleA_Top And ball_Top <= paddleA_Bot Then
        HorDir = 1
    End If
    
    If ball_Top <= table_Top Then
        VertDir = 1
    End If
    
    If ball_Bot >= table_Bot Then
        VertDir = -1
    End If
    
    If ball_Right >= table_Right Or ball_Left <= table_Left Then
        MsgBox "You Lose"
        Exit Sub
    End If
    
    TimeOut (0.01)
    
Loop Until rep_count = 500

End Sub

Sub TimeOut(duration_ms As Double)
    start_time = Timer
    Do
    DoEvents
    Loop Until (Timer - start_time) >= duration_ms
End Sub
