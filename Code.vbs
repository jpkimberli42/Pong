Sub MoveBall()

Dim HorDir As Integer
Dim VertDir As Integer

HorDir = 1
VertDir = -1

rep_count = 0

Do
rep_count = rep_count + 1
DoEvents

    ActiveSheet.Shapes.Range(Array("Oval 1")).IncrementLeft HorDir
    ActiveSheet.Shapes.Range(Array("Oval 1")).IncrementTop VertDir
    
    'ActiveSheet.Shapes.Range(Array("Rectangle 5")).IncrementTop VertDir
    
    If ActiveSheet.Shapes.Range(Array("Oval 1")).Left + ActiveSheet.Shapes.Range(Array("Oval 1")).Width >= ActiveSheet.Shapes.Range(Array("Rectangle 5")).Left Then
        HorDir = -1
    End If
    
    If ActiveSheet.Shapes.Range(Array("Oval 1")).Left <= ActiveSheet.Shapes.Range(Array("Rectangle 3")).Left + ActiveSheet.Shapes.Range(Array("Rectangle 3")).Width Then
        HorDir = 1
    End If
    
    If ActiveSheet.Shapes.Range(Array("Oval 1")).Top <= ActiveSheet.Shapes.Range(Array("Rectangle 2")).Top Then
        VertDir = 1
    End If
    
    If ActiveSheet.Shapes.Range(Array("Oval 1")).Top + ActiveSheet.Shapes.Range(Array("Oval 1")).Height >= ActiveSheet.Shapes.Range(Array("Rectangle 2")).Height + ActiveSheet.Shapes.Range(Array("Rectangle 2")).Top Then
        VertDir = -1
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
