Attribute VB_Name = "я_фигуры"

Const hh = 2.5
Const iHeight = 54

Public Sub shapes_all()
    Call shapes_height
    Call shapes_leftt
    Call shapes_top
End Sub

Private Sub shapes_height()
    With ActiveSheet
        For i = 1 To 10
            .Shapes("cmbt_" & i).Height = iHeight
        Next
    End With
End Sub


Private Sub shapes_leftt()
    With ActiveSheet
        .Shapes("cmbt_6").Left = .Shapes("cmbt_1").Left + .Shapes("cmbt_1").Width + hh
    End With
    
    With ActiveSheet.Shapes("cmbt_1")
        ActiveSheet.Shapes("cmbt_2").Left = .Left
        ActiveSheet.Shapes("cmbt_3").Left = .Left
        ActiveSheet.Shapes("cmbt_4").Left = .Left
        ActiveSheet.Shapes("cmbt_5").Left = .Left
    End With
    
    With ActiveSheet.Shapes("cmbt_6")
        ActiveSheet.Shapes("cmbt_7").Left = .Left
        ActiveSheet.Shapes("cmbt_8").Left = .Left
        ActiveSheet.Shapes("cmbt_9").Left = .Left
        ActiveSheet.Shapes("cmbt_10").Left = .Left
    End With
End Sub

Private Sub shapes_top()

    With ActiveSheet
    
        .Shapes("cmbt_6").Top = .Shapes("cmbt_1").Top
        
        ind = 2
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        
        ind = 3
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        
        ind = 4
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        
        ind = 5
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        
        
        
        ind = 7
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        ind = 8
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        ind = 9
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        ind = 10
        .Shapes("cmbt_" & ind).Top = .Shapes("cmbt_" & ind - 1).Top + .Shapes("cmbt_" & ind - 1).Height + hh
        
    End With
    
End Sub





Public Sub ertert()
    dddd = ActiveSheet.Shapes("cmbt_4").Height
End Sub
