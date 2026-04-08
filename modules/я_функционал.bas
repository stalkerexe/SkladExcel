Attribute VB_Name = "я_функционал"
Option Explicit

Public Sub A__999()

    With Sheets("Главная")
        .Shapes("cmbt_1").Visible = False
        .Shapes("cmbt_2").Visible = False
        .Shapes("cmbt_3").Visible = False
        .Shapes("cmbt_4").Visible = False
        .Shapes("cmbt_5").Visible = False
        .Shapes("cmbt_6").Visible = False
        .Shapes("cmbt_7").Visible = False
        .Shapes("cmbt_8").Visible = False
        .Shapes("cmbt_9").Visible = False
        .Shapes("cmbt_10").Visible = False
    End With
    
    With Sheets("Склад")
        .Shapes("grCmbBox").Visible = False
    End With
    
    Sheets("Cклад").Visible = 2
    Sheets("Приход").Visible = 2
    Sheets("Отложено_приход").Visible = 2
    Sheets("Расход").Visible = 2
    Sheets("Отложено_расход").Visible = 2

End Sub

