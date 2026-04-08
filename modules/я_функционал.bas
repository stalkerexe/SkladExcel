Attribute VB_Name = "я_функционал"
Option Explicit

Public Sub A__999()
Dim skladSheet As Worksheet

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
    
    Set skladSheet = GetSheetByName(SHEET_SKLAD)
    If skladSheet Is Nothing Then Exit Sub

    With skladSheet
        .Shapes("grCmbBox").Visible = False
        .Visible = 2
    End With
    Sheets("Приход").Visible = 2
    Sheets("Отложено_приход").Visible = 2
    Sheets("Расход").Visible = 2
    Sheets("Отложено_расход").Visible = 2

End Sub

