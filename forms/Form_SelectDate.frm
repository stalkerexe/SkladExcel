VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_SelectDate 
   Caption         =   "Выбор даты"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3255
   OleObjectBlob   =   "Form_SelectDate.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_SelectDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim mozno As Boolean
Dim dt_2
Dim MyYear
Dim MyMonth
Dim MyDay
Dim MyWeekDay
Dim MyCountDay
Dim l_start

Private Sub ComboBox_Month_Click()
TextBox_Дата.SetFocus
End Sub

Private Sub Image_Вперед_День_Click()
On Error Resume Next
dt_1 = dt_1 + 1
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub


Private Sub Image_Назад_День_Click()
On Error Resume Next
dt_1 = dt_1 - 1
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
Me.StartUpPosition = 0

With iForm
    If iTb = 0 Then
        Me.Top = .Top + .tb_dt1.Top + .tb_dt1.Height * 2
        Me.Left = .Left + .tb_dt1.Left
    Else
        Me.Top = .Top + .tb_dt2.Top + .tb_dt2.Height * 2
        Me.Left = .Left + .tb_dt2.Left
    End If
End With

dt_2 = dt_1

With ComboBox_Month
    .AddItem "Январь"
    .AddItem "Февраль"
    .AddItem "Март"
    .AddItem "Апрель"
    .AddItem "Май"
    .AddItem "Июнь"
    .AddItem "Июль"
    .AddItem "Август"
    .AddItem "Сентябрь"
    .AddItem "Октябрь"
    .AddItem "Ноябрь"
    .AddItem "Декабрь"
End With

Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)

mozno = True
Set_Mоnth (dt_1)
mozno = False


Dim hg
Dim wd

hg = 14
wd = 20

Controls("Cell_" & 1 & "_" & 1).Top = 2
Controls("Cell_" & 1 & "_" & 1).Left = 2

For rw = 1 To 6
    For cm = 1 To 7
        Controls("Cell_" & rw & "_" & cm).Height = hg
        Controls("Cell_" & rw & "_" & cm).Width = wd
        
        Controls("Cell_" & rw & "_" & cm).Font.Size = 9
    Next
Next

For rw = 2 To 6
    Controls("Cell_" & rw & "_" & 1).Top = Controls("Cell_" & rw - 1 & "_" & 1).Top + hg + 1
Next

For cm = 2 To 7
    rw = 1:    Controls("Cell_" & rw & "_" & cm).Top = Controls("Cell_" & rw & "_" & 1).Top
    
    rw = 2:    Controls("Cell_" & rw & "_" & cm).Top = Controls("Cell_" & rw & "_" & 1).Top
    rw = 3:    Controls("Cell_" & rw & "_" & cm).Top = Controls("Cell_" & rw & "_" & 1).Top
    rw = 4:    Controls("Cell_" & rw & "_" & cm).Top = Controls("Cell_" & rw & "_" & 1).Top
    rw = 5:    Controls("Cell_" & rw & "_" & cm).Top = Controls("Cell_" & rw & "_" & 1).Top
    rw = 6:    Controls("Cell_" & rw & "_" & cm).Top = Controls("Cell_" & rw & "_" & 1).Top
Next


For rw = 2 To 6
    Controls("Cell_" & rw & "_" & 1).Left = Controls("Cell_" & 1 & "_" & 1).Left
Next

For cm = 2 To 7
    rw = 1:    Controls("Cell_" & rw & "_" & cm).Left = Controls("Cell_" & rw & "_" & cm - 1).Left + wd + 1
    rw = 2:    Controls("Cell_" & rw & "_" & cm).Left = Controls("Cell_" & rw & "_" & cm - 1).Left + wd + 1
    rw = 3:    Controls("Cell_" & rw & "_" & cm).Left = Controls("Cell_" & rw & "_" & cm - 1).Left + wd + 1
    rw = 4:    Controls("Cell_" & rw & "_" & cm).Left = Controls("Cell_" & rw & "_" & cm - 1).Left + wd + 1
    rw = 5:    Controls("Cell_" & rw & "_" & cm).Left = Controls("Cell_" & rw & "_" & cm - 1).Left + wd + 1
    rw = 6:    Controls("Cell_" & rw & "_" & cm).Left = Controls("Cell_" & rw & "_" & cm - 1).Left + wd + 1
Next

Frame_dt.Height = hg * 6 + 12
Frame_dt.Width = wd * 7 + 14
Frame_dt.Top = Frame_week.Top + Frame_week.Height + 2
Frame_button.Top = Frame_dt.Top + Frame_dt.Height + 2

wd = 22
Frame_week.Left = Frame_dt.Left

Controls("lb_1").Left = Controls("Cell_" & 1 & "_" & 1).Left

    For cm = 1 To 7
        Controls("lb_" & cm).Height = hg
        Controls("lb_" & cm).Width = wd
    Next

For cm = 2 To 7
    Controls("lb_" & cm).Left = Controls("lb_" & cm - 1).Left + wd - 1
Next

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
On Error Resume Next
If CloseMode = 0 Then dt_1 = dt_2
End Sub
Private Sub Set_TextBox_Дата(MyDate As Date)
End Sub
Private Sub Set_TextBox_Year(MyDate As Date)
TextBox_Year.value = Format(MyDate, "yyyy")
End Sub
Private Sub Set_Mоnth(MyDate As Date)
On Error Resume Next
Dim i As Integer
Dim j As Integer
MyYear = VBA.Year(MyDate)
MyMonth = Month(MyDate)
MyDay = Day(MyDate)
Label_Year.Caption = MyYear
ComboBox_Month.ListIndex = MyMonth - 1
MyWeekDay = Weekday(DateSerial(MyYear, MyMonth, 1), vbMonday)
MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)
l_start = 2 - MyWeekDay
For i = 1 To 6
For j = 1 To 7
If l_start >= 1 And l_start <= MyCountDay Then
Me.Controls("Cell_" & i & "_" & j).Caption = l_start
Else
Me.Controls("Cell_" & i & "_" & j).Caption = ""
End If
If l_start = MyDay Then
Set_On_Off CInt(i), CInt(j)
End If
l_start = l_start + 1
Next j, i
End Sub
Private Sub Cmd_Текущий_День_Click()
On Error Resume Next
dt_1 = VBA.Now
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub
Private Sub Cmd_Назад_День_Click()
On Error Resume Next
dt_1 = dt_1 - 1
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub
Private Sub Cmd_Вперед_День_Click()
On Error Resume Next
dt_1 = dt_1 + 1
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub

Private Sub Set_Дата(iRow As Integer, jCol As Integer)
On Error Resume Next
If Me.Controls("Cell_" & iRow & "_" & jCol).Caption = "" Then Exit Sub
MyYear = VBA.Year(dt_1)
MyMonth = Month(dt_1)
MyDay = CInt(Me.Controls("Cell_" & iRow & "_" & jCol).Caption)
dt_1 = DateSerial(MyYear, MyMonth, MyDay)
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False

With iForm
    If iTb = 0 Then .tb_dt1.Text = dt_1
    If iTb = 1 Then .tb_dt2.Text = dt_1
End With

Unload Me
End Sub
Private Sub ComboBox_Month_Change()
On Error Resume Next
If mozno Then Exit Sub
MyYear = VBA.Year(dt_1)
MyMonth = CInt(ComboBox_Month.ListIndex + 1)
MyDay = Day(dt_1)
MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)
If MyDay > MyCountDay Then MyDay = MyCountDay
dt_1 = DateSerial(MyYear, MyMonth, MyDay)
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub
Private Sub SpinButton_Year_SpinDown()
On Error Resume Next
MyYear = VBA.Year(dt_1) - 1
MyMonth = Month(dt_1)
MyDay = Day(dt_1)
MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)
If MyDay > MyCountDay Then MyDay = MyCountDay
dt_1 = DateSerial(MyYear, MyMonth, MyDay)
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub
Private Sub SpinButton_Year_SpinUp()
On Error Resume Next
MyYear = VBA.Year(dt_1) + 1
MyMonth = Month(dt_1)
MyDay = Day(dt_1)
MyCountDay = Day(DateSerial(MyYear, MyMonth + 1, 1) - 1)
If MyDay > MyCountDay Then MyDay = MyCountDay
dt_1 = DateSerial(MyYear, MyMonth, MyDay)
Set_TextBox_Дата (dt_1)
Set_TextBox_Year (dt_1)
mozno = True
Set_Mоnth (dt_1)
mozno = False
End Sub
Private Sub Cell_1_1_Click()
Set_On_Off 1, 1
Set_Дата 1, 1
End Sub
Private Sub Cell_1_2_Click()
Set_On_Off 1, 2
Set_Дата 1, 2
End Sub
Private Sub Cell_1_3_Click()
Set_On_Off 1, 3
Set_Дата 1, 3
End Sub
Private Sub Cell_1_4_Click()
Set_On_Off 1, 4
Set_Дата 1, 4
End Sub
Private Sub Cell_1_5_Click()
Set_On_Off 1, 5
Set_Дата 1, 5
End Sub
Private Sub Cell_1_6_Click()
Set_On_Off 1, 6
Set_Дата 1, 6
End Sub
Private Sub Cell_1_7_Click()
Set_On_Off 1, 7
Set_Дата 1, 7
End Sub
Private Sub Cell_2_1_Click()
Set_On_Off 2, 1
Set_Дата 2, 1
End Sub
Private Sub Cell_2_2_Click()
Set_On_Off 2, 2
Set_Дата 2, 2
End Sub
Private Sub Cell_2_3_Click()
Set_On_Off 2, 3
Set_Дата 2, 3
End Sub
Private Sub Cell_2_4_Click()
Set_On_Off 2, 4
Set_Дата 2, 4
End Sub
Private Sub Cell_2_5_Click()
Set_On_Off 2, 5
Set_Дата 2, 5
End Sub
Private Sub Cell_2_6_Click()
Set_On_Off 2, 6
Set_Дата 2, 6
End Sub
Private Sub Cell_2_7_Click()
Set_On_Off 2, 7
Set_Дата 2, 7
End Sub
Private Sub Cell_3_1_Click()
Set_On_Off 3, 1
Set_Дата 3, 1
End Sub
Private Sub Cell_3_2_Click()
Set_On_Off 3, 2
Set_Дата 3, 2
End Sub
Private Sub Cell_3_3_Click()
Set_On_Off 3, 3
Set_Дата 3, 3
End Sub
Private Sub Cell_3_4_Click()
Set_On_Off 3, 4
Set_Дата 3, 4
End Sub
Private Sub Cell_3_5_Click()
Set_On_Off 3, 5
Set_Дата 3, 5
End Sub
Private Sub Cell_3_6_Click()
Set_On_Off 3, 6
Set_Дата 3, 6
End Sub
Private Sub Cell_3_7_Click()
Set_On_Off 3, 7
Set_Дата 3, 7
End Sub
Private Sub Cell_4_1_Click()
Set_On_Off 4, 1
Set_Дата 4, 1
End Sub
Private Sub Cell_4_2_Click()
Set_On_Off 4, 2
Set_Дата 4, 2
End Sub
Private Sub Cell_4_3_Click()
Set_On_Off 4, 3
Set_Дата 4, 3
End Sub
Private Sub Cell_4_4_Click()
Set_On_Off 4, 4
Set_Дата 4, 4
End Sub
Private Sub Cell_4_5_Click()
Set_On_Off 4, 5
Set_Дата 4, 5
End Sub
Private Sub Cell_4_6_Click()
Set_On_Off 4, 6
Set_Дата 4, 6
End Sub
Private Sub Cell_4_7_Click()
Set_On_Off 4, 7
Set_Дата 4, 7
End Sub
Private Sub Cell_5_1_Click()
Set_On_Off 5, 1
Set_Дата 5, 1
End Sub
Private Sub Cell_5_2_Click()
Set_On_Off 5, 2
Set_Дата 5, 2
End Sub
Private Sub Cell_5_3_Click()
Set_On_Off 5, 3
Set_Дата 5, 3
End Sub
Private Sub Cell_5_4_Click()
Set_On_Off 5, 4
Set_Дата 5, 4
End Sub
Private Sub Cell_5_5_Click()
Set_On_Off 5, 5
Set_Дата 5, 5
End Sub
Private Sub Cell_5_6_Click()
Set_On_Off 5, 6
Set_Дата 5, 6
End Sub
Private Sub Cell_5_7_Click()
Set_On_Off 5, 7
Set_Дата 5, 7
End Sub
Private Sub Cell_6_1_Click()
Set_On_Off 6, 1
Set_Дата 6, 1
End Sub
Private Sub Cell_6_2_Click()
Set_On_Off 6, 2
Set_Дата 6, 2
End Sub
Private Sub Cell_6_3_Click()
Set_On_Off 6, 3
Set_Дата 6, 3
End Sub
Private Sub Cell_6_4_Click()
Set_On_Off 6, 4
Set_Дата 6, 4
End Sub
Private Sub Cell_6_5_Click()
Set_On_Off 6, 5
Set_Дата 6, 5
End Sub
Private Sub Cell_6_6_Click()
Set_On_Off 6, 6
Set_Дата 6, 6
End Sub
Private Sub Cell_6_7_Click()
Set_On_Off 6, 7
Set_Дата 6, 7
End Sub
Private Sub Set_On_Off(iRow As Integer, jCol As Integer)
Dim i As Integer
Dim j As Integer
If Me.Controls("Cell_" & iRow & "_" & jCol).Caption = "" Then Exit Sub
For i = 1 To 6
For j = 1 To 7
Me.Controls("Cell_" & i & "_" & j).BackColor = RGB(255, 255, 255)
Me.Controls("Cell_" & i & "_" & j).BorderColor = RGB(255, 255, 255)
Next j
Next i
Me.Controls("Cell_" & iRow & "_" & jCol).BackColor = RGB(204, 255, 204)
Me.Controls("Cell_" & iRow & "_" & jCol).BorderColor = RGB(150, 150, 150)
End Sub


Private Sub UserForm_Terminate()
On Error Resume Next
Set iForm = Nothing
End Sub

