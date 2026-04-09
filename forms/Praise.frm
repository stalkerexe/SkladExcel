VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Praise 
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   OleObjectBlob   =   "Praise.frm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Praise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim rr As Long
Dim cell As Range
Dim cc(): Dim aa(): Dim find_Flag As Byte
Private Const iCm As Integer = 9

Dim iCmF As Integer
Dim txtF As String

Private Const iSplin As Long = 16



Private Sub tbFind_Nm_Change()
        If find_Flag = 1 Then Exit Sub
        If find_Flag = 3 Then Exit Sub
        
iCmF = 4
        str_ = tbFind_Nm.Text
        Call find_in_text
End Sub

Private Sub tbFind_Cod_Change()
        If find_Flag = 2 Then Exit Sub
        If find_Flag = 3 Then Exit Sub

iCmF = 3
        str_ = tbFind_Cod.Text
        Call find_in_text
End Sub

Private Sub find_in_text()
        On Error Resume Next
        
        ListBox1.Clear
        
        If Len(str_) = 0 Then ListBox1.List = c: Exit Sub
        
        ReDim cc(LBound(c) To UBound(c), 1 To iCm)
        textS = UCase(str_)
        
        iCol = 0
        For i = LBound(c) To UBound(c)
            txtF = c(i, iCmF)
            sCod = c(i, 3)
            
            If Len(str_) = 1 Then
            If sCod <> "---------------------------------" Then
            If UCase(VBA.Left(txtF, 1)) = textS Then
            For cm = 1 To iCm
            cc(i, cm) = c(i, cm)
            Next
            iCol = iCol + 1
            End If
            End If
            Else
            If InStr(1, UCase(txtF), textS) > 0 Then
            If sCod <> "---------------------------------" Then
            For cm = 1 To iCm
            cc(i, cm) = c(i, cm)
            Next
            iCol = iCol + 1
            End If
            End If
            End If
        Next
        
        If iCol = 0 Then Exit Sub
        ReDim w(1 To iCol, 1 To iCm)
        
        j = 1
        For i = LBound(cc) To UBound(cc)
        If cc(i, 1) <> "" Then
        For cm = 1 To iCm
        w(j, cm) = cc(i, cm)
        Next
        j = j + 1
        End If
        Next
        
        ListBox1.List = w
End Sub



Private Sub comb_sk_Change()
        Call load_nomenclature
End Sub

Private Sub load_nomenclature()
        On Error Resume Next
        If comb_sk.ListIndex = -1 Then Exit Sub
        Call Load_from_sk
        tbFind_Nm.SetFocus
End Sub


Private Sub Load_from_sk()
        On Error Resume Next
        ListBox1.Clear
        comb_gr.Clear
        tbFind_Cod = ""
        tbFind_Nm = ""
        sSk = Me.comb_sk.value
        
        Call arr_select_sk
        
        ListBox1.List = c
        
        comb_gr.List = arr_sk_gr_1
        
        ListBox1.AddItem ""
        ListBox1.AddItem ""
        

End Sub


Private Sub comb_sk_set_Change()
On Error Resume Next
With ThisWorkbook.Sheets("my_set")
sSk = comb_sk_set.value
.Range("p2").value = sSk
End With
tbFind_Cod = ""
tbFind_Nm = ""
comb_sk.ListIndex = comb_sk_set.ListIndex
comb_sk_set.Visible = False
End Sub




Private Sub ListBox1_Click()
    comb_sk_set.Visible = False
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        On Error Resume Next
        
        ThisWorkbook.Activate
        
        ind = ListBox1.ListIndex
        If ind = -1 Then Exit Sub
        If ListBox1.List(ind, 2) = "---------------------------------" Then Exit Sub
        If ThisWorkbook.ActiveSheet.Name <> "Приход" And ThisWorkbook.ActiveSheet.Name <> "Расход" Then Sheets("Расход").Select
        
        sSk = Me.comb_sk.value
        With ListBox1
            sID = .List(ind, 0)
            sCod = .List(ind, 2)
            sNm = .List(ind, 3)
            sEd = .List(ind, 4)
            sCnZ = .List(ind, 6)
            sCnR = .List(ind, 7)
            sCol = .List(ind, 5)
        End With
        
        If ThisWorkbook.ActiveSheet.Name = "Приход" Then
            Call ent_Pr
        Else
            Call ent_Zv
        End If
        
        Call remove_green

End Sub

Private Sub ent_Zv()
        On Error Resume Next
        
        iRow = Cells(Rows.Count, zvNm).End(xlUp).Row + 1
        If iRow < rwZv Then iRow = rwZv: GoTo 22
        For Each cell In Range(Cells(rwZv, 1), Cells(iRow, 1))
        If Cells(cell.Row, zvSk) = sSk Then
        If CStr(Cells(cell.Row, zvNm)) = sNm And CStr(Cells(cell.Row, zvCod)) = sCod Then
        rr = cell.Row
        Cells(rr, zvCol) = Cells(rr, zvCol) + 1
        Call add_to_box
        Cells(rr, zvCol).Select
        GoTo 99
        End If
        End If
        Next
22
        Cells(iRow, zvCod).NumberFormat = "@"

        Cells(iRow, 1) = sID
        Cells(iRow, zvSk) = sSk
        Cells(iRow, zvNm) = sNm
        Cells(iRow, zvCod) = sCod
        Cells(iRow, zvEd) = sEd
        
Cells(iRow, zvCn) = sCnR
        Cells(iRow, zvCnR) = sCnR
        Cells(iRow, zvCnZ) = sCnZ
        
        
        Cells(iRow, zvOst) = sCol

        Cells(iRow, zvCol) = 1
        
        '--------------------------------
        'skid
        If Cells(rwZv_mj, zvOst).value = "" Then GoTo 77
        iSkid = Cells(rwZv_mj, zvOst).value
        sCn = Cells(iRow, zvCn) 'цена_без_скидки
        sCnR = sCn - (sCn * iSkid / 100)
        Cells(iRow, zvCnR) = sCnR
77
        '--------------------------------
        Call format_zv
        Call add_to_box
99
End Sub

Private Sub add_to_box()
        On Error Resume Next
        With ThisWorkbook.Sheets("корзина")
            n = .Cells(Rows.Count, zvNm).End(xlUp).Row + 1
            If n < rwZv Then GoTo 22
            For Each cell In Range(.Cells(rwZv, 1), .Cells(n, 1))
            If .Cells(cell.Row, zvSk) = sSk Then
            If CStr(.Cells(cell.Row, zvNm)) = sNm And CStr(.Cells(cell.Row, zvCod)) = sCod Then
            rr = cell.Row
            .Cells(rr, zvCol) = .Cells(rr, zvCol) + 1
            Exit Sub
            End If
            End If
            Next
        End With
22
        With ThisWorkbook.Sheets("корзина")
        n = .Cells(Rows.Count, zvNm).End(xlUp).Row + 1
        If n < rwZv Then n = rwZv
        Range(Cells(iRow, 1), Cells(iRow, 100)).Copy
        .Cells(n, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        End With
        
        With ThisWorkbook.Sheets("корзина")
            .Cells(n, zvCol).NumberFormat = "#,##0.00"
            .Cells(n, zvCnZ).NumberFormat = "#,##0.00"
            .Cells(n, zvCnR).NumberFormat = "#,##0.00"
            .Cells(n, zvSm).NumberFormat = "#,##0.00"
        End With
        
        iRowBox = n
        Call formula_in_box
                
        With ThisWorkbook.Sheets("корзина")
            n = .Cells(Rows.Count, zvNm).End(xlUp).Row
            j = 1
            For i = rwZv To n
            .Cells(i, zvNN) = j
            j = j + 1
            Next
        End With
End Sub

Private Sub format_zv()
        On Error Resume Next
        
        n = Cells(Rows.Count, zvNm).End(xlUp).Row
        Cells(iRow, zvNN) = n - rwZv + 1

        row1 = iRow:  row2 = iRow
        Call format_zv_
        
        Cells(iRow, zvCol).Select
End Sub



Private Sub ent_Pr()
        On Error Resume Next
        
        iRow = Cells(Rows.Count, prNm).End(xlUp).Row + 1
        If iRow < rwZv Then iRow = rwZv: GoTo 22
        For Each cell In Range(Cells(rwZv, 1), Cells(iRow, 1))
        If Cells(cell.Row, prSk) = sSk Then
        If CStr(cell.value) = sID Then
        rr = cell.Row
        Cells(rr, prCol) = Cells(rr, prCol) + 1
        Cells(rr, prCol).Select
        GoTo 99
        End If
        End If
        Next
22
        Cells(iRow, zvCod).NumberFormat = "@"

        Cells(iRow, 1) = sID
        Cells(iRow, prSk) = sSk
        Cells(iRow, prNm) = sNm
        Cells(iRow, prCod) = sCod
        Cells(iRow, prEd) = sEd
        Cells(iRow, prCnR) = sCnR
        Cells(iRow, prCnZ) = sCnZ
        
        Cells(iRow, prCol) = 1
        
        Call ban_input
        Call format_pr
99
End Sub

Private Sub format_pr()
        On Error Resume Next
        n = Cells(Rows.Count, prNm).End(xlUp).Row
        Cells(iRow, prNN) = n - rwZv + 1

        row1 = iRow:  row2 = iRow
        Call format_pr_
        
        Cells(iRow, prCol).Select
End Sub


Private Sub ban_input()
        On Error Resume Next
        With Range(Cells(iRow, prNm), Cells(iRow, prEd)).Validation
            .Delete
            .Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="99999999"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "Запрет редактирования"
            .InputMessage = ""
            .ErrorMessage = "Нельзя изменять данные в этой ячейке!"
            .ShowInput = True
            .showError = True
        End With
End Sub


Private Sub lb_nm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
Me.comb_gr.DropDown
find_Flag = 3
tbFind_Nm.Text = ""
tbFind_Cod.Text = ""
End Sub
Private Sub comb_gr_Change()
On Error Resume Next
ind = comb_gr.ListIndex
If ind = -1 Then Exit Sub
ListBox1.Clear
On Error GoTo 12
row1 = comb_gr.List(ind, 0) + 1
row2 = comb_gr.List(ind + 1, 0) - 1
If row2 < row1 Then Exit Sub
12
If ind = comb_gr.ListCount - 1 Then row2 = UBound(c)
ReDim aa(1 To row2 - row1 + 1, 1 To iCm)
j = 1
For i = row1 To row2
For cm = 1 To iCm
aa(j, cm) = c(i, cm)
Next
j = j + 1
Next
ListBox1.List = aa
Me.Caption = "      " & comb_gr.List(ind, 1)
End Sub




Private Sub UserForm_Initialize()
    Call load_all
End Sub

Private Sub UserForm_Click()
    On Error Resume Next
    comb_sk_set.Visible = False
End Sub

Private Sub load_all()
    On Error Resume Next
    Call load_sklads
    
    Me.comb_gr.ListRows = 25
    With ThisWorkbook.Sheets("my_set")
        sSk = .Range("p2").value
        comb_sk_set.value = sSk
        
        If comb_sk.ListCount = 0 Then GoTo 11
        If comb_sk.value = "" Then comb_sk.ListIndex = 0
    End With
11
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 15
    Me.Left = Application.Width - Me.Width - 15
    Me.Height = Application.Height - Me.Top - 20
    Me.ListBox1.Height = Me.Height - 60
    
    With ThisWorkbook.Sheets("setting")
    If .Range("b6").value = 0 Then
    lb_cod.Width = 0
    tbFind_Cod.Width = 0
    tbFind_Nm.Left = lb_nm.Left
    tbFind_Nm.Width = lb_nm.Width + 13
    ListBox1.ColumnWidths = "0;0;0;300;0;40;0;0;0"
    Else
    ListBox1.ColumnWidths = "0;0;60;250;0;40;0;0;0"
    End If
    End With
    
    With comb_gr
    .Left = lb_nm.Left
    .Top = lb_nm.Top
    .Width = lb_nm.Width + 13
    .ZOrder 1
    End With
       
       
    NO.ForeColor = RGB(255, 255, 255)
    
End Sub

Private Sub load_sklads()
        Call LoadSkToControl(comb_sk)
        Call LoadSkToControl(comb_sk_set)

End Sub


        
Private Sub NO_Click()
Unload Me
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
NO.SpecialEffect = 0
NO.Top = 18
NO.Height = 15.75
End Sub
Private Sub lb_ost_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
NO.SpecialEffect = 0
NO.Top = 18
NO.Height = 15.75
End Sub
Private Sub tbFind_Nm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
NO.SpecialEffect = 0
NO.Top = 18
NO.Height = 15.75
End Sub
Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
NO.SpecialEffect = 0
NO.Top = 18
NO.Height = 15.75
End Sub
Private Sub NO_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
NO.SpecialEffect = 6
NO.Top = 16
NO.Height = 20
End Sub
Private Sub comb_sk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
comb_sk.DropDown
End Sub
Private Sub comb_sk_set_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
comb_sk_set.DropDown
End Sub


Private Sub tbFind_Cod_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
find_Flag = 1
tbFind_Nm.Text = ""
End Sub
Private Sub tbFind_Nm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
find_Flag = 2
tbFind_Cod.Text = ""
End Sub

Private Sub lb_cod_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
comb_gr.DropDown
find_Flag = 3
tbFind_Nm.Text = ""
tbFind_Cod.Text = ""
End Sub

Private Sub CommandButton1_Click()
comb_sk_set.Visible = True
End Sub
Private Sub Label9_Click()
comb_sk.DropDown
End Sub

Private Sub SpinButton1_SpinDown()
        On Error Resume Next
        ListBox1.TopIndex = ListBox1.TopIndex + iSplin
End Sub

Private Sub SpinButton1_SpinUp()
        On Error Resume Next
        If ListBox1.TopIndex <= iSplin Then ListBox1.TopIndex = 0: Exit Sub
        ListBox1.TopIndex = ListBox1.TopIndex - iSplin
End Sub

Private Sub UserForm_Terminate()
On Error Resume Next
Erase aa
Erase c
Erase cc
Erase gr
Erase nm
Erase cod
Erase ed
Erase ost
Erase cnZ
Erase cnR
End Sub
