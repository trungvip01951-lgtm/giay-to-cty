Option Explicit

' Cong no mot khach hang - Trung Kien. Danh sach don tren TongQuan: cot B, dong 13 den 62.
Private Const DONG_DAU As Long = 13
Private Const DONG_CUOI As Long = 62

Public Sub Auto_Open()
    On Error Resume Next
    CapNhatTongQuan
End Sub

' Nut "TAO DON MOI": chep sheet MAU_DON, dat ten theo ngay hom nay, them vao TongQuan
Public Sub TaoDonMoi()
    Dim goc As String, ten As String, i As Long, ws As Worksheet
    goc = ChrW(272) & "H-" & Format(Date, "dd") & "." & Format(Date, "mm") & "." & Format(Date, "yy")
    ten = goc
    i = 1
    Do While CoSheet(ten)
        i = i + 1
        ten = goc & "-" & i
    Loop
    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets("MAU_DON").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ws.Name = ten
    On Error Resume Next
    ws.Tab.Color = RGB(46, 117, 182)
    On Error GoTo 0
    ws.Range("C4").Value = ten
    ws.Range("C5").Value = Date
    CapNhatTongQuan
    Application.ScreenUpdating = True
    ws.Activate
    On Error Resume Next
    ws.Range("C6").Select
    On Error GoTo 0
End Sub

' Nut "CAP NHAT": ghi lai ten tat ca sheet don hang vao cot Ma don cua TongQuan
Public Sub CapNhatTongQuan()
    Dim tq As Worksheet, s As Worksheet, r As Long, ds() As String, n As Long, khac As Boolean
    Set tq = ThisWorkbook.Worksheets("TongQuan")
    ReDim ds(1 To ThisWorkbook.Worksheets.Count)
    For Each s In ThisWorkbook.Worksheets
        If s.Name <> "TongQuan" And s.Name <> "MAU_DON" And s.Name <> "HuongDan" And s.Name <> "BangGia" Then
            n = n + 1
            ds(n) = s.Name
        End If
    Next s
    If n > DONG_CUOI - DONG_DAU + 1 Then
        MsgBox "Qua " & (DONG_CUOI - DONG_DAU + 1) & " don hang: Tong quan chi hien " & (DONG_CUOI - DONG_DAU + 1) & " don dau. Nen luu file sang ban moi.", vbExclamation
        n = DONG_CUOI - DONG_DAU + 1
    End If
    ' chi ghi lai khi danh sach thay doi (tranh hoi luu file khong can thiet)
    For r = DONG_DAU To DONG_CUOI
        If r - DONG_DAU + 1 <= n Then
            If CStr(tq.Cells(r, 2).Value) <> ds(r - DONG_DAU + 1) Then khac = True
        ElseIf CStr(tq.Cells(r, 2).Value) <> "" Then
            khac = True
        End If
    Next r
    If khac Then
        tq.Range("B" & DONG_DAU & ":B" & DONG_CUOI).ClearContents
        For r = 1 To n
            tq.Cells(DONG_DAU + r - 1, 2).Value = ds(r)
        Next r
    End If
    Application.Calculate
End Sub

' Nut "VE TONG QUAN" tren moi sheet don
Public Sub VeTongQuan()
    CapNhatTongQuan
    ThisWorkbook.Worksheets("TongQuan").Activate
    On Error Resume Next
    ThisWorkbook.Worksheets("TongQuan").Range("B13").Select
    On Error GoTo 0
End Sub

Private Function CoSheet(ByVal ten As String) As Boolean
    Dim s As Object
    For Each s In ThisWorkbook.Sheets
        If StrComp(s.Name, ten, vbTextCompare) = 0 Then
            CoSheet = True
            Exit Function
        End If
    Next s
End Function
