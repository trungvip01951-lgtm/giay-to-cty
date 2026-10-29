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
    Dim tq As Worksheet, s As Worksheet, r As Long
    Set tq = ThisWorkbook.Worksheets("TongQuan")
    tq.Range("B" & DONG_DAU & ":B" & DONG_CUOI).ClearContents
    r = DONG_DAU
    For Each s In ThisWorkbook.Worksheets
        If s.Name <> "TongQuan" And s.Name <> "MAU_DON" And s.Name <> "HuongDan" And s.Name <> "BangGia" Then
            If r <= DONG_CUOI Then
                tq.Cells(r, 2).Value = s.Name
                r = r + 1
            End If
        End If
    Next s
    Application.CalculateFull
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
