' ============================================================
' קנדל BI — כלי בדיקת דוחות תשומות v1.0
' kandelbi.org
' ============================================================
' הוראות התקנה:
' 1. פתח אקסל חדש ושמור כ-.xlsm (אקסל עם מאקרואים)
' 2. פתח עורך VBA: Alt+F11
' 3. לחץ Insert → Module
' 4. הדבק את כל הקוד הזה
' 5. שנה את FILE_ID, DOWNLOAD_DATE ו-USER_EMAIL לערכים שקיבלת
' 6. שמור וסגור את עורך ה-VBA
' ============================================================

' ===== פרטי רישיון — למלא בעת ההורדה =====
Const FILE_ID As String = "REPLACE_FILE_ID"
Const DOWNLOAD_DATE As String = "REPLACE_DATE"  ' פורמט: YYYY-MM-DD
Const USER_EMAIL As String = "REPLACE_EMAIL"
Const LICENSE_URL As String = "https://license-check.orkandel11.workers.dev/"
' ============================================

' ===== בדיקת תוקף רישיון =====
Function CheckLicense() As Boolean
    Dim downloaded As Date
    Dim expiry As Date
    
    On Error GoTo LocalCheck
    downloaded = CDate(DOWNLOAD_DATE)
    expiry = DateAdd("m", 3, downloaded)
    
    ' בדיקה מקומית
    If Now() > expiry Then
        Call ShowExpired(expiry)
        CheckLicense = False
        Exit Function
    End If
    
    ' בדיקה מול שרת
    On Error GoTo LocalCheck
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    Dim payload As String
    payload = "{""fileId"":""" & FILE_ID & """,""downloadDate"":""" & DOWNLOAD_DATE & """,""email"":""" & USER_EMAIL & """}"
    http.Open "POST", LICENSE_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send payload
    
    If http.Status = 200 Then
        If InStr(http.responseText, """valid"":false") > 0 Then
            Dim msg As String
            msg = ExtractJson(http.responseText, "message")
            If msg = "" Then msg = "הקובץ אינו תקף. הורד גרסה חדשה מ-kandelbi.org"
            MsgBox Chr(9940) & " " & msg, vbCritical, "קנדל BI"
            CheckLicense = False
            Exit Function
        End If
        Dim daysStr As String
        daysStr = ExtractJson(http.responseText, "daysLeft")
        If daysStr <> "" And CInt(daysStr) <= 14 Then
            MsgBox Chr(9888) & " תוקף הקובץ יפוג בעוד " & daysStr & " ימים." & Chr(13) & "הורד גרסה חדשה מ-kandelbi.org", vbExclamation, "קנדל BI"
        End If
    End If
    
    CheckLicense = True
    Exit Function

LocalCheck:
    If Now() > expiry Then
        Call ShowExpired(expiry)
        CheckLicense = False
    Else
        CheckLicense = True
    End If
End Function

Sub ShowExpired(expiry As Date)
    MsgBox Chr(9940) & " פג תוקף הקובץ" & Chr(13) & Chr(13) & _
           "תוקף הקובץ פג בתאריך " & Format(expiry, "dd/mm/yyyy") & Chr(13) & _
           "אנא הורד גרסה חדשה מ-kandelbi.org", vbCritical, "קנדל BI"
End Sub

Function ExtractJson(json As String, key As String) As String
    Dim pos As Long, s As Long, e As Long
    pos = InStr(json, """" & key & """:")
    If pos = 0 Then Exit Function
    s = pos + Len(key) + 3
    If Mid(json, s, 1) = """" Then
        s = s + 1: e = InStr(s, json, """")
        ExtractJson = Mid(json, s, e - s)
    Else
        e = s
        Do While Mid(json, e, 1) <> "," And Mid(json, e, 1) <> "}" And e < Len(json)
            e = e + 1
        Loop
        ExtractJson = Mid(json, s, e - s)
    End If
End Function

' ===== פונקציות עזר =====
Function TimeToMin(v As Variant) As Double
    On Error GoTo Err1
    If IsEmpty(v) Or v = "" Then TimeToMin = -1: Exit Function
    If IsDate(v) Then
        TimeToMin = (CDbl(CDate(v)) - Int(CDbl(CDate(v)))) * 1440
    ElseIf IsNumeric(v) Then
        If CDbl(v) < 1 Then TimeToMin = CDbl(v) * 1440 Else TimeToMin = -1
    Else
        Dim p() As String: p = Split(CStr(v), ":")
        If UBound(p) >= 1 Then TimeToMin = CDbl(p(0)) * 60 + CDbl(p(1)) Else TimeToMin = -1
    End If
    Exit Function
Err1: TimeToMin = -1
End Function

Function FindCol(ws As Worksheet, kws As Variant) As Integer
    Dim i As Integer, j As Integer
    Dim lc As Integer: lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lc
        Dim h As String: h = LCase(ws.Cells(1, i).Value)
        For j = 0 To UBound(kws)
            If InStr(h, LCase(kws(j))) > 0 Then FindCol = i: Exit Function
        Next j
    Next i
    FindCol = 0
End Function

' ===== בדיקה ראשית =====
Sub RunAudit()
    If Not CheckLicense() Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("נתונים")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "אנא צור גיליון בשם 'נתונים' והכנס את דוח השעות", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Call CheckOverlaps(ws)
    Call CheckReasonableness(ws)
    Call MakeSummary
    Application.ScreenUpdating = True
    
    MsgBox Chr(9989) & " הבדיקה הושלמה! ראה גיליונות: חפיפות שעות, בדיקת סבירות, דוח סיכום", vbInformation, "קנדל BI"
End Sub

' ===== חפיפות שעות =====
Sub CheckOverlaps(ws As Worksheet)
    DeleteSheet "חפיפות שעות"
    Dim rpt As Worksheet
    Set rpt = AddSheet("חפיפות שעות")
    
    Dim hdrs As Variant
    hdrs = Array("מזהה עובד", "תאריך", "כניסה 1", "יציאה 1", "כניסה 2", "יציאה 2", "חפיפה (דק')", "סוג")
    WriteHeaders rpt, hdrs, RGB(13, 31, 60)
    
    Dim empC As Integer, dtC As Integer, stC As Integer, enC As Integer
    empC = FindCol(ws, Array("עובד", "שם", "מזהה", "id"))
    dtC  = FindCol(ws, Array("תאריך", "date", "יום"))
    stC  = FindCol(ws, Array("כניסה", "התחלה", "start", "from", "משעה"))
    enC  = FindCol(ws, Array("יציאה", "סיום", "end", "to", "עד"))
    
    If empC = 0 Or dtC = 0 Or stC = 0 Or enC = 0 Then
        MsgBox "לא נמצאו עמודות נדרשות. וודא עמודות: עובד, תאריך, שעת כניסה, שעת יציאה", vbExclamation
        Exit Sub
    End If
    
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, empC).End(xlUp).Row
    Dim row As Long: row = 2
    Dim i As Long, j As Long
    
    For i = 2 To lr - 1
        For j = i + 1 To lr
            If CStr(ws.Cells(i, empC).Value) = CStr(ws.Cells(j, empC).Value) And _
               CStr(ws.Cells(i, dtC).Value) = CStr(ws.Cells(j, dtC).Value) Then
                
                Dim s1 As Double, e1 As Double, s2 As Double, e2 As Double
                s1 = TimeToMin(ws.Cells(i, stC).Value)
                e1 = TimeToMin(ws.Cells(i, enC).Value)
                s2 = TimeToMin(ws.Cells(j, stC).Value)
                e2 = TimeToMin(ws.Cells(j, enC).Value)
                
                If s1 >= 0 And e1 > 0 And s2 >= 0 And e2 > 0 Then
                    Dim os As Double: os = Application.Max(s1, s2)
                    Dim oe As Double: oe = Application.Min(e1, e2)
                    If oe > os Then
                        rpt.Cells(row, 1).Value = ws.Cells(i, empC).Value
                        rpt.Cells(row, 2).Value = ws.Cells(i, dtC).Value
                        rpt.Cells(row, 3).Value = ws.Cells(i, stC).Value
                        rpt.Cells(row, 4).Value = ws.Cells(i, enC).Value
                        rpt.Cells(row, 5).Value = ws.Cells(j, stC).Value
                        rpt.Cells(row, 6).Value = ws.Cells(j, enC).Value
                        rpt.Cells(row, 7).Value = oe - os
                        
                        Dim isFullOverlap As Boolean
                        isFullOverlap = (s2 >= s1 And e2 <= e1) Or (s1 >= s2 And e1 <= e2)
                        rpt.Cells(row, 8).Value = IIf(isFullOverlap, "חפיפה מלאה", "חפיפה חלקית")
                        rpt.Rows(row).Interior.Color = IIf(isFullOverlap, RGB(254, 242, 242), RGB(255, 248, 237))
                        rpt.Cells(row, 8).Font.Color = IIf(isFullOverlap, RGB(217, 48, 37), RGB(230, 126, 34))
                        row = row + 1
                    End If
                End If
            End If
        Next j
    Next i
    rpt.Columns.AutoFit
End Sub

' ===== סבירות שעות =====
Sub CheckReasonableness(ws As Worksheet)
    DeleteSheet "בדיקת סבירות"
    Dim rpt As Worksheet
    Set rpt = AddSheet("בדיקת סבירות")
    WriteHeaders rpt, Array("עובד", "תאריך", "סה""כ שעות", "ממצא", "פירוט"), RGB(13, 31, 60)
    
    Dim empC As Integer, dtC As Integer, stC As Integer, enC As Integer
    empC = FindCol(ws, Array("עובד", "שם", "מזהה"))
    dtC  = FindCol(ws, Array("תאריך", "date"))
    stC  = FindCol(ws, Array("כניסה", "התחלה", "start"))
    enC  = FindCol(ws, Array("יציאה", "סיום", "end"))
    
    If empC = 0 Or dtC = 0 Then Exit Sub
    
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, empC).End(xlUp).Row
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    
    For i = 2 To lr
        Dim k As String
        k = CStr(ws.Cells(i, empC).Value) & "||" & Format(ws.Cells(i, dtC).Value, "dd/mm/yyyy")
        Dim hrs As Double: hrs = 0
        If stC > 0 And enC > 0 Then
            Dim sm As Double: sm = TimeToMin(ws.Cells(i, stC).Value)
            Dim em As Double: em = TimeToMin(ws.Cells(i, enC).Value)
            If sm >= 0 And em > sm Then hrs = (em - sm) / 60
        End If
        If dict.Exists(k) Then dict(k) = dict(k) + hrs Else dict.Add k, hrs
    Next i
    
    Dim row As Long: row = 2
    Dim kk As Variant
    For Each kk In dict.Keys
        Dim th As Double: th = dict(kk)
        Dim parts() As String: parts = Split(kk, "||")
        Dim issue As String: issue = ""
        Dim detail As String: detail = ""
        
        ' שבת
        On Error Resume Next
        Dim dv As Date: dv = CDate(parts(1))
        On Error GoTo 0
        If Weekday(dv) = 7 Then
            issue = "עבודה בשבת": detail = "נדרש אישור מיוחד"
        End If
        
        ' שעות חריגות
        If th > 12 Then
            issue = issue & IIf(issue <> "", " + ", "") & "שעות חריגות"
            detail = detail & " " & Format(th, "0.0") & " שעות"
        ElseIf th > 10 Then
            issue = issue & IIf(issue <> "", " + ", "") & "שעות גבוהות"
            detail = detail & " " & Format(th, "0.0") & " שעות"
        End If
        
        If issue <> "" Then
            rpt.Cells(row, 1).Value = parts(0)
            rpt.Cells(row, 2).Value = parts(1)
            rpt.Cells(row, 3).Value = th
            rpt.Cells(row, 4).Value = issue
            rpt.Cells(row, 5).Value = Trim(detail)
            rpt.Rows(row).Interior.Color = IIf(th > 12 Or InStr(issue, "שבת") > 0, RGB(254, 242, 242), RGB(255, 248, 237))
            row = row + 1
        End If
    Next kk
    rpt.Columns.AutoFit
End Sub

' ===== סיכום =====
Sub MakeSummary()
    DeleteSheet "דוח סיכום"
    Dim rpt As Worksheet
    Set rpt = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    rpt.Name = "דוח סיכום"
    
    With rpt
        .Range("A1:E1").Merge
        .Cells(1, 1).Value = "דוח בדיקת דוחות תשומות — קנדל BI"
        .Cells(1, 1).Font.Bold = True: .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Interior.Color = RGB(13, 31, 60)
        .Cells(1, 1).Font.Color = RGB(255, 255, 255)
        .Cells(1, 1).HorizontalAlignment = xlCenter
        
        .Cells(2, 1).Value = "תאריך:"
        .Cells(2, 2).Value = Format(Now(), "dd/mm/yyyy hh:mm")
        .Cells(3, 1).Value = "מזהה קובץ:"
        .Cells(3, 2).Value = FILE_ID
        .Cells(4, 1).Value = "משתמש:"
        .Cells(4, 2).Value = USER_EMAIL
        
        .Cells(6, 1).Value = "ממצא"
        .Cells(6, 2).Value = "מספר מקרים"
        With .Range("A6:B6")
            .Interior.Color = RGB(30, 106, 191)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
        
        Dim oc As Long, rc As Long
        On Error Resume Next
        oc = ThisWorkbook.Sheets("חפיפות שעות").Cells(Rows.Count, 1).End(xlUp).Row - 1
        rc = ThisWorkbook.Sheets("בדיקת סבירות").Cells(Rows.Count, 1).End(xlUp).Row - 1
        On Error GoTo 0
        If oc < 0 Then oc = 0
        If rc < 0 Then rc = 0
        
        .Cells(7, 1).Value = "חפיפות שעות"
        .Cells(7, 2).Value = oc
        If oc > 0 Then .Cells(7, 2).Interior.Color = RGB(254, 242, 242)
        
        .Cells(8, 1).Value = "בעיות סבירות"
        .Cells(8, 2).Value = rc
        If rc > 0 Then .Cells(8, 2).Interior.Color = RGB(255, 248, 237)
        
        .Cells(10, 1).Value = "הופק ע""י קנדל BI | kandelbi.org | " & USER_EMAIL
        .Cells(10, 1).Font.Italic = True
        .Cells(10, 1).Font.Color = RGB(107, 122, 153)
        
        .Columns.AutoFit
    End With
End Sub

' ===== עזרי גיליונות =====
Sub DeleteSheet(name As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(name).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Function AddSheet(name As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = name
    Set AddSheet = ws
End Function

Sub WriteHeaders(ws As Worksheet, hdrs As Variant, color As Long)
    Dim i As Integer
    For i = 0 To UBound(hdrs)
        ws.Cells(1, i + 1).Value = hdrs(i)
    Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(hdrs) + 1))
        .Interior.Color = color
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

Sub Auto_Open()
    If Not CheckLicense() Then Exit Sub
    MsgBox "ברוך הבא לכלי בדיקת דוחות תשומות — קנדל BI" & Chr(13) & Chr(13) & _
           "1. הכנס נתוני שעות בגיליון 'נתונים'" & Chr(13) & _
           "2. הפעל מאקרו 'RunAudit' לביצוע הבדיקה" & Chr(13) & Chr(13) & _
           "kandelbi.org", vbInformation, "קנדל BI"
End Sub
