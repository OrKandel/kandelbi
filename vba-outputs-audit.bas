' ============================================================
' קנדל BI — כלי בדיקת תפוקות v1.0 | kandelbi.org
' ============================================================
' הוראות: פתח אקסל חדש → שמור כ-.xlsm → Alt+F11 → Insert Module → הדבק

Const FILE_ID As String = "REPLACE_FILE_ID"
Const DOWNLOAD_DATE As String = "REPLACE_DATE"
Const USER_EMAIL As String = "REPLACE_EMAIL"
Const LICENSE_URL As String = "https://license-check.orkandel11.workers.dev/"

Function CheckLicense() As Boolean
    Dim dl As Date, ex As Date
    On Error GoTo lc
    dl = CDate(DOWNLOAD_DATE): ex = DateAdd("m", 3, dl)
    If Now() > ex Then
        MsgBox Chr(9940) & " פג תוקף הקובץ" & Chr(13) & "תוקף: " & Format(ex, "dd/mm/yyyy") & Chr(13) & "הורד חדש מ-kandelbi.org", vbCritical, "קנדל BI"
        CheckLicense = False: Exit Function
    End If
    On Error GoTo lc
    Dim h As Object: Set h = CreateObject("MSXML2.XMLHTTP")
    h.Open "POST", LICENSE_URL, False
    h.setRequestHeader "Content-Type", "application/json"
    h.send "{""fileId"":""" & FILE_ID & """,""downloadDate"":""" & DOWNLOAD_DATE & """,""email"":""" & USER_EMAIL & """}"
    If h.Status = 200 And InStr(h.responseText, """valid"":false") > 0 Then
        MsgBox Chr(9940) & " הקובץ אינו תקף. הורד חדש מ-kandelbi.org", vbCritical, "קנדל BI"
        CheckLicense = False: Exit Function
    End If
    CheckLicense = True: Exit Function
lc: CheckLicense = (Now() <= ex)
End Function

Function FindCol(ws As Worksheet, kws As Variant) As Integer
    Dim i As Integer, j As Integer, lc As Integer
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lc
        Dim h As String: h = LCase(ws.Cells(1, i).Value)
        For j = 0 To UBound(kws)
            If InStr(h, LCase(kws(j))) > 0 Then FindCol = i: Exit Function
        Next j
    Next i
End Function

Sub DeleteSheet(n As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(n).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Function AddSheet(n As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = n: Set AddSheet = ws
End Function

Sub Hdrs(ws As Worksheet, h As Variant, c As Long)
    Dim i As Integer
    For i = 0 To UBound(h): ws.Cells(1, i + 1).Value = h(i): Next i
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(h) + 1))
        .Interior.Color = c: .Font.Color = RGB(255, 255, 255): .Font.Bold = True
    End With
End Sub

Sub RunOutputsAudit()
    If Not CheckLicense() Then Exit Sub
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Sheets("נתונים"): On Error GoTo 0
    If ws Is Nothing Then MsgBox "צור גיליון 'נתונים' עם דוח התפוקות", vbExclamation: Exit Sub
    Application.ScreenUpdating = False
    Call NewOutputTypes(ws)
    Call ReasonCheck(ws)
    Call RiskSample(ws)
    Call MonthlyDashboard(ws)
    Call OutputsSummary
    Application.ScreenUpdating = True
    MsgBox Chr(9989) & " הושלם! גיליונות: סוגי תפוקה חדשים, בדיקת סבירות, מדגם, דשבורד, סיכום", vbInformation, "קנדל BI"
End Sub

Sub NewOutputTypes(ws As Worksheet)
    DeleteSheet "סוגי תפוקה חדשים"
    Dim r As Worksheet: Set r = AddSheet("סוגי תפוקה חדשים")
    Hdrs r, Array("סוג תפוקה", "חודש ראשון", "כמות שורות", "הערה"), RGB(13, 31, 60)
    Dim tc As Integer, mc As Integer
    tc = FindCol(ws, Array("סוג תפוקה", "תפוקה", "שירות", "type"))
    mc = FindCol(ws, Array("חודש", "תאריך", "month", "date"))
    If tc = 0 Then MsgBox "חסרה עמודת סוג תפוקה", vbExclamation: Exit Sub
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, tc).End(xlUp).Row
    Dim first As Object: Set first = CreateObject("Scripting.Dictionary")
    Dim cnts As Object: Set cnts = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 2 To lr
        Dim tp As String: tp = CStr(ws.Cells(i, tc).Value)
        Dim mn As String: mn = IIf(mc > 0, Format(ws.Cells(i, mc).Value, "mm/yyyy"), "לא ידוע")
        If tp <> "" Then
            If Not first.Exists(tp) Or mn < first(tp) Then
                If Not first.Exists(tp) Then first.Add tp, mn Else first(tp) = mn
            End If
            If cnts.Exists(tp) Then cnts(tp) = cnts(tp) + 1 Else cnts.Add tp, 1
        End If
    Next i
    Dim minM As String: minM = "99/9999"
    Dim k As Variant
    For Each k In first.Keys: If first(k) < minM Then minM = first(k): Next k
    Dim row As Long: row = 2
    For Each k In first.Keys
        If first(k) > minM Then
            r.Cells(row, 1).Value = k: r.Cells(row, 2).Value = first(k)
            r.Cells(row, 3).Value = cnts(k): r.Cells(row, 4).Value = "סוג חדש — לא היה בחודש הראשון"
            r.Rows(row).Interior.Color = RGB(255, 248, 237): row = row + 1
        End If
    Next k
    If row = 2 Then r.Cells(2, 1).Value = "לא נמצאו סוגי תפוקה חדשים"
    r.Columns.AutoFit
End Sub

Sub ReasonCheck(ws As Worksheet)
    DeleteSheet "בדיקת סבירות תפוקות"
    Dim r As Worksheet: Set r = AddSheet("בדיקת סבירות תפוקות")
    Hdrs r, Array("עובד", "סוג תפוקה", "חודש", "כמות", "ממוצע", "סטיה%", "ממצא"), RGB(13, 31, 60)
    Dim ec As Integer, tc As Integer, mc As Integer, qc As Integer
    ec = FindCol(ws, Array("עובד", "שם", "id"))
    tc = FindCol(ws, Array("סוג תפוקה", "תפוקה", "שירות"))
    mc = FindCol(ws, Array("חודש", "תאריך"))
    qc = FindCol(ws, Array("כמות", "qty", "quantity"))
    If tc = 0 Or qc = 0 Then Exit Sub
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, IIf(ec > 0, ec, tc)).End(xlUp).Row
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    Dim cnts As Object: Set cnts = CreateObject("Scripting.Dictionary")
    Dim byM As Object: Set byM = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 2 To lr
        Dim emp As String: emp = IIf(ec > 0, CStr(ws.Cells(i, ec).Value), "כללי")
        Dim tp As String: tp = CStr(ws.Cells(i, tc).Value)
        Dim mn As String: mn = IIf(mc > 0, Format(ws.Cells(i, mc).Value, "mm/yyyy"), "כללי")
        Dim qty As Double: qty = IIf(IsNumeric(ws.Cells(i, qc).Value), CDbl(ws.Cells(i, qc).Value), 0)
        Dim k As String: k = emp & "||" & tp
        Dim km As String: km = k & "||" & mn
        If sums.Exists(k) Then sums(k) = sums(k) + qty: cnts(k) = cnts(k) + 1 Else sums.Add k, qty: cnts.Add k, 1
        If byM.Exists(km) Then byM(km) = byM(km) + qty Else byM.Add km, qty
    Next i
    Dim row As Long: row = 2
    Dim kk As Variant
    For Each kk In byM.Keys
        Dim parts() As String: parts = Split(kk, "||")
        Dim bk As String: bk = parts(0) & "||" & parts(1)
        Dim avg As Double: avg = IIf(cnts.Exists(bk) And cnts(bk) > 0, sums(bk) / cnts(bk), 0)
        Dim mq As Double: mq = byM(kk)
        Dim dev As Double: dev = IIf(avg > 0, (mq - avg) / avg * 100, 0)
        If Abs(dev) > 50 Then
            r.Cells(row, 1).Value = parts(0): r.Cells(row, 2).Value = parts(1)
            r.Cells(row, 3).Value = parts(2): r.Cells(row, 4).Value = mq
            r.Cells(row, 5).Value = Format(avg, "0.0"): r.Cells(row, 6).Value = Format(dev, "0.0") & "%"
            r.Cells(row, 7).Value = IIf(dev > 100, "קפיצה חריגה מאוד", IIf(dev > 50, "גידול חריג", "ירידה חריגה"))
            r.Rows(row).Interior.Color = IIf(dev > 100, RGB(254, 242, 242), IIf(dev > 50, RGB(255, 248, 237), RGB(238, 242, 251)))
            row = row + 1
        End If
    Next kk
    If row = 2 Then r.Cells(2, 1).Value = "לא נמצאו סטיות חריגות (>50%)"
    r.Columns.AutoFit
End Sub

Sub RiskSample(ws As Worksheet)
    DeleteSheet "מדגם סיכונים"
    Dim r As Worksheet: Set r = AddSheet("מדגם סיכונים")
    Hdrs r, Array("שורה", "עובד", "סוג תפוקה", "חודש", "כמות", "סיכון", "סיבה"), RGB(13, 31, 60)
    Dim ec As Integer, tc As Integer, mc As Integer, qc As Integer
    ec = FindCol(ws, Array("עובד", "שם")): tc = FindCol(ws, Array("סוג תפוקה", "תפוקה"))
    mc = FindCol(ws, Array("חודש", "תאריך")): qc = FindCol(ws, Array("כמות", "qty"))
    If tc = 0 Then Exit Sub
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, IIf(ec > 0, ec, tc)).End(xlUp).Row
    Dim avgs As Object: Set avgs = CreateObject("Scripting.Dictionary")
    Dim cnts As Object: Set cnts = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 2 To lr
        Dim tp As String: tp = CStr(ws.Cells(i, tc).Value)
        Dim qty As Double: qty = IIf(IsNumeric(ws.Cells(i, qc).Value), CDbl(ws.Cells(i, qc).Value), 0)
        If avgs.Exists(tp) Then avgs(tp) = avgs(tp) + qty: cnts(tp) = cnts(tp) + 1 Else avgs.Add tp, qty: cnts.Add tp, 1
    Next i
    Dim row As Long: row = 2
    For i = 2 To lr
        Dim emp As String: emp = IIf(ec > 0, CStr(ws.Cells(i, ec).Value), "")
        Dim tpv As String: tpv = CStr(ws.Cells(i, tc).Value)
        Dim mn As String: mn = IIf(mc > 0, Format(ws.Cells(i, mc).Value, "mm/yyyy"), "")
        Dim qv As Double: qv = IIf(IsNumeric(ws.Cells(i, qc).Value), CDbl(ws.Cells(i, qc).Value), 0)
        Dim avg As Double: avg = IIf(cnts.Exists(tpv) And cnts(tpv) > 0, avgs(tpv) / cnts(tpv), 0)
        Dim risk As String: risk = "": Dim reason As String: reason = ""
        If avg > 0 And qv > avg * 2 Then risk = "גבוה": reason = "כמות גבוהה פי 2+"
        If avg > 0 And qv < avg * 0.3 Then risk = "בינוני": reason = "כמות נמוכה מ-30%"
        If risk = "" And (i Mod 10 = 0) Then risk = "נמוך": reason = "מדגם אקראי 10%"
        If risk <> "" Then
            r.Cells(row, 1).Value = i: r.Cells(row, 2).Value = emp
            r.Cells(row, 3).Value = tpv: r.Cells(row, 4).Value = mn
            r.Cells(row, 5).Value = qv: r.Cells(row, 6).Value = risk: r.Cells(row, 7).Value = reason
            r.Rows(row).Interior.Color = IIf(risk = "גבוה", RGB(254, 242, 242), IIf(risk = "בינוני", RGB(255, 248, 237), RGB(238, 242, 251)))
            row = row + 1
        End If
    Next i
    r.Columns.AutoFit
End Sub

Sub MonthlyDashboard(ws As Worksheet)
    DeleteSheet "דשבורד חודשי"
    Dim db As Worksheet: Set db = AddSheet("דשבורד חודשי")
    db.Range("A1:H1").Merge
    db.Cells(1, 1).Value = "דשבורד תפוקות חודשי — קנדל BI"
    db.Cells(1, 1).Font.Bold = True: db.Cells(1, 1).Font.Size = 14
    db.Cells(1, 1).Interior.Color = RGB(13, 31, 60): db.Cells(1, 1).Font.Color = RGB(255, 255, 255)
    db.Cells(1, 1).HorizontalAlignment = xlCenter
    Dim tc As Integer, mc As Integer, qc As Integer
    tc = FindCol(ws, Array("סוג תפוקה", "תפוקה", "שירות"))
    mc = FindCol(ws, Array("חודש", "תאריך"))
    qc = FindCol(ws, Array("כמות", "qty"))
    If tc = 0 Or mc = 0 Or qc = 0 Then db.Cells(3, 1).Value = "חסרות עמודות": Exit Sub
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, tc).End(xlUp).Row
    Dim mSet As Object: Set mSet = CreateObject("Scripting.Dictionary")
    Dim tSet As Object: Set tSet = CreateObject("Scripting.Dictionary")
    Dim dat As Object: Set dat = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 2 To lr
        Dim mn As String: mn = Format(ws.Cells(i, mc).Value, "mm/yyyy")
        Dim tp As String: tp = CStr(ws.Cells(i, tc).Value)
        Dim qty As Double: qty = IIf(IsNumeric(ws.Cells(i, qc).Value), CDbl(ws.Cells(i, qc).Value), 0)
        If Not mSet.Exists(mn) Then mSet.Add mn, True
        If Not tSet.Exists(tp) Then tSet.Add tp, True
        Dim k As String: k = mn & "||" & tp
        If dat.Exists(k) Then dat(k) = dat(k) + qty Else dat.Add k, qty
    Next i
    Dim mArr() As String: ReDim mArr(mSet.Count - 1)
    Dim mi As Integer: mi = 0
    Dim mk As Variant
    For Each mk In mSet.Keys: mArr(mi) = mk: mi = mi + 1: Next mk
    Dim tmp As String, m1 As Integer, m2 As Integer
    For m1 = 0 To UBound(mArr) - 1
        For m2 = m1 + 1 To UBound(mArr)
            If mArr(m1) > mArr(m2) Then tmp = mArr(m1): mArr(m1) = mArr(m2): mArr(m2) = tmp
        Next m2
    Next m1
    Dim sr As Long: sr = 3
    db.Cells(sr, 1).Value = "סוג תפוקה": db.Cells(sr, 1).Font.Bold = True
    Dim c As Integer
    For c = 0 To UBound(mArr)
        db.Cells(sr, c + 2).Value = mArr(c): db.Cells(sr, c + 2).Font.Bold = True
        db.Cells(sr, c + 2).HorizontalAlignment = xlCenter
    Next c
    db.Cells(sr, UBound(mArr) + 3).Value = "סה""כ": db.Cells(sr, UBound(mArr) + 3).Font.Bold = True
    With db.Range(db.Cells(sr, 1), db.Cells(sr, UBound(mArr) + 3))
        .Interior.Color = RGB(30, 106, 191): .Font.Color = RGB(255, 255, 255)
    End With
    Dim dr As Long: dr = sr + 1
    Dim tk As Variant
    For Each tk In tSet.Keys
        db.Cells(dr, 1).Value = tk
        Dim rt As Double: rt = 0
        For c = 0 To UBound(mArr)
            Dim dk As String: dk = mArr(c) & "||" & tk
            Dim v As Double: v = IIf(dat.Exists(dk), dat(dk), 0)
            db.Cells(dr, c + 2).Value = v: db.Cells(dr, c + 2).HorizontalAlignment = xlCenter
            rt = rt + v
            If c > 0 Then
                Dim pk As String: pk = mArr(c - 1) & "||" & tk
                Dim pv As Double: pv = IIf(dat.Exists(pk), dat(pk), 0)
                If pv > 0 And v > pv * 1.5 Then db.Cells(dr, c + 2).Interior.Color = RGB(254, 242, 242)
                If pv > 0 And v < pv * 0.5 Then db.Cells(dr, c + 2).Interior.Color = RGB(255, 248, 237)
            End If
        Next c
        db.Cells(dr, UBound(mArr) + 3).Value = rt: db.Cells(dr, UBound(mArr) + 3).Font.Bold = True
        If dr Mod 2 = 0 Then db.Rows(dr).Interior.Color = RGB(248, 250, 252)
        dr = dr + 1
    Next tk
    db.Cells(dr, 1).Value = "סה""כ": db.Cells(dr, 1).Font.Bold = True
    For c = 0 To UBound(mArr)
        Dim ct As Double: ct = 0
        For Each tk In tSet.Keys
            Dim cdk As String: cdk = mArr(c) & "||" & tk
            If dat.Exists(cdk) Then ct = ct + dat(cdk)
        Next tk
        db.Cells(dr, c + 2).Value = ct: db.Cells(dr, c + 2).Font.Bold = True
    Next c
    With db.Rows(dr): .Interior.Color = RGB(13, 31, 60): .Font.Color = RGB(255, 255, 255): .Font.Bold = True: End With
    db.Cells(dr + 2, 1).Value = "אגדה: אדום = קפיצה >50% | צהוב = ירידה >50%"
    db.Cells(dr + 2, 1).Font.Italic = True: db.Cells(dr + 2, 1).Font.Color = RGB(107, 122, 153)
    db.Columns.AutoFit
End Sub

Sub OutputsSummary()
    DeleteSheet "דוח סיכום"
    Dim r As Worksheet: Set r = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    r.Name = "דוח סיכום"
    With r
        .Range("A1:F1").Merge
        .Cells(1, 1).Value = "דוח בדיקת תפוקות — קנדל BI"
        .Cells(1, 1).Font.Bold = True: .Cells(1, 1).Font.Size = 14
        .Cells(1, 1).Interior.Color = RGB(13, 31, 60): .Cells(1, 1).Font.Color = RGB(255, 255, 255)
        .Cells(1, 1).HorizontalAlignment = xlCenter
        .Cells(2, 1).Value = "תאריך:": .Cells(2, 2).Value = Format(Now(), "dd/mm/yyyy hh:mm")
        .Cells(3, 1).Value = "מזהה קובץ:": .Cells(3, 2).Value = FILE_ID
        .Cells(4, 1).Value = "משתמש:": .Cells(4, 2).Value = USER_EMAIL
        With .Range("A6:B6")
            .Cells(1).Value = "ממצא": .Cells(2).Value = "כמות"
            .Interior.Color = RGB(30, 106, 191): .Font.Color = RGB(255, 255, 255): .Font.Bold = True
        End With
        Dim n1 As Long, n2 As Long, n3 As Long
        On Error Resume Next
        n1 = ThisWorkbook.Sheets("סוגי תפוקה חדשים").Cells(Rows.Count, 1).End(xlUp).Row - 1
        n2 = ThisWorkbook.Sheets("בדיקת סבירות תפוקות").Cells(Rows.Count, 1).End(xlUp).Row - 1
        n3 = ThisWorkbook.Sheets("מדגם סיכונים").Cells(Rows.Count, 1).End(xlUp).Row - 1
        On Error GoTo 0
        If n1 < 0 Then n1 = 0: If n2 < 0 Then n2 = 0: If n3 < 0 Then n3 = 0
        .Cells(7, 1).Value = "סוגי תפוקה חדשים": .Cells(7, 2).Value = n1
        .Cells(8, 1).Value = "חריגות סבירות": .Cells(8, 2).Value = n2
        .Cells(9, 1).Value = "שורות במדגם": .Cells(9, 2).Value = n3
        If n1 > 0 Then .Cells(7, 2).Interior.Color = RGB(255, 248, 237)
        If n2 > 0 Then .Cells(8, 2).Interior.Color = RGB(254, 242, 242)
        .Cells(11, 1).Value = "הופק ע""י קנדל BI | kandelbi.org"
        .Cells(11, 1).Font.Italic = True: .Cells(11, 1).Font.Color = RGB(107, 122, 153)
        .Columns.AutoFit
    End With
End Sub

Sub Auto_Open()
    If Not CheckLicense() Then Exit Sub
    MsgBox "ברוך הבא לכלי בדיקת תפוקות — קנדל BI" & Chr(13) & Chr(13) & _
           "1. הכנס נתוני תפוקות בגיליון 'נתונים'" & Chr(13) & _
           "   עמודות: עובד, סוג תפוקה, חודש, כמות" & Chr(13) & _
           "2. הפעל מאקרו 'RunOutputsAudit'", vbInformation, "קנדל BI"
End Sub
