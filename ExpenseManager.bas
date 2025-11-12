Attribute VB_Name = "ExpenseManager"
' ====================================
' Module: ExpenseManager
' برنامج إدارة المصروفات الشهرية
' Monthly Expenses Manager
' ====================================

Option Explicit

' استيراد البيانات من WhatsApp
Sub ImportFromWhatsApp()
    Dim wsData As Worksheet
    Dim wsDashboard As Worksheet
    Dim pastedText As String
    Dim lines() As String
    Dim i As Long
    Dim currentDate As String
    Dim lastRow As Long
    Dim amount As Double
    Dim description As String
    Dim category As String
    Dim expenseDate As Date
    
    Set wsData = ThisWorkbook.Sheets("Database")
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    
    ' الحصول على النص من منطقة الإدخال
    On Error Resume Next
    pastedText = wsDashboard.Range("B5").Value
    On Error GoTo 0
    
    If pastedText = "" Or IsEmpty(pastedText) Then
        MsgBox "الرجاء لصق بيانات WhatsApp أولاً!", vbExclamation, "تنبيه"
        Exit Sub
    End If
    
    ' تقسيم النص إلى أسطر
    lines = Split(pastedText, vbLf)
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
    
    ' معالجة كل سطر
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        If line <> "" Then
            ' التحقق إذا كان السطر يحتوي على تاريخ
            If IsDateLine(line) Then
                currentDate = ExtractArabicDate(line)
                expenseDate = ConvertArabicDate(currentDate)
            ' التحقق إذا كان السطر يحتوي على مصروف
            ElseIf IsNumeric(Left(ConvertArabicNumbers(line), 1)) Then
                amount = ExtractAmount(line)
                description = ExtractDescription(line)
                category = ClassifyExpense(description)
                
                ' إضافة البيانات إلى قاعدة البيانات
                wsData.Cells(lastRow, 1).Value = expenseDate
                wsData.Cells(lastRow, 2).Value = amount
                wsData.Cells(lastRow, 3).Value = description
                wsData.Cells(lastRow, 4).Value = category
                wsData.Cells(lastRow, 5).Value = GetFinancialMonth(expenseDate)
                
                lastRow = lastRow + 1
            End If
        End If
    Next i
    
    ' تنسيق البيانات
    FormatDatabase
    
    ' تحديث لوحة المعلومات
    RefreshDashboard
    
    ' مسح منطقة الإدخال
    wsDashboard.Range("B5:E15").ClearContents
    
    MsgBox "تم استيراد البيانات بنجاح!" & vbCrLf & "عدد المعاملات: " & (lastRow - 2), vbInformation, "نجح الاستيراد"
End Sub

' التحقق إذا كان السطر يحتوي على تاريخ
Function IsDateLine(line As String) As Boolean
    IsDateLine = (InStr(line, "يناير") > 0 Or InStr(line, "فبراير") > 0 Or _
                  InStr(line, "مارس") > 0 Or InStr(line, "أبريل") > 0 Or _
                  InStr(line, "مايو") > 0 Or InStr(line, "يونيو") > 0 Or _
                  InStr(line, "يوليو") > 0 Or InStr(line, "أغسطس") > 0 Or _
                  InStr(line, "سبتمبر") > 0 Or InStr(line, "أكتوبر") > 0 Or _
                  InStr(line, "نوفمبر") > 0 Or InStr(line, "ديسمبر") > 0)
End Function

' استخراج التاريخ العربي
Function ExtractArabicDate(line As String) As String
    Dim cleanLine As String
    cleanLine = Trim(line)
    
    ' إزالة الأقواس والمحتوى بينها [١/‏١١, ١:٣٣ ص]
    If InStr(cleanLine, "]") > 0 Then
        cleanLine = Trim(Mid(cleanLine, InStr(cleanLine, "]") + 1))
    End If
    
    ' إزالة الأسماء (كل شيء بعد ":")
    If InStr(cleanLine, ":") > 0 Then
        cleanLine = Trim(Mid(cleanLine, InStr(cleanLine, ":") + 1))
    End If
    
    ExtractArabicDate = cleanLine
End Function

' تحويل التاريخ العربي إلى تاريخ فعلي
Function ConvertArabicDate(arabicDate As String) As Date
    Dim day As Integer
    Dim monthName As String
    Dim monthNum As Integer
    Dim year As Integer
    Dim parts() As String
    
    ' تحويل الأرقام العربية إلى إنجليزية
    arabicDate = ConvertArabicNumbers(arabicDate)
    
    parts = Split(arabicDate, " ")
    If UBound(parts) >= 1 Then
        day = CInt(parts(0))
        monthName = parts(1)
        
        ' تحويل اسم الشهر إلى رقم
        Select Case monthName
            Case "يناير": monthNum = 1
            Case "فبراير": monthNum = 2
            Case "مارس": monthNum = 3
            Case "أبريل": monthNum = 4
            Case "مايو": monthNum = 5
            Case "يونيو": monthNum = 6
            Case "يوليو": monthNum = 7
            Case "أغسطس": monthNum = 8
            Case "سبتمبر": monthNum = 9
            Case "أكتوبر": monthNum = 10
            Case "نوفمبر": monthNum = 11
            Case "ديسمبر": monthNum = 12
        End Select
        
        ' افتراض السنة الحالية
        year = year(Date)
        
        ConvertArabicDate = DateSerial(year, monthNum, day)
    End If
End Function

' تحويل الأرقام العربية إلى إنجليزية
Function ConvertArabicNumbers(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "٠", "0")
    result = Replace(result, "١", "1")
    result = Replace(result, "٢", "2")
    result = Replace(result, "٣", "3")
    result = Replace(result, "٤", "4")
    result = Replace(result, "٥", "5")
    result = Replace(result, "٦", "6")
    result = Replace(result, "٧", "7")
    result = Replace(result, "٨", "8")
    result = Replace(result, "٩", "9")
    ConvertArabicNumbers = result
End Function

' استخراج المبلغ من السطر
Function ExtractAmount(line As String) As Double
    Dim parts() As String
    Dim amount As String
    
    line = ConvertArabicNumbers(line)
    parts = Split(line, " ")
    
    If UBound(parts) >= 0 Then
        amount = parts(0)
        ExtractAmount = CDbl(amount)
    End If
End Function

' استخراج الوصف من السطر
Function ExtractDescription(line As String) As String
    Dim parts() As String
    Dim description As String
    
    parts = Split(line, " ", 2)
    If UBound(parts) >= 1 Then
        description = Trim(parts(1))
    Else
        description = "غير محدد"
    End If
    
    ExtractDescription = description
End Function

' تصنيف المصروف تلقائياً
Function ClassifyExpense(description As String) As String
    Dim category As String
    description = LCase(description)
    
    ' تصنيف الطعام
    If InStr(description, "مطعم") > 0 Or InStr(description, "بقالة") > 0 Or _
       InStr(description, "طعام") > 0 Or InStr(description, "كافيه") > 0 Or _
       InStr(description, "بوفية") > 0 Or InStr(description, "كبايبو") > 0 Or _
       InStr(description, "عريكة") > 0 Then
        category = "طعام وشراب"
    
    ' تصنيف المواصلات
    ElseIf InStr(description, "بنزين") > 0 Or InStr(description, "توصيلة") > 0 Or _
           InStr(description, "سيارة") > 0 Or InStr(description, "نقل") > 0 Or _
           InStr(description, "طوف") > 0 Then
        category = "مواصلات"
    
    ' تصنيف المصروفات الشخصية
    ElseIf InStr(description, "حلاق") > 0 Or InStr(description, "ملابس") > 0 Or _
           InStr(description, "عطر") > 0 Or InStr(description, "شخصي") > 0 Or _
           InStr(description, "رياضية") > 0 Then
        category = "شخصي"
    
    ' تصنيف المنزل
    ElseIf InStr(description, "كهربائية") > 0 Or InStr(description, "أغراض") > 0 Or _
           InStr(description, "امازون") > 0 Or InStr(description, "بيت") > 0 Or _
           InStr(description, "منزل") > 0 Or InStr(description, "تريندول") > 0 Then
        category = "منزل ومستلزمات"
    
    ' أخرى
    Else
        category = "أخرى"
    End If
    
    ClassifyExpense = category
End Function

' حساب الشهر المالي (من 27 إلى 26)
Function GetFinancialMonth(expenseDate As Date) As String
    Dim financialMonth As String
    Dim monthNum As Integer
    Dim year As Integer
    
    year = year(expenseDate)
    monthNum = Month(expenseDate)
    
    ' إذا كان التاريخ قبل 27، فهو ينتمي للشهر الحالي
    If Day(expenseDate) < 27 Then
        financialMonth = Format(DateSerial(year, monthNum, 1), "yyyy-mm")
    Else
        ' إذا كان 27 أو بعده، فهو ينتمي للشهر التالي
        If monthNum = 12 Then
            financialMonth = Format(DateSerial(year + 1, 1, 1), "yyyy-mm")
        Else
            financialMonth = Format(DateSerial(year, monthNum + 1, 1), "yyyy-mm")
        End If
    End If
    
    GetFinancialMonth = financialMonth
End Function

' تنسيق قاعدة البيانات
Sub FormatDatabase()
    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets("Database")
    
    With wsData
        ' تنسيق العناوين
        With .Range("A1:E1")
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' تنسيق عمود التاريخ
        .Columns("A:A").NumberFormat = "dd/mm/yyyy"
        
        ' تنسيق عمود المبلغ
        .Columns("B:B").NumberFormat = "#,##0.00"
        
        ' توسيط الأعمدة
        .Columns("A:E").AutoFit
    End With
End Sub

' تحديث لوحة المعلومات
Sub RefreshDashboard()
    Dim wsData As Worksheet
    Dim wsDashboard As Worksheet
    Dim currentMonth As String
    Dim totalExpense As Double
    Dim categoryTotals As Object
    Dim i As Long
    Dim lastRow As Long
    
    Set wsData = ThisWorkbook.Sheets("Database")
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set categoryTotals = CreateObject("Scripting.Dictionary")
    
    ' الحصول على الشهر المالي الحالي
    currentMonth = GetFinancialMonth(Date)
    
    ' حساب المجاميع
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If wsData.Cells(i, 5).Value = currentMonth Then
            ' إجمالي المصروفات
            totalExpense = totalExpense + wsData.Cells(i, 2).Value
            
            ' مجاميع الفئات
            Dim cat As String
            cat = wsData.Cells(i, 4).Value
            If categoryTotals.exists(cat) Then
                categoryTotals(cat) = categoryTotals(cat) + wsData.Cells(i, 2).Value
            Else
                categoryTotals.Add cat, wsData.Cells(i, 2).Value
            End If
        End If
    Next i
    
    ' تحديث لوحة المعلومات
    wsDashboard.Range("C18").Value = totalExpense
    
    ' حساب المتوسط اليومي
    Dim daysInPeriod As Integer
    daysInPeriod = Day(Date) - 26
    If daysInPeriod <= 0 Then daysInPeriod = daysInPeriod + 30
    
    If daysInPeriod > 0 Then
        wsDashboard.Range("C19").Value = totalExpense / daysInPeriod
    End If
    
    ' تحديث توزيع الفئات
    UpdateCategoryChart categoryTotals
End Sub

' تحديث رسم الفئات
Sub UpdateCategoryChart(categoryTotals As Object)
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    
    Dim startRow As Long
    startRow = 22
    
    ' مسح البيانات السابقة
    wsDashboard.Range("B22:D30").ClearContents
    
    ' كتابة العناوين
    wsDashboard.Cells(startRow, 2).Value = "الفئة"
    wsDashboard.Cells(startRow, 3).Value = "المبلغ"
    wsDashboard.Cells(startRow, 4).Value = "النسبة"
    
    ' حساب الإجمالي
    Dim total As Double
    Dim key As Variant
    For Each key In categoryTotals.Keys
        total = total + categoryTotals(key)
    Next key
    
    ' كتابة البيانات
    Dim row As Long
    row = startRow + 1
    For Each key In categoryTotals.Keys
        wsDashboard.Cells(row, 2).Value = key
        wsDashboard.Cells(row, 3).Value = categoryTotals(key)
        If total > 0 Then
            wsDashboard.Cells(row, 4).Value = Format(categoryTotals(key) / total, "0.0%")
        End If
        row = row + 1
    Next key
End Sub

' مسح منطقة الإدخال
Sub ClearImportArea()
    Dim wsDashboard As Worksheet
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    wsDashboard.Range("B5:E15").ClearContents
End Sub

' إنشاء تقرير شهري
Sub CreateMonthlyReport()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim currentMonth As String
    Dim lastRow As Long
    Dim i As Long
    Dim reportRow As Long
    
    Set wsData = ThisWorkbook.Sheets("Database")
    
    ' إنشاء ورقة التقرير
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Monthly_Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = "Monthly_Report"
    
    ' إعداد التقرير
    currentMonth = GetFinancialMonth(Date)
    
    wsReport.Cells(1, 1).Value = "تقرير المصروفات الشهرية"
    wsReport.Cells(2, 1).Value = "الشهر المالي: " & currentMonth
    wsReport.Cells(3, 1).Value = "من 27 إلى 26"
    
    ' نسخ البيانات
    wsReport.Cells(5, 1).Value = "التاريخ"
    wsReport.Cells(5, 2).Value = "المبلغ"
    wsReport.Cells(5, 3).Value = "الوصف"
    wsReport.Cells(5, 4).Value = "الفئة"
    
    reportRow = 6
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If wsData.Cells(i, 5).Value = currentMonth Then
            wsReport.Cells(reportRow, 1).Value = wsData.Cells(i, 1).Value
            wsReport.Cells(reportRow, 2).Value = wsData.Cells(i, 2).Value
            wsReport.Cells(reportRow, 3).Value = wsData.Cells(i, 3).Value
            wsReport.Cells(reportRow, 4).Value = wsData.Cells(i, 4).Value
            reportRow = reportRow + 1
        End If
    Next i
    
    ' تنسيق التقرير
    With wsReport
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        .Range("A5:D5").Font.Bold = True
        .Range("A5:D5").Interior.Color = RGB(68, 114, 196)
        .Range("A5:D5").Font.Color = RGB(255, 255, 255)
        .Columns("A:D").AutoFit
    End With
    
    MsgBox "تم إنشاء التقرير الشهري بنجاح!", vbInformation, "تقرير"
End Sub

' إضافة معاملة يدوياً
Sub AddManualTransaction()
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim transDate As Date
    Dim amount As Double
    Dim description As String
    Dim category As String
    
    Set wsData = ThisWorkbook.Sheets("Database")
    
    transDate = InputBox("أدخل التاريخ (dd/mm/yyyy):", "تاريخ المعاملة", Format(Date, "dd/mm/yyyy"))
    If transDate = 0 Then Exit Sub
    
    amount = InputBox("أدخل المبلغ:", "المبلغ")
    If amount = 0 Then Exit Sub
    
    description = InputBox("أدخل الوصف:", "الوصف")
    If description = "" Then Exit Sub
    
    category = InputBox("أدخل الفئة:" & vbCrLf & "طعام وشراب / مواصلات / شخصي / منزل ومستلزمات / أخرى", "الفئة")
    If category = "" Then category = "أخرى"
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row + 1
    
    wsData.Cells(lastRow, 1).Value = transDate
    wsData.Cells(lastRow, 2).Value = amount
    wsData.Cells(lastRow, 3).Value = description
    wsData.Cells(lastRow, 4).Value = category
    wsData.Cells(lastRow, 5).Value = GetFinancialMonth(transDate)
    
    RefreshDashboard
    
    MsgBox "تمت إضافة المعاملة بنجاح!", vbInformation
End Sub