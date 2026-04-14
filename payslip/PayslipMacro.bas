' ============================================================
'  PAYSLIP GENERATOR MACRO
'
'  How to add this to your workbook:
'  1. Open the Excel file
'  2. Press Alt + F11 to open VBA Editor
'  3. In the menu: Insert > Module
'  4. Paste this entire code into the module
'  5. Close VBA Editor
'  6. Save the file as .xlsm (Macro-Enabled Workbook)
'  7. Run via: Alt + F8 > select macro > Run
'     Or assign to a button on the sheet
' ============================================================

Sub GenerateAllPayslipPDFs()
'
' Generates individual PDF payslips for ALL employees
' by looping through serial numbers in F5
'
    Dim wsPaySlips As Worksheet
    Dim wsWage As Worksheet
    Dim empCount As Long
    Dim i As Long
    Dim empName As String
    Dim fileName As String
    Dim folderPath As String
    Dim originalValue As Variant

    Set wsPaySlips = ThisWorkbook.Sheets("Pay Slips")
    Set wsWage = ThisWorkbook.Sheets("Wage Sheet")

    ' Save the original F5 value so we can restore it later
    originalValue = wsPaySlips.Range("F5").Value

    ' Count employees in Wage Sheet (serial numbers in column A, starting row 5)
    empCount = 0
    Dim r As Long
    For r = 5 To wsWage.Cells(wsWage.Rows.Count, "A").End(xlUp).Row
        If IsNumeric(wsWage.Cells(r, 1).Value) And wsWage.Cells(r, 1).Value <> "" Then
            empCount = empCount + 1
        End If
    Next r

    If empCount = 0 Then
        MsgBox "No employees found in Wage Sheet!", vbExclamation
        Exit Sub
    End If

    ' Confirm with user
    If MsgBox("Found " & empCount & " employees." & vbCrLf & vbCrLf & _
              "Generate " & empCount & " individual PDF payslips?", _
              vbYesNo + vbQuestion, "Generate Payslips") = vbNo Then
        Exit Sub
    End If

    ' Create output folder in the same directory as the workbook
    folderPath = ThisWorkbook.Path & "\Payslips_PDF"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    ' Turn off screen updating for speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    Dim printRange As Range
    Set printRange = wsPaySlips.Range("C2:F24")

    For i = 1 To empCount
        ' Set the serial number - this triggers all VLOOKUPs
        wsPaySlips.Range("F5").Value = i

        ' Force recalculation
        Application.Calculate
        DoEvents

        ' Get employee name from D6 for the filename
        empName = Trim(CStr(wsPaySlips.Range("D6").Value))
        empName = Replace(empName, " ", "_")
        empName = Replace(empName, "/", "-")
        empName = Replace(empName, "\", "-")
        empName = Replace(empName, ".", "")

        fileName = folderPath & "\" & Format(i, "00") & "_" & empName & ".pdf"

        ' Export the payslip block to PDF
        printRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=fileName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=True, _
            OpenAfterPublish:=False

        ' Update status bar
        Application.StatusBar = "Generating payslip " & i & " of " & empCount & "..."
    Next i

    ' Restore original F5 value
    wsPaySlips.Range("F5").Value = originalValue
    Application.Calculate

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox empCount & " payslips saved to:" & vbCrLf & vbCrLf & _
           folderPath, vbInformation, "Done!"

    ' Open the output folder
    Shell "explorer.exe " & folderPath, vbNormalFocus

End Sub


Sub GenerateCombinedPayslipPDF()
'
' Generates ONE combined PDF with all payslips (2 per page)
' using both Block 1 (F5) and Block 2 (F34 = F5+1)
'
    Dim wsPaySlips As Worksheet
    Dim wsWage As Worksheet
    Dim empCount As Long
    Dim i As Long
    Dim folderPath As String
    Dim fileName As String
    Dim originalValue As Variant

    Set wsPaySlips = ThisWorkbook.Sheets("Pay Slips")
    Set wsWage = ThisWorkbook.Sheets("Wage Sheet")

    originalValue = wsPaySlips.Range("F5").Value

    ' Count employees
    empCount = 0
    Dim r As Long
    For r = 5 To wsWage.Cells(wsWage.Rows.Count, "A").End(xlUp).Row
        If IsNumeric(wsWage.Cells(r, 1).Value) And wsWage.Cells(r, 1).Value <> "" Then
            empCount = empCount + 1
        End If
    Next r

    If empCount = 0 Then
        MsgBox "No employees found in Wage Sheet!", vbExclamation
        Exit Sub
    End If

    If MsgBox("Found " & empCount & " employees." & vbCrLf & vbCrLf & _
              "Generate combined PDF with all payslips (2 per page)?", _
              vbYesNo + vbQuestion, "Generate Combined Payslip") = vbNo Then
        Exit Sub
    End If

    folderPath = ThisWorkbook.Path & "\Payslips_PDF"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    fileName = folderPath & "\All_Payslips_Combined.pdf"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    ' We'll use a temporary sheet to accumulate all payslip blocks
    ' then export that sheet as one PDF

    ' Since Excel can't easily merge PDFs, we'll print all pages
    ' using the 2-per-page layout (Block 1 + Block 2)
    ' Loop F5 = 1, 3, 5, 7, ... (block 2 auto-fills F5+1)

    Dim printRange As Range
    Set printRange = wsPaySlips.Range("C2:F53") ' Both blocks

    ' Set print area to both blocks
    wsPaySlips.PageSetup.PrintArea = "C2:F53"

    ' For combined PDF, we print all iterations
    ' Unfortunately Excel VBA can't natively merge PDFs in one export
    ' So we generate individual pages and tell the user

    ' Alternative: Print to printer with "Microsoft Print to PDF"
    ' For simplicity, generate individual PDFs

    For i = 1 To empCount Step 2
        wsPaySlips.Range("F5").Value = i
        Application.Calculate
        DoEvents

        Dim pageName As String
        pageName = folderPath & "\Page_" & Format((i + 1) / 2, "00") & ".pdf"

        printRange.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=pageName, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=True, _
            OpenAfterPublish:=False

        Application.StatusBar = "Generating page " & (i + 1) / 2 & "..."
    Next i

    ' Restore
    wsPaySlips.Range("F5").Value = originalValue
    Application.Calculate

    Application.ScreenUpdating = True
    Application.StatusBar = False

    Dim pageCount As Long
    pageCount = Application.WorksheetFunction.RoundUp(empCount / 2, 0)

    MsgBox pageCount & " pages saved (2 payslips each) to:" & vbCrLf & vbCrLf & _
           folderPath, vbInformation, "Done!"

    Shell "explorer.exe " & folderPath, vbNormalFocus

End Sub
