
'Sub DropDown18_Change()
'    Worksheets("Invoice").DropDown18.List = Application.WorksheetFunction.Transpose(Worksheets("Config").Range("Consultant"))
'End Sub

Sub AddArrayToDropDownOnWorksheet()
'DropDown16 is DropDowns(1)
'DropDown18 is DropDowns(2)
'Populates the DropDown list contents from Horizontal range in Worksheet "Config"
  Sheets("Invoice").DropDowns(2).List = Sheets("Config").Range("ConfigConsultant")

End Sub

Sub transCon()
    Worksheets("Config").Range("Consultant").Copy
    Worksheets("Config").Range("Z1").PasteSpecial Paste:=xlPasteValues, Transpose:=True
End Sub

Sub backup()
    Worksheets("Invoice").Range("Invoice").Copy Destination:=Worksheets("Backup").Range("A1")
End Sub

Sub restore()
    Worksheets("Backup").Range("Invoice").Copy Destination:=Worksheets("Invoice").Range("A1")
End Sub

Sub checkInvoiceNo()
    If Range("Ledger_Invoice_No").End(xlDown).Value < 1000 Then
        Range("Invoice_No").Value = 1000
    Else
        Range("Invoice_No").Value = Range("Ledger_Invoice_No").End(xlDown).Value + 1
    End If
End Sub

Sub Auto_Open()
    Worksheets("Invoice").Activate
    Range("Date").Value = Date
    checkInvoiceNo
End Sub

Sub PostToRegister()
    Dim WS1 As Worksheet
    Dim WS2 As Worksheet
    Set WS1 = Worksheets("Invoice")
    Set WS2 = Worksheets("Ledger")
    NextRow = WS2.Cells(Rows.Count, 1).End(xlUp).Row + 1
    WS2.Cells(NextRow, 1).Resize(1, 5).Value = Array(WS1.Range("Date"), WS1.Range("Invoice_No"), WS1.Range("PO_Number"), WS1.Range("InvoiceClient"), WS1.Range("TOTAL"))
End Sub

Sub ClearReset()
    Range("Date").Value = Date
'    Range("Invoie_No").Value = 1000
    checkInvoiceNo
    Range("Details").ClearContents
    Range("E1").ClearContents
    Range("E2").ClearContents
    Range("F3").ClearContents
    Range("F4").ClearContents
    Range("PO_Number").ClearContents
End Sub

Sub NextInvoice()
'    Temp = Range("Invoice_No").Value + 1
    ClearReset
'    Range("Invoice_No").Value = Temp
    ActiveWorkbook.Save
End Sub

' This Function only required for MAC SaveInvWithNewName
Function FileExists(ByVal AFileName As Variant) As Boolean
'    On Error GoTo Catch

    FileSystem.FileLen AFileName

'    FileExists = True

'    GoTo Finally

'Catch:
'        FileExists = False
'Finally:
End Function

Function SaveInvWithNewNameMAC()
    Dim NewFN As Variant
    Dim NewPDF As Variant
'    \ for Windows / For Mac??
    NewFN = ThisWorkbook.Path & "/" & "Invoice " & Range("Invoice_No").Value & ".xlsx"
    NewPDF = ThisWorkbook.Path & "/" & "Invoice " & Range("Invoice_No").Value & ".pdf"

    Dim FSO
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    If FSO.FileExists(NewPDF) Or FSO.FileExists(NewFN) Then
    If FileExists(NewPDF) Or FileExists(NewFN) Then
        GoTo Terminate
    End If
End Function

Function SaveInvWithNewNameWIN()

    Dim NewFN As Variant
    Dim NewPDF As Variant
    NewFN = ThisWorkbook.Path & "\" & "Invoice " & Range("Invoice_No").Value & ".xlsx"
    NewPDF = ThisWorkbook.Path & "\" & "Invoice " & Range("Invoice_No").Value & ".pdf"

    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(NewPDF) Or FSO.FileExists(NewFN) Then
        GoTo Terminate
    End If

End Function

Function WINorMAC()
' Test the conditional compiler constant
  #If Mac Then
    ' Iam a Mac and will test if it is Excel 2011 or higher
    If Val(Application.Version) > 14 Then
      SaveInvWithNewNameMAC
    End If
  #Else
    ' I am Windows
    SaveInvWithNewNameWIN
  #End If
End Function

Sub SaveInvWithNewName()

    WINorMAC

    PostToRegister

    ActiveSheet.Copy
'    For Each Shape In ActiveSheet.Shapes
'        Shape.Delete
'    Next Shape

    Dim ws As Worksheet

    Set ws = ActiveSheet

    With ws.UsedRange
        .Copy
        .PasteSpecial Paste:=xlPasteValues, _
        Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

    ActiveSheet.Protect "password", True, True
    ActiveWorkbook.SaveAs NewFN, xlOpenXMLWorkbook, ReadOnlyRecommended
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=NewPDF, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ActiveWorkbook.Close SaveChanges:=False
    NextInvoice

    Exit Sub
Terminate:
        MsgBox "Specified File Exists", vbInformation, "Exists"
    End
End Sub


'Private Sub Worksheet_Change(ByVal Target As Range)
'
'  Dim NextRow As Long
'
'  Application.EnableEvents = False
'    If Target.Address = "$J$2" Then
'       NextRow = Cells(Rows.Count, "K").End(xlUp).Row + 1
'       Cells(NextRow, "K") = Target.Value
'    End If
'  Application.EnableEvents = True
'
'End Sub







Private Sub Workbook_Open()

End Sub




