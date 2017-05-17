Attribute VB_Name = "Hulpfuncties"
Sub UnhideSheets() 'v41
Dim x As Integer

With ActiveWorkbook
    x = .Worksheets.Count
    For t = 1 To x
        .Worksheets(t).Visible = True
    Next
End With

End Sub

Sub VeryHideSheet() 'v41
    ActiveSheet.Visible = xlVeryHidden
End Sub

Sub RealEmpty()
Dim rng As Range, selectrange As Range
    With ActiveWorkbook.ActiveSheet
    Set selectrange = Range("a1").Resize(.Cells(.Rows.Count, 1).End(xlUp).Row, _
    .Cells(1, Columns.Count).End(xlToLeft).Column)
        For Each rng In selectrange
            If rng = "" Then rng.ClearContents
        Next
    End With
End Sub

Sub UnprotectAllSheets()
    Dim sh As Worksheet
    Dim myPassword As String
    
    myPassword = InputBox("Voer wachtwoord in: ", "Unlock tabbladen", "S3rVic3")
    
    'myPassword = "S3rVic3"
    shcnt = ActiveWorkbook.Sheets.Count
    For Each sh In ActiveWorkbook.Worksheets
    i_cnt = i_cnt + 1
        sh.Unprotect Password:=myPassword
        sh.Protect Password:=""
        sh.Unprotect
        sh.Visible = xlSheetVisible
        On Error GoTo errh
    Next sh
    MsgBox ("Aantal tabbladen vrijgegeven = " & i_cnt)
    Exit Sub
errh:
MsgBox ("Fout in sheet " & sh.Name)
Resume Next
End Sub

Public Function fCountColor(range_data As Range, criteria As Range) As Long
    Application.Volatile
    Dim datax As Range
    Dim xcolor As Long
xcolor = criteria.Interior.ColorIndex
For Each datax In range_data
    If datax.Interior.ColorIndex = xcolor Then
        fCountColor = fCountColor + 1
    End If
Next datax
End Function

Function fConvertToLetter(iCol As Integer) As String
    Application.Volatile
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      fConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      fConvertToLetter = fConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Function fJoinRange(rg As Range, Optional delimiter As String = ", ") As String
    Application.Volatile
    For Each s In rg
    If fJoinRange = "" Then fJoinRange = s Else fJoinRange = fJoinRange & delimiter & s
    Next
End Function

Public Function fMaxif(srcrw As Range, Optional argrw As Range, Optional arg As Variant) As Variant
Dim getal As Double
Dim tempmax As Double
Dim inhoud() As Variant
Dim temp As Variant

    'Beide reeksen in een Array
    
    ReDim inhoud(1 To srcrw.Count, 1 To 2)
    
    teller = 0
    For Each temp In srcrw
        teller = teller + 1
        inhoud(teller, 1) = temp
    Next
    
    teller = 0
    For Each temp In argrw
        teller = teller + 1
        inhoud(teller, 2) = temp
    Next
    
    For teller = 1 To UBound(inhoud, 1)
        If inhoud(teller, 2) = arg And inhoud(teller, 1) > tempmax Then tempmax = inhoud(teller, 1)
    Next
    
    
    fMaxif = tempmax
End Function
Function fThConditiescore(Bouwjaar As Integer, ThLevensduur As Integer, Optional peiljaar As Integer) As Variant 'v52
Application.Volatile
Dim c As Double
Dim t As Double, L As Double
'C = 1 +  ½ log (1 – t / L)

'L = theoretische levensduur

    If IsEmpty(peiljaar) Or peiljaar < 1952 Or peiljaar > Year(Date) + 50 Then peiljaar = Year(Date)
    If ThLevensduur <= 0 Or ThLevensduur > 50 Then GoTo err
    
    t = peiljaar - Bouwjaar 'leeftijd
    If t = 0 Then 'nieuw
        c = 1
    Else
    
    p = t / ThLevensduur
        If t / ThLevensduur > 0.75 Then
            c = 3
        Else
            c = 1 + Application.WorksheetFunction.Log(1 - (t / ThLevensduur), 0.5)
            If c < 1 Then GoTo err 'Foutafvanging component met bouwjaar in de toekomst
        End If
    
    End If
    fThConditiescore = CDbl(Round(c, 0))
    
resumeexit:
    Exit Function

err:
    fThConditiescore = ""
    GoTo resumeexit

End Function

