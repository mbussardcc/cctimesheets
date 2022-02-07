Private Sub CopyNClearButton_Click()
' Copy
    Sheets("Timesheet").Copy after:=Sheets(Sheets.Count)
    On Error Resume Next
' Rename and the clear the working sheet
    Dim endDate As Variant
    Dim startSun As String
    endDate = Range("B21").Value
    startSun = Range("C2").Value
    startSun = Replace(startSun, "/", "")
    ActiveSheet.Name = startSun
    On Error GoTo 0
' Remove button from backup sheet
    ActiveSheet.Shapes("CopyNClearButton").Delete
' Switch back to working sheet
    Worksheets("Timesheet").Activate
' Clear checkboxes in active shet
    Dim o As Object
        For Each o In ActiveSheet.OLEObjects
            If InStr(1, o.Name, "CheckBox") > 0 Then
                o.Object.Value = False
        End If
    Next
' Clear time entry data
    Application.GoTo Reference:="DataEntry"
    Selection.ClearContents
' Advance to next pay period
    Range("C2").Value = endDate + 1
End Sub