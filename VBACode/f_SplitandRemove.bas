Attribute VB_Name = "F_SplitandRemove"
Function Custom_SplitandRemove(InitialRange As String, RemoveMatchesFromThisRange As String) As String

'Troubleshooters
On Error GoTo Troubleshooter1
If InitialRange = Empty Or InitialRange = "" Then
    GoTo Troubleshooter_EmptyFirstValue
End If
If RemoveMatchesFromThisRange = Empty Or Trim(RemoveMatchesFromThisRange) = "" Then
    GoTo Troubleshooter_EmptySecondValue
End If

'Set Initial Variables
Z = 0
b = 0
n = 0
Dim NewArray() As String
ReDim NewArray(0 To 5000)

'Split Strings
SplitValuesInitialRange = Split(InitialRange, ";")
SplitValuesSecondaryRange = Split(RemoveMatchesFromThisRange, ";")

'Loop Through Strings, Remove Duplicates Entirely
For Each InitialValue In SplitValuesInitialRange
    If WorksheetFunction.Trim(InitialValue) <> "" Then
        For i = 0 To UBound(SplitValuesSecondaryRange)
            If WorksheetFunction.Proper(WorksheetFunction.Trim(InitialValue)) = WorksheetFunction.Proper(WorksheetFunction.Trim(SplitValuesSecondaryRange(i))) Then
                TheresAMatch = True
            End If
        Next i
        If TheresAMatch = False Then
            NewArray(n) = WorksheetFunction.Trim(InitialValue) & "; "
            n = n + 1
        End If
        TheresAMatch = Empty
    End If
Next InitialValue
If n > 0 Then
    ReDim Preserve NewArray(0 To n - 1)
End If

'Bring Array Into String
For b = 0 To n - 1
NewString = NewArray(b) & NewString
Next b

'Reset Variables
b = 0
n = 0
Erase NewArray

'Set Function for Cell Input
If NewString = Empty Then
    Custom_SplitandRemove = ""
    Else
    Custom_SplitandRemove = NewString
End If
NewString = Empty

Exit Function

'Troubleshooters
Troubleshooter1:
MsgBox ("Something went wrong" & vbCrLf & Err.Description & vbCrLf & Err.Number & vbCrLf & Err.Source & vbCrLf & Error & vbCrLf & vbObjectError & vbCrLf & InitialAddress & " " & RemoveMatchesAddress)
Exit Function

Troubleshooter_EmptyFirstValue:
Custom_SplitandRemove = ""
Exit Function

Troubleshooter_EmptySecondValue:
Custom_SplitandRemove = Trim(InitialRange)
Exit Function

End Function

