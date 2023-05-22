Option Explicit
Function Clipboard$(Optional s$)
    Dim v: v = s  'Cast to variant for 64-bit VBA support
    With CreateObject("htmlfile")
    With .parentWindow.clipboardData
        Select Case True
            Case Len(s): .setData "text", v
            Case Else:   Clipboard = .GetData("text")
        End Select
    End With
    End With
End Function

Sub timesheet()

Dim coll As String
Dim items As Object
Set items = CreateObject("System.Collections.ArrayList")

items.Add "moderating_solutions: "
items.Add "communication: "
items.Add "testing_phases_number: "
items.Add "testing_phases_hours: "
items.Add "projects: "


Dim i As Integer

For i = 3 To 7
    If Range("O" & CStr(i)).Value <> 0 Then
        coll = coll + (items(i - 3) + Replace(CStr(Range("o" + CStr(i)).Value), ",", ".") + ";" + vbCr)
    End If
Next i

coll = coll + "comments: "

If Range("P8").Value <> 0 Then
    coll = coll + CStr(Range("P8").Value) + " solutions fixed manually, "
End If
If Range("O8").Value <> 0 Then
    coll = coll + "resolved " + CStr(Range("O8").Value) + " flagged solutions, "
End If
If Range("Q8").Value <> 0 Then
    coll = coll + "submitted " + CStr(Range("Q8").Value) + " feedback forms, "
End If
coll = Left(coll, Len(coll) - 2) + ";"

Range("Q3").Value = Time

Clipboard coll

End Sub
