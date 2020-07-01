Sub AddEA(ByVal Target As Range)
'Code adapted from Sumit Bansal from https://trumpexcel.com
'Allows multiple selections from data validation list. Repeat selection removes value from list.
'In this instance, the selections are always made from column 19.
Dim dict As Object
Dim Newvalue As String
Dim Oldvalue As String
Dim previousPeople() As String
Set dict = CreateObject("Scripting.Dictionary")

Application.EnableEvents = True
If Target.Column = 19 Then
    If Target.SpecialCells(xlCellTypeAllValidation) Is Nothing Then
    GoTo Exitsub
    Else: If Target.Value = "" Then GoTo Exitsub Else
    Application.EnableEvents = False
    Newvalue = Target.Value
    Debug.Print "New Value: " + Newvalue
    Application.Undo
    Oldvalue = Target.Value
    Debug.Print "Oldvalue: " + Oldvalue
        ' If the box is empty, then just fill it with the selected name
        If Oldvalue = "" Then
            Debug.Print "Box was empty"
            Target.Value = Newvalue
            Debug.Print "Target.Value = " + Target.Value + "and we're done"
        Else
            Debug.Print "Box has names in it already."
            previousPeople = Split(Oldvalue, ", ")
            For Each Person In previousPeople
                dict.Add Person, 1
            Next Person
            
            'If it's already in the list, delete it
            If dict.Exist(Target.Value) Then
                dict.Remove (Target.Value)
                Debug.Print Target.Value + "is already listed. Deleting."
            
            'If it's NOT already in the list, add it
            Else
                dict.Add Target.Value, 1
                Debug.Print Target.Value + "is NOT already listed. Adding."
            End If
            ' Reconstruct the string
            Target.Value = Join(dict.Keys, ", ")
            Debug.Print "Setting Target.Value to: " + Target.Value
           
        End If
    End If
End If
Application.EnableEvents = True
Exitsub:
Application.EnableEvents = True
End Sub
