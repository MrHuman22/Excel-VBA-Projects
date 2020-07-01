'Uses the event handler Worksheet_Change
'For this reason, store this as a worksheet-specific sub
'Code adapted from Sumit Bansal from https://trumpexcel.com
'Allows multiple selections from data validation list. Repeat selection removes value from list.
Sub Worksheet_Change(ByVal Target As Range)

    Dim dict As Object
    Dim Newvalue As String
    Dim Oldvalue As String
    Dim previousPeople() As String
    Set dict = CreateObject("Scripting.Dictionary")

    Application.EnableEvents = True
    On Error GoTo Exitsub
    If Target.Column = 19 Then
        If Target.SpecialCells(xlCellTypeAllValidation) Is Nothing Then
            GoTo Exitsub
        Else
            If Target.Value = "" Then
                GoTo Exitsub
            Else
                Application.EnableEvents = False
                Newvalue = Target.Value
                Application.Undo
                Oldvalue = Target.Value
                'Debug.Print "New Value: " + Newvalue, "Oldvalue: " + Oldvalue

                ' If the box is empty, then just fill it with the selected name
                If Oldvalue = "" Then
                    'Debug.Print "Box was empty"
                    Target.Value = Newvalue
                    'Debug.Print "Target.Value = " + Target.Value + " and we're done"
                Else
                    'Debug.Print "Box has names in it already."
                    previousPeople = Split(Oldvalue, ", ")

                    'Adding the previous people to the dictionary
                    For Each Person In previousPeople
                        dict.Add Person, 1
                    Next Person

                    'Debugging to make sure the dictionary is constructed properly
                    'Debug.Print "Printing keys: "
                    'For Each Key In dict.Keys
                        'Debug.Print Key, dict(Key)
                    'Next Key

                    'If it's already in the list, delete it
                    'Debug.Print "Now adding or subtracting"
                    If dict.Exists(Newvalue) Then
                        dict.Remove (Newvalue)
                        'Debug.Print Newvalue + "is already listed. Deleting."
                        'Debug.Print "Updated keys: "
                        'For Each Key In dict.Keys
                            'Debug.Print Key, dict(Key)
                        'Next Key
                        Target.Value = Join(dict.Keys, ", ")
                        'Debug.Print "Setting Target.Value to: " + Target.Value
                    'If it's NOT already in the list, add it
                    Else
                        dict.Add Newvalue, 1
                        'Debug.Print Newvalue + "is NOT already listed. Adding to the list."
                        Target.Value = Join(dict.Keys, ", ")
                        'Debug.Print "Setting Target.Value to: " + Target.Value
                    End If
                End If
            End If
        End If
    End If
    
    Application.EnableEvents = True
    
    Exitsub:
        Application.EnableEvents = True

End Sub
