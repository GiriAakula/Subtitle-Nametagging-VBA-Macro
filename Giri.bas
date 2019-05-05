Attribute VB_Name = "Nametagging for SRT"
Sub Giri()
Dim first As String
Dim second As String
Dim third As String
Dim loopvalue As Integer
Dim myInputBoxVariable As String

    
    Selection.EndKey Unit:=wdStory
    Selection.MoveUp Unit:=wdLine, Count:=3
    first = Selection.Sentences(1)
    If IsNumeric(first) = True Then
    third = first
    Else
    Selection.MoveUp Unit:=wdLine, Count:=1
    second = Selection.Sentences(1)
    third = second
    End If
    Selection.HomeKey Unit:=wdStory
    myInputBoxVariable = InputBox(Prompt:="Enter text that you want to put at the beginning of each sentence", Title:="Giri", Default:="Eg: [Giri to Asma]")
    
    For loopvalue = 0 To third
    Call Asma(myInputBoxVariable)
    Next
    

End Sub



Sub Asma(ByVal inputText As String)
    
    Selection.MoveDown Unit:=wdLine, Count:=2
    Selection.TypeText Text:=inputText
    Selection.HomeKey Unit:=wdLine
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    
    If InStr(1, Selection.Range.Text, "A") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "B") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "C") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "D") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "E") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "F") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "G") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "H") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "I") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "J") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "K") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "L") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "M") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "N") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "O") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "P") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "Q") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "R") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "S") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "T") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "U") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "V") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "W") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "X") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "Y") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "Z") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "a") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "b") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "c") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "d") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "e") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "f") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "g") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "h") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "i") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "j") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "k") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "l") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "m") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "n") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "o") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "p") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "q") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "r") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "s") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "t") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "u") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "v") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "w") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "x") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "y") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "z") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "-") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, """") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    ElseIf InStr(1, Selection.Range.Text, "''") = 1 Then
    Selection.MoveDown Unit:=wdLine, Count:=2
    
    Else
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    End If
    Selection.HomeKey Unit:=wdLine
    End Sub

