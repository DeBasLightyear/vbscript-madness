Do
    Dim input
    input = InputBox("Enter normal text:", "I implore you to act", "Text goes here")

    If IsEmpty(input) Or input = "2" Then
        WScript.Quit()
    ElseIf input = "" Then
        MsgBox("No input.")
    End If
Loop Until input <> ""
