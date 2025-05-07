Linkesíti a wordbe a linkeket (JÁP, Commitatus)

1. Megnyitod a word-ot
2. alt+f11
3. insert -> new module
4. aztán beszúrod ezt és futtatod:

Sub LinkesitMindenURL()
    Dim rgx As Object
    Dim myRange As Range
    Dim matches As Object
    Dim match As Object
    Dim cleanURL As String
    Dim startPos As Long
  
    Set rgx = CreateObject("VBScript.RegExp")
    rgx.Pattern = "(https?://[^\s)]+)"
    rgx.Global = True
    
 
    Set myRange = ActiveDocument.Range
    
    If rgx.Test(myRange.Text) Then
        Set matches = rgx.Execute(myRange.Text)
        
        startPos = 0
        
        For Each match In matches
            cleanURL = match.Value
            
            
            Do While Right(cleanURL, 1) = "." Or Right(cleanURL, 1) = ")"
                cleanURL = Left(cleanURL, Len(cleanURL) - 1)
            Loop
            
            
            Set myRange = ActiveDocument.Range(startPos, ActiveDocument.Content.End)
            
            
            With myRange.Find
                .Text = match.Value
                .Forward = True
                .Wrap = wdFindStop
                .Execute
                
                If .Found Then
                    Set myRange = ActiveDocument.Range(Start:=.Parent.Start, End:=.Parent.Start + Len(cleanURL))
                    ActiveDocument.Hyperlinks.Add Anchor:=myRange, Address:=cleanURL
                    
                    startPos = .Parent.Start + Len(cleanURL)
                End If
            End With
        Next
    End If
    
    Set rgx = Nothing
End Sub
