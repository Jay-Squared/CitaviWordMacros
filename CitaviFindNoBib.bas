Attribute VB_Name = "CitaviFindNoBib"
Sub FindFirstNoBibField()

    Dim selectedContentControlsCount As Long
    Dim contentControlIndex As Long
    
    Dim encodedPlaceholder As String
    Dim decodedPlaceholder As String
    Dim formattedPlaceholder As String
    Dim firstFormattedPlaceholderBeginning As String
    Dim firstFormattedPlaceholderEnd As String
    Dim formattedPlaceholders As String
    
    Dim newPlaceholderEntries As String
    Dim newFormattedPlaceholderEntries As String
    
    Dim newFormattedPlaceholder As String
    Dim newFormattedAndNumberedPlaceholder As String
    Dim newPlaceholder As String
    Dim newEncodedPlaceholder As String
    Dim newFieldCode As String
    
    Dim Json As Object
    
    Dim I As Integer
    Dim j As Integer
    
    Dim regExp As Object
    Set regExp = CreateObject("vbscript.regexp")
    
    Dim selectedContentControls As ContentControls

    Set selectedContentControls = Selection.ContentControls
    selectedContentControlsCount = selectedContentControls.Count

    For contentControlIndex = 1 To selectedContentControlsCount
    
        j = 1
        
        For Each Field In selectedContentControls(contentControlIndex).Range.Fields
        If Field.Code Like "*ADDIN CitaviPlaceholder*" Then
        
            With regExp
            .Pattern = "(ADDIN CitaviPlaceholder\{)([^\}]+)(\})"
            .Global = True
            encodedPlaceholder = .Replace(Field.Code, "$2")
            End With
            decodedPlaceholder = StrConv(DecodeBase64(encodedPlaceholder), vbUnicode)
            
            Set Json = JsonConverter.ParseJson(decodedPlaceholder)
            
            formattedPlaceholder = JsonConverter.ConvertToJson(Json, Whitespace:=2)
            formattedPlaceholder = UnescapeUTF8(formattedPlaceholder)
            
            CopyTextToClipboard (formattedPlaceholder)
            
            If InStr(1, formattedPlaceholder, """NoBib"": true,") Then
            
                Selection.SetRange Start:=selectedContentControls(contentControlIndex).Range.Start, _
                End:=selectedContentControls(contentControlIndex).Range.End
                Exit Sub
            End If
        
        
        
        End If
        Next
    
    Next

End Sub


