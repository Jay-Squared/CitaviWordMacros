Attribute VB_Name = "CitaviMergePlaceholders"
Sub CitaviMergePlaceholders()
    Application.ScreenUpdating = False
    
    ' Variables
    
    Dim selectedContentControls As ContentControls
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
    
    Dim i As Integer
    Dim j As Integer
    
    Dim regExp As Object
    Set regExp = CreateObject("vbscript.regexp")
    
    Set selectedContentControls = Selection.ContentControls
    selectedContentControlsCount = selectedContentControls.Count
    
    i = 1

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
                
                formattedPlaceholders = formattedPlaceholders & formattedPlaceholder
                
                If i = 1 Then
                    firstFormattedPlaceholderBeginning = ExtractPlaceholderBeginning(formattedPlaceholder)
                    firstFormattedPlaceholderEnd = ExtractPlaceholderEnd(formattedPlaceholder)
                End If
                
                If j > 1 Then selectedContentControls(contentControlIndex).Range.Fields(j).Delete
    
                j = j + 1
                i = i + 1
            End If
        Next
        
    Next
    
    If Not Trim(formattedPlaceholders & vbNullString) = vbNullString Then
    
        newFormattedPlaceholderEntries = ExtractPlaceholderEntries(formattedPlaceholders)
        newFormattedPlaceholder = firstFormattedPlaceholderBeginning & newFormattedPlaceholderEntries & vbCr & firstFormattedPlaceholderEnd
        newFormattedAndNumberedPlaceholder = RenumberIDs(newFormattedPlaceholder)
        newFormattedAndNumberedPlaceholder = RedirectReferences(newFormattedAndNumberedPlaceholder)
        newPlaceholder = Replace(JsonConverter.ConvertToJson(JsonConverter.ParseJson(newFormattedAndNumberedPlaceholder), Whitespace:=0), vbCrLf, "")
        newPlaceholder = UnescapeUTF8(newPlaceholder)
        newPlaceholder = RemoveEscapedLineBreaks(newPlaceholder)
        newEncodedPlaceholder = encodeBase64(StrConv(newPlaceholder, vbFromUnicode))
        newFieldCode = Replace("ADDIN CitaviPlaceholder{" & newEncodedPlaceholder & "}", vbLf, "")
        
        i = 1
        
        For i = selectedContentControlsCount To 1 Step -1
            If i > 1 Then
                selectedContentControls(i).Delete
            End If
        Next
        
        selectedContentControls(1).Range.Fields(1).Code.Text = newFieldCode
        
    End If
    
    Application.ScreenUpdating = True
    
End Sub
