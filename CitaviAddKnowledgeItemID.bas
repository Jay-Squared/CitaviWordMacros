Attribute VB_Name = "CitaviAddKnowledgeItemID"
Sub CitaviAddKnowledgeItemIDToSelectedPlaceholder()
    
    ' Variables

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
    
    Dim KnowledgeItemId
    
    Dim Json As Object
    
    Dim i As Integer
    Dim j As Integer
    
    Dim regExp As Object
    Set regExp = CreateObject("vbscript.regexp")
    
    Dim objData As New MSForms.DataObject
    
    Dim selectedContentControls As ContentControls
    
    Set selectedContentControls = Selection.ContentControls
    selectedContentControlsCount = selectedContentControls.Count
    
    If selectedContentControlsCount > 1 Then Exit Sub
    
    objData.GetFromClipboard
    
    replacementPattern = "$1$2$3$4""AssociateWithKnowledgeItemId"": """ & objData.GetText() & """,$4"

    For contentControlIndex = 1 To selectedContentControlsCount
        
        i = 1
        j = 1
        
        formattedPlaceholders = ""
    
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
                                
                With regExp
                    .Pattern = "(\""Id\"": "")([0-9a-z\-]+)("",)([\s\n]+)"
                    .Global = True
                    formattedPlaceholder = .Replace(formattedPlaceholder, replacementPattern)
                End With
                                
                formattedPlaceholders = formattedPlaceholders & formattedPlaceholder
                
                If i = 1 Then
                    firstFormattedPlaceholderBeginning = ExtractPlaceholderBeginning(formattedPlaceholder)
                    firstFormattedPlaceholderEnd = ExtractPlaceholderEnd(formattedPlaceholder)
                End If
                 
                j = j + 1
                i = i + 1
                
            End If
        Next
        
        newFormattedPlaceholderEntries = ExtractPlaceholderEntries(formattedPlaceholders)
        newFormattedPlaceholder = firstFormattedPlaceholderBeginning & newFormattedPlaceholderEntries & vbCr & firstFormattedPlaceholderEnd
        newFormattedAndNumberedPlaceholder = RenumberIDs(newFormattedPlaceholder)
        newPlaceholder = Replace(JsonConverter.ConvertToJson(JsonConverter.ParseJson(newFormattedAndNumberedPlaceholder), Whitespace:=0), vbCrLf, "")
        newPlaceholder = UnescapeUTF8(newPlaceholder)
        newPlaceholder = RemoveEscapedLineBreaks(newPlaceholder)
        newEncodedPlaceholder = encodeBase64(StrConv(newPlaceholder, vbFromUnicode))
        newFieldCode = Replace("ADDIN CitaviPlaceholder{" & newEncodedPlaceholder & "}", vbLf, "")
        
        selectedContentControls(contentControlIndex).Range.Fields(1).Code.Text = newFieldCode
        
    Next

End Sub
