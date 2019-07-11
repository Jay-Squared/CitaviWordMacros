Attribute VB_Name = "CitaviCopyKnowledgeItemID"
Sub CopyKnowledgeItemID()

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
    
    Dim selectedContentControls As ContentControls
    
    Set selectedContentControls = Selection.ContentControls
    selectedContentControlsCount = selectedContentControls.Count
    
    If selectedContentControlsCount > 1 Then Exit Sub

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
                
                With regExp
                    .Pattern = "(""AssociateWithKnowledgeItemId"": "")([0-9a-z\-]+)("")"
                    .Global = True

                End With
                
                Set allMatches = regExp.Execute(formattedPlaceholder)

                If allMatches.Count <> 0 Then
                    associatedKnowleddgeItemID = allMatches.Item(0).SubMatches.Item(1)
                End If
                
                CopyTextToClipboard (associatedKnowleddgeItemID)
            End If
        Next
    Next

End Sub

Private Function CopyTextToClipboard(txt)
    
    Dim obj As New DataObject
    obj.SetText txt
    obj.PutInClipboard

End Function
