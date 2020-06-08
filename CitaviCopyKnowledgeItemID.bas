Attribute VB_Name = "CitaviCopyKnowledgeItemID"
Sub CitaviCopyKnowledgeItemID()

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
    Dim CCtrl As Field
    
    Dim arr() As String
    ReDim Preserve arr(0)
    
    Set selectedContentControls = Selection.ContentControls
    
    selectedContentControlsCount = selectedContentControls.Count
    
    If selectedContentControlsCount < 1 Then
        If Selection.Range.ParentContentControl Is Nothing Then Exit Sub
        For Each Field In Selection.Range.ParentContentControl.Range.Fields
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
                
                For Each s In regExp.Execute(formattedPlaceholder)
                i = UBound(arr) + 1
                ReDim Preserve arr(i)
                arr(i) = s.SubMatches.Item(1)
                Next
            End If
        Next
    Else
        For contentControlIndex = 1 To selectedContentControlsCount
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
                    
                    For Each s In regExp.Execute(formattedPlaceholder)
                    i = UBound(arr) + 1
                    ReDim Preserve arr(i)
                    arr(i) = s.SubMatches.Item(1)
                    Next
                End If
            Next
        Next
    End If
        
    CopyTextToClipboard Join(arr, vbCrLf)

End Sub

Private Function CopyTextToClipboard(txt)
    
    Dim obj As New DataObject
    obj.SetText txt
    obj.PutInClipboard

End Function
