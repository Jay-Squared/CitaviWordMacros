Attribute VB_Name = "CitaviFieldHelperFunctions"

Function RedirectReferences(myString As String)
    
    Dim objRegExp As regExp
    Set objRegExp = New regExp
    
    Dim objMatch As Match
    Dim colMatches As matchCollection
    
    Dim strProjectID As String
      
    objRegExp.Pattern = "(" & Chr(34) & "Project" & Chr(34) & ": {)([\s]+)(" & Chr(34) & "\$id" & Chr(34) & ": " & Chr(34) & ")([0-9]+)(" & Chr(34) & ")"
    objRegExp.MultiLine = False
    objRegExp.IgnoreCase = True
    objRegExp.Global = False
    
    Set colMatches = objRegExp.Execute(myString)
    strProjectID = colMatches(0).SubMatches(3)
    
        With objRegExp
           .Pattern = "(" & Chr(34) & "Project" & Chr(34) & ": {)([\s]+)(" & Chr(34) & "\$ref" & Chr(34) & ": " & Chr(34) & ")([0-9]+)(" & Chr(34) & ")"
           .Global = False
           .IgnoreCase = True
           .MultiLine = False
           Do While .Test(myString)
               Set objRegMC = .Execute(myString)
               For Each objRegM In objRegMC
                    myString = .Replace(myString, "$1$2" & Chr(34) & "$newref" & Chr(34) & ":  " & Chr(34) & strProjectID & "$5")
               Next
           Loop
    End With
    
    With objRegExp
           .Pattern = "\$newref"
           .Global = False
           .IgnoreCase = True
           .MultiLine = False
           Do While .Test(myString)
               Set objRegMC = .Execute(myString)
               For Each objRegM In objRegMC
                    myString = .Replace(myString, "$ref")
               Next
           Loop
    End With
    
    RedirectReferences = myString
End Function

Function RemoveEscapedLineBreaks(myString As String)
    
    Dim objRegExp As regExp
    Set objRegExp = New regExp

    With objRegExp
           .Pattern = "(^.?|(?!\\)..)(\\r|\\n)"
           .Global = False
           .IgnoreCase = True
           .MultiLine = False
           Do While .Test(myString)
               Set objRegMC = .Execute(myString)
               For Each objRegM In objRegMC
                    myString = .Replace(myString, "$1")
               Next
           Loop
    End With
    
    RemoveEscapedLineBreaks = myString

End Function

Function RenumberIDs(myString As String)

    Dim objRegExp As regExp
    
    Set objRegExp = New regExp
    
    i = 1

    With objRegExp
           .Pattern = "(" & Chr(34) & "\$id" & Chr(34) & ": " & Chr(34) & ")([0-9]+)(" & Chr(34) & ")"
           .Global = False
           .IgnoreCase = True
           .MultiLine = False
           Do While .Test(myString)
               Set objRegMC = .Execute(myString)
               For Each objRegM In objRegMC
                    myString = .Replace(myString, Chr(34) & "$uniqueid" & Chr(34) & ": " & Chr(34) & i & "$3")
                    i = i + 1
               Next
           Loop
    End With
    
    With objRegExp
           .Pattern = "\$uniqueid"
           .Global = False
           .IgnoreCase = True
           .MultiLine = False
           Do While .Test(myString)
               Set objRegMC = .Execute(myString)
               For Each objRegM In objRegMC
                    myString = .Replace(myString, "$id")
               Next
           Loop
    End With
    
    RenumberIDs = myString
   
End Function

Function ExtractPlaceholderBeginning(myString As String)

    Dim objRegExp As regExp
    Set objRegExp = New regExp
    
    Dim objMatch As Match
    Dim colMatches As matchCollection
      
    objRegExp.Pattern = "^(\s*)([\S\s]+?)(\s*)(\n    {)"
    objRegExp.MultiLine = False
    objRegExp.IgnoreCase = True
    objRegExp.Global = False
    
    Set colMatches = objRegExp.Execute(myString)
    
    ExtractPlaceholderBeginning = colMatches(0).SubMatches(0) & colMatches(0).SubMatches(1) & colMatches(0).SubMatches(2)
   
End Function

Function ExtractPlaceholderEnd(myString As String)

    Dim objRegExp As regExp
    Set objRegExp = New regExp
    
    Dim objMatch As Match
    Dim colMatches As matchCollection
      
    objRegExp.Pattern = "(\n)(  ],)(\s*)([\S\s]+?)(\s*)$"
    objRegExp.MultiLine = False
    objRegExp.IgnoreCase = True
    objRegExp.Global = False
    
    Set colMatches = objRegExp.Execute(myString)
    
    ExtractPlaceholderEnd = colMatches(0).SubMatches(0) & colMatches(0).SubMatches(1) & colMatches(0).SubMatches(2) & colMatches(0).SubMatches(3)
   
    ExtractPlaceholderEnd = colMatches(0).SubMatches(0) & colMatches(0).SubMatches(1) & colMatches(0).SubMatches(2) & colMatches(0).SubMatches(3)
   
End Function

Function ExtractPlaceholderEntries(myString As String)
   
    Dim objRegExp As regExp
    Dim objMatch As Match
    Dim colMatches   As matchCollection
    Dim RetStr As String
    
    Set objRegExp = New regExp
    
    objRegExp.Pattern = "(\n    {)\s*([\S\s]+?)\s*(\n    })"
    objRegExp.MultiLine = True
    objRegExp.IgnoreCase = True
    objRegExp.Global = True

    Set colMatches = objRegExp.Execute(myString)
    
    i = 1
        
    For Each objMatch In colMatches
        If i < colMatches.Count Then
            RetStr = RetStr & objMatch.Value & ","
        Else
            RetStr = RetStr & objMatch.Value
        End If
        i = i + 1
    Next
    
    ExtractPlaceholderEntries = RetStr
   
End Function

Function DecodeBase64(ByVal strData As String) As Byte()
 
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
    
    Set objNode = Nothing
    Set objXML = Nothing
 
End Function

Function encodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument
    
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    encodeBase64 = objNode.Text
 
    Set objNode = Nothing
    Set objXML = Nothing
End Function


'modified from http://needtec.sakura.ne.jp/codeviewer/index.php?id=2
'https://gist.github.com/yadimon/ce1d04b88de17064bfae

Function UnescapeUTF8(ByVal StringToDecode As String) As String
    Dim i As Long
    Dim acode As Integer, sTmp As String
    
    On Error Resume Next
    
    If InStr(1, StringToDecode, "\") = 0 And InStr(1, StringToDecode, "%") = 0 Then
        UnescapeUTF8 = StringToDecode
        Exit Function
    End If
    For i = Len(StringToDecode) To 1 Step -1
        acode = Asc(Mid$(StringToDecode, i, 1))
        Select Case acode
        Case 48 To 57, 65 To 90, 97 To 122
            ' don't touch alphanumeric chars
            DoEvents

        Case 92, 37: ' Decode \ or % value with uXXXX format
            If Mid$(StringToDecode, i + 1, 1) = "u" Then
                sTmp = CStr(CLng("&H" & Mid$(StringToDecode, i + 2, 4)))
                If IsNumeric(sTmp) Then
                    StringToDecode = Left$(StringToDecode, i - 1) & ChrW$(CInt("&H" & Mid$(StringToDecode, i + 2, 4))) & Mid$(StringToDecode, i + 6)
                End If
            End If
            
        Case 37: ' % not %uXXXX but %XX format
            
            sTmp = CStr(CLng("&H" & Mid$(StringToDecode, i + 1, 2)))
            If IsNumeric(sTmp) Then
                StringToDecode = Left$(StringToDecode, i - 1) & ChrW$(CInt("&H" & Mid$(StringToDecode, i + 1, 2))) & Mid$(StringToDecode, i + 3)
            End If
            
        End Select
    Next

    UnescapeUTF8 = StringToDecode
End Function
