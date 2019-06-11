Attribute VB_Name = "CitaviMovePlaceholders"
Sub CitaviMovePlaceholdersToFootnote()

    Application.ScreenUpdating = False

    Dim selectedContentControls As ContentControls
    Dim selectedContentControlsCount As Long
    Dim oCC As ContentControl
    Dim insertionPointRange As Range
    Dim newFn As Footnote
    Dim newFnRange As Range
    Dim newCC As ContentControl
    
    Set selectedContentControls = Selection.Range.ContentControls
    selectedContentControlsCount = selectedContentControls.Count
    
    For contentControlIndex = 1 To selectedContentControlsCount
    
        Set oCC = selectedContentControls(contentControlIndex)
        Set insertionPointRange = Application.activeDocument.Range(selectedContentControls(contentControlIndex).Range.End + 1, selectedContentControls(contentControlIndex).Range.End + 1)

        insertionPointRange.Text = ""
    
        With insertionPointRange.FootnoteOptions
            .Location = wdBottomOfPage
            .NumberingRule = wdRestartContinuous
            .StartingNumber = 1
            .NumberStyle = wdNoteNumberStyleArabic
            .LayoutColumns = 0
        End With
        
        Set newFn = Application.activeDocument.Footnotes.Add(insertionPointRange, "")
       
        Set newField = newFn.Range.Fields.Add(newFn.Range, oCC.Range.Fields(1).Type)
        Set newCC = Application.activeDocument.ContentControls.Add(oCC.Type, newFn.Range)
        oCC.Range.Copy
        newCC.Range.Paste
    
    Next
    
    For i = selectedContentControlsCount To 1 Step -1
        selectedContentControls(i).Delete
    Next
    
    
    Application.ScreenUpdating = True

End Sub
