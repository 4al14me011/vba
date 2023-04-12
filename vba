Sub FindReplaceInPDF()

    ' Declare variables
    Dim AcroApp As Acrobat.CAcroApp
    Dim AcroAVDoc As Acrobat.CAcroAVDoc
    Dim AcroPDDoc As Acrobat.CAcroPDDoc
    Dim AcroTextSelect As Acrobat.CAcroPDTextSelect
    Dim AcroRect As Acrobat.CAcroRect
    Dim AcroPoint As Acrobat.CAcroPoint
    Dim SearchString As String
    Dim ReplaceString As String
    Dim i As Integer
    
    ' Set the search and replace strings
    SearchString = "old text"
    ReplaceString = "new text"
    
    ' Create a new instance of Acrobat
    Set AcroApp = CreateObject("AcroExch.App")
    
    ' Open the PDF file you want to work with
    Set AcroAVDoc = CreateObject("AcroExch.AVDoc")
    If AcroAVDoc.Open("C:\Path\To\Your\Document.pdf", "") Then
        Set AcroPDDoc = AcroAVDoc.GetPDDoc()
        
        ' Search for the text in the PDF file
        Set AcroTextSelect = AcroPDDoc.CreateTextSelect(0, 0)
        If AcroTextSelect.Find(SearchString, False, False) Then
            ' Loop through all instances of the text in the PDF file
            Do While AcroTextSelect.FindNext()
                ' Highlight the text
                Set AcroRect = AcroTextSelect.GetBoundingRect()
                Set AcroPoint = AcroRect.BottomRight
                AcroRect.Left = AcroRect.Left + 1
                AcroRect.Right = AcroRect.Right + 1
                AcroTextSelect.AddHighlightEx AcroRect, 1
                
                ' Replace the text
                AcroTextSelect.Replace(ReplaceString)
            Loop
        End If
        
        ' Clean up
        AcroPDDoc.Save 1, ""
        AcroAVDoc.Close True
    End If
    AcroApp.Exit
    
End Sub
