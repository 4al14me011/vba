Sub FindReplaceInPDF()

    ' Declare variables
    Dim AcroApp As Acrobat.CAcroApp
    Dim AcroAVDoc As Acrobat.CAcroAVDoc
    Dim AcroPDDoc As Acrobat.CAcroPDDoc
    Dim AcroHiliteList As Acrobat.CAcroHiliteList
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
        If AcroTextSelect.FindText(SearchString, False, False) Then
            ' Loop through all instances of the text in the PDF file
            Do While AcroTextSelect.FindNext()
                ' Highlight the text
                Set AcroHiliteList = AcroTextSelect.GetHiliteList()
                For i = 0 To AcroHiliteList.GetCount() - 1
                    Set AcroRect = AcroHiliteList.GetRect(i)
                    Set AcroPoint = AcroRect.BottomRight
                    
                    ' Replace the text
                    AcroTextSelect.Replace(ReplaceString)
                Next
            Loop
        End If
        
        ' Clean up
        AcroPDDoc.Close
    End If
    AcroAVDoc.Close (True)
    AcroApp.Exit
    
End Sub
