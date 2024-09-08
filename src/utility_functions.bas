Attribute VB_Name = "utility_functions"
Option Explicit


Private Function printB2bPdfUsingAdobeSdkPageRange(ByVal filePath As String, ByVal startPage As Integer, ByVal endPage As Integer) As Variant
'    print provided range of page & return boolean if file name exist on any pages
    Dim acroApp As acroApp
    Dim avDoc As AcroAVDoc
    Dim pdDoc As AcroPDDoc
    
    Set acroApp = New acroApp
    Set avDoc = New AcroAVDoc
    Set pdDoc = New AcroPDDoc
    
    Dim resultArr As Variant
    ReDim resultArr(1 To 1, 1 To 2)
    
    Dim fileName As String
    Dim isFileNameExist As Boolean
    Dim methodeReturn As Variant
    
    methodeReturn = acroApp.Hide() ' this methode must call bfore call "Exit()" methode

    If avDoc.Open(filePath, "") Then
        
        Set pdDoc = avDoc.GetPDDoc()
        fileName = Left$(pdDoc.GetFileName, Len(pdDoc.GetFileName) - 4)
        isFileNameExist = avDoc.FindText(fileName, 0, 1, 1) ' find the file name
        
        methodeReturn = avDoc.PrintPagesSilent(startPage, endPage, 2, 0, 0)
        
    End If
    
'    methodeReturn = acroApp.CloseAllDocs() ' this methode must call bfore call "Exit()" methode
    
'    methodeReturn = acroApp.Exit()

    resultArr(1, 1) = fileName
    
    If isFileNameExist Then
        resultArr(1, 2) = "OK"
    Else
        resultArr(1, 2) = "Mismatch"
    End If
    
   printB2bPdfUsingAdobeSdkPageRange = resultArr
    

End Function

Private Function getBtlPiPageNum(ByVal strFilename As String) As Integer
    Dim acroApp As New acroApp
    Dim objAVDoc As New AcroAVDoc
    Dim objPDDoc As New AcroPDDoc
    Dim objPage As AcroPDPage
    Dim objSelection As AcroPDTextSelect
    Dim objHighlight As AcroHiliteList
    Dim pageNum As Long
      
    Dim methodeReturn As Variant
    methodeReturn = acroApp.Hide() ' this methode must call bfore call "Exit()" methode

    If objAVDoc.Open(strFilename, "") Then
    
       Set objPDDoc = objAVDoc.GetPDDoc
       
       For pageNum = 0 To objPDDoc.GetNumPages() - 1
       
          Set objPage = objPDDoc.AcquirePage(pageNum)
          Set objHighlight = New AcroHiliteList
          objHighlight.Add 0, 10000 ' Adjust this up if it's not getting all the text on the page
          Set objSelection = objPage.CreatePageHilite(objHighlight)
            
            If objSelection Is Nothing Then ' if page have no selected text, that's mean page is scan page
            
                getBtlPiPageNum = pageNum
                
'                methodeReturn = acroApp.CloseAllDocs() ' this methode must call bfore call "Exit()" methode
'                methodeReturn = acroApp.Exit()
                
                Exit Function
                
            End If
            
       Next pageNum
       
'       objAVDoc.Close 1
              
    End If
    
'    methodeReturn = acroApp.CloseAllDocs() ' this methode must call bfore call "Exit()" methode
'    methodeReturn = acroApp.Exit()
    
    If 3 <= objPDDoc.GetNumPages() - 1 Then
    
        getBtlPiPageNum = 3
        
    Else
    
        getBtlPiPageNum = 0 ' set first page if there is no text pages
        
    End If
    
 End Function
 
Private Function returnSelectedFilesFullPathArr(ByVal initialPath As String) As Variant
    Dim fileDialog As Object
    Dim selectedFiles As Variant
    Dim i As Long
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select Files"
        .AllowMultiSelect = True
         .InitialFileName = initialPath
        If .Show = -1 Then
            ReDim selectedFiles(1 To .SelectedItems.Count)
            For i = 1 To .SelectedItems.Count
                selectedFiles(i) = .SelectedItems.Item(i)
            Next i
        End If
    End With

    returnSelectedFilesFullPathArr = selectedFiles
End Function


Private Function getTextFromPDF(ByVal strFilename As String) As String
    Dim objAVDoc As New AcroAVDoc
    Dim objPDDoc As New AcroPDDoc
    Dim objPage As AcroPDPage
    Dim objSelection As AcroPDTextSelect
    Dim objHighlight As AcroHiliteList
    Dim pageNum As Long
    Dim strText As String
 
    strText = ""
    If objAVDoc.Open(strFilename, "") Then
       Set objPDDoc = objAVDoc.GetPDDoc
       For pageNum = 0 To objPDDoc.GetNumPages() - 1
          Set objPage = objPDDoc.AcquirePage(pageNum)
          Set objHighlight = New AcroHiliteList
          objHighlight.Add 0, 10000 ' Adjust this up if it's not getting all the text on the page
          Set objSelection = objPage.CreatePageHilite(objHighlight)
            Dim tCount As Variant
          If Not objSelection Is Nothing Then
             For tCount = 0 To objSelection.GetNumText - 1
                strText = strText & objSelection.GetText(tCount)
                Debug.Print strText
             Next tCount
          End If
       Next pageNum
       objAVDoc.Close 1
    End If
 
    getTextFromPDF = strText
 
 End Function
 

Private Function printPdfUsingAdobeSdkPageRange(ByVal filePath As String, ByVal startPage As Integer, ByVal endPage As Integer) As Variant
    '    print provided range of page
        Dim acroApp As acroApp
        Dim avDoc As AcroAVDoc

        Set acroApp = New acroApp
        Set avDoc = New AcroAVDoc
        Dim methodeReturn As Variant

        methodeReturn = acroApp.Hide() ' this methode must call bfore call "Exit()" methode

        If avDoc.Open(filePath, "") Then

            methodeReturn = avDoc.PrintPagesSilent(startPage, endPage, 2, 0, 0)

        End If

    '    methodeReturn = acroApp.CloseAllDocs() ' this methode must call bfore call "Exit()" methode

    '    methodeReturn = acroApp.Exit()

       printPdfUsingAdobeSdkPageRange = methodeReturn


End Function
