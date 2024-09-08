Attribute VB_Name = "Main"
Option Explicit

 Sub b2bPrint()
     
    Dim filePathArr As Variant
    filePathArr = Application.Run("utility_functions.returnSelectedFilesFullPathArr", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2024")  ' all b2b path
    
    Dim filePath As Variant
    Dim btlPiPageNum As Integer
    Dim fileNameExistReturnArr As Variant
    Dim resultArr As Variant
    ReDim resultArr(1 To UBound(filePathArr), 1 To 2)
    Dim i As Long
    
    
    Dim clipboard As Object
    Set clipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") ' create clipboard data object
        
    
    For i = LBound(filePathArr) To UBound(filePathArr)
        
        filePath = filePathArr(i)
        
        btlPiPageNum = Application.Run("utility_functions.getBtlPiPageNum", filePath)
                
        fileNameExistReturnArr = Application.Run("utility_functions.printB2bPdfUsingAdobeSdkPageRange", filePath, 0, btlPiPageNum)
        
        resultArr(i, 1) = fileNameExistReturnArr(1, 1)
        resultArr(i, 2) = fileNameExistReturnArr(1, 2)
        
        clipboard.SetText fileNameExistReturnArr(1, 1) 'clipboard text
        clipboard.PutInClipboard 'put to clipboard
        
    Next i
    
    Dim wsRange As Range
    Set wsRange = ActiveSheet.Range("A1").Resize(UBound(resultArr, 1), UBound(resultArr, 2))
    wsRange.NumberFormat = "@"
    wsRange.Value = resultArr
    
    Dim App As acroApp
    Dim methodeReturn As Variant
    Set App = New acroApp
    methodeReturn = App.Hide() ' this methode must call bfore call "Exit()" methode
    methodeReturn = App.CloseAllDocs() ' this methode must call bfore call "Exit()" methode
    methodeReturn = App.Exit()
    
    
    MsgBox "Printing Process Completed"
    
 End Sub


Sub printPdfWithPageRange()

    Dim filePathArr As Variant
    filePathArr = Application.Run("utility_functions.returnSelectedFilesFullPathArr", "")  'default path nothing

    Dim filePath As Variant
    Dim methodeReturn As Variant

    Dim startPageNumber, endPageNumber As Integer
    startPageNumber = InputBox("Enter Start Page Number", "Enter a Number") - 1
    endPageNumber = InputBox("Enter End Page Number", "Enter a Number") - 1

    Dim i As Long

    For i = LBound(filePathArr) To UBound(filePathArr)

        filePath = filePathArr(i)

        methodeReturn = Application.Run("utility_functions.printPdfUsingAdobeSdkPageRange", filePath, startPageNumber, endPageNumber) 'this function return a arr, that's why just received

    Next i

    Dim App As acroApp
    Set App = New acroApp
    methodeReturn = App.Hide() ' this methode must call bfore call "Exit()" methode
    methodeReturn = App.CloseAllDocs() ' this methode must call bfore call "Exit()" methode
    methodeReturn = App.Exit()

    MsgBox "Printing Process Completed"

 End Sub


