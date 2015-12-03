Attribute VB_Name = "Module1"
Option Explicit
Private mobjutil2 As C_GEN_MyUtilities

Sub startProcessing()
    
    Set mobjutil2 = New C_GEN_MyUtilities
    
    Dim strSourceDataFileLocation As String
    strSourceDataFileLocation = mobjutil2.WB_OpenFile("Please Select the Raw Datasource File.", XLSX)
    
    If strSourceDataFileLocation = "False" Then GoTo AllDone
    
    Dim strWBName As String
    strWBName = mobjutil2.STR_GiveMeAllLettersAfterLast("\", strSourceDataFileLocation)
    
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name = strWBName Then
            MsgBox "The Source Data Workbook you chose (or another workbook with the same name)is already open. Please close it before proceeding.", vbOKOnly + vbInformation, "Cannot Proceed"
            wb.Activate
            Set wb = Nothing
            GoTo AllDone
        End If
    Next wb
    
    Set wb = Nothing
    
    
    Dim objDT As C_Proj_DataTier
    Dim objPR As C_Proj_Presentation
    
    Set objDT = New C_Proj_DataTier
    Set objPR = New C_Proj_Presentation
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Dashboard"))
    
    
    
    
    
    objDT.initiate strSourceDataFileLocation
    objPR.initiate ws
    
    
    Dim objRS As Object
    
    Set objRS = objDT.giveMeDistinctTrainingName
    objPR.writeTrainingNames objRS
    Set objRS = Nothing
    
    Dim objRs1 As Object
    
    Set objRs1 = objDT.giveMeStudentDetails
    objPR.writeStudentDetails objRs1
    Set objRs1 = Nothing
    
    Dim objRs2 As Object
    Set objRs2 = objDT.giveMeTraining
    
    Debug.Print objRs2.RecordCount
    
    objPR.getPercentageumber objRs2
    
    Set objRs2 = Nothing
    
AllDone:
    On Error Resume Next
    objDT.demolish
    objPR.demolish
    Set mobjutil2 = Nothing
    
End Sub

Private Function testSecond()

End Function


Private Function getFileLocation() As String

    Dim strPath As String
        strPath = Application.GetOpenFilename("Text Files (*.xlsx), *.xlsx, Add-in Files (*.xla), *.xla", 2, "Open My Files", , True)
        
        If strPath <> "" Then
            Debug.Print strPath
        Else
            MsgBox "Slect FIle"
        End If
        
        
    Dim wb As Workbook
    
        Set wb = Workbooks.Open(strPath)
            wb.Worksheets(1).Cells(1, 1) = "Name"
            wb.Save
            wb.Close
        Set wb = Nothing
        

End Function
