VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Proj_DataTier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjutil As C_GEN_MyUtilities
Private mstrDataSourceLocation As String
Private i As Integer
Private j As Integer
Private mWs As Worksheet

Public Function initiate(strWBSourceDataLocation As String)
    Set mobjutil = New C_GEN_MyUtilities
    
    mstrDataSourceLocation = strWBSourceDataLocation
    
    Dim wb As Workbook

    Set wb = Workbooks.Open(mstrDataSourceLocation)
        wb.Worksheets("Raw RSP Report (User List)").Cells(1, 1).Value = "Name"
        wb.Save
        wb.Close
    Set wb = Nothing
    
    
End Function
Public Function demolish()
    Set mobjutil = Nothing
    Set mWs = Nothing
End Function


Public Function giveMeDistinctTrainingName() As Object

    Dim objAdo As Object
'    Dim objColumnName As Object
    
    
    
    'Select distinct Mobile From [Table1$] Group By distinct Mobile
    
    'Name    Mobile  Email   Location    City    Zone    State   Retailer    Retailer Type   Retailer Type 2 Contract Number Agency  Last login  Points  Manage Account  Agency Manager  Corporate Manager
    
    
    Set objAdo = mobjutil.ADO_GiveMeRecordsetFromExternalWorkbook(mstrDataSourceLocation, "SELECT distinct [Training attended] FROM [Raw Training Taken Report$]")
    
    Set giveMeDistinctTrainingName = objAdo
    
    Set objAdo = Nothing
    
End Function

'Public Function giveMeDistinctTrainingName() As Object
'
'    Dim objAdo As Object
'
'    Set objAdo = mobjUtil.ADO_GiveMeRecordsetFromExternalWorkbook("D:\workingFile\Gopal\working\Raw_TRN_TKN_-1st_Oct_to_12th_Oct-15.xlsx", "SELECT distinct Mobile FROM [Raw RSP Report (User List)$] Group By Name,Mobile,Email,Location,City,Zone,State,Retailer,[Retailer Type],[Retailer Type 2],[Contract Number],Agency, [Last login],Points,[Manage Account],[Agency Manager],[Corporate Manager]")
'
'    Set giveMeDistinctTrainingName = objAdo
'
'    Set mobjUtil = Nothing
'
'    Set objAdo = Nothing
'
'End Function
'

'Public Function giveMeAllValue(wsSource As Worksheet)
'
''    Set mWs = wsSource
''
''    For i = 1 To mobjUtil.WB_GetLastUsedRowNumber(mWs)
''
''        For j = 1 To mobjUtil.WB_GetLastUsedColumnNumber(mWs)
''             mWs.Cells(i, j).Value = ThisWorkbook.Worksheets(5).Cells(i, j)
''
''        Next j
''    Next i
'
'    Dim objAdo As Object
''    Dim objColumnName As Object
'
'
'
'    'Select distinct Mobile From [Table1$] Group By distinct Mobile
'
'    'Name    Mobile  Email   Location    City    Zone    State   Retailer    Retailer Type   Retailer Type 2 Contract Number Agency  Last login  Points  Manage Account  Agency Manager  Corporate Manager
'
'
'    Set objAdo = mobjUtil.ADO_GiveMeRecordsetFromExternalWorkbook(mstrDataSourceLocation, "SELECT distinct [Training attended] FROM [Raw Training Taken Report$]")
'
'    Set giveMeDistinctTrainingName = objAdo
'
'    Set objAdo = Nothing
'
'
'End Function


Public Function giveMeTraining() As Object

    Dim objTraining As Object
        
    Set objTraining = mobjutil.ADO_GiveMeRecordsetFromExternalWorkbook(mstrDataSourceLocation, "SELECT Name, Mobile,[Training attended],Percentage FROM [Raw Training Taken Report$]")
    
    Set giveMeTraining = objTraining
    
    'Set giveMeTraining = Nothing
    
End Function


Public Function giveMeStudentDetails() As Object

    Dim objDistinctValue As Object
        
    Set objDistinctValue = mobjutil.ADO_GiveMeRecordsetFromExternalWorkbook(mstrDataSourceLocation, "SELECT * FROM [Raw RSP Report (User List)$] Order By Name, Mobile")
    
    Set giveMeStudentDetails = objDistinctValue
    
    Set objDistinctValue = Nothing
    
End Function
'

