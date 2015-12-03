VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Proj_Presentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mWs As Worksheet
Private mobjutil As C_GEN_MyUtilities


Public Function writeTrainingNames(objRsSource As Object)
    
    Dim i As Integer
    i = 19
    objRsSource.movefirst
    Do While objRsSource.EOF = False
        
        mWs.Cells(1, i).Value = objRsSource.Fields(0).Value
        
        objRsSource.moveNext
        i = i + 1
    Loop

End Function

Public Function writeStudentDetails(objRsSource As Object)
    
    'Getting Header Name
    Dim z As Long
    For z = 0 To objRsSource.Fields.Count - 1
        mWs.Cells(1, z + 1).Value = objRsSource.Fields(z).Name
    Next z
    'End of Header Name
    
    Dim i As Integer
    i = 2
    Dim j As Integer

    objRsSource.movefirst
    
    Do While objRsSource.EOF = False
        
        If (UCase(Trim(mWs.Cells(i - 1, 1))) <> UCase(Trim(objRsSource.Fields(0).Value))) And UCase(Trim((mWs.Cells(i - 1, 2))) <> UCase(Trim(objRsSource.Fields(1).Value))) Then
            
            For z = 0 To objRsSource.Fields.Count - 1
                mWs.Cells(i, z + 1).Value = objRsSource.Fields(z).Value
            Next z
            i = i + 1
            
        End If
        
        objRsSource.moveNext
        
    Loop

    'mWs.Range("A2").CopyFromRecordset objRsSource

End Function


Public Function getPercentageumber(objRsSource As Object)
    
    Dim i As Integer
    i = 2
    Dim intColumnNumber As Integer
    
    Dim intCount As Integer
    'intCount = 0
    Dim lngTotalRecs As Long
    lngTotalRecs = mobjutil.WB_GetLastUsedRowNumber(mWs)
    
    
    Dim lngCourseQty As Long
    
    For i = 2 To lngTotalRecs
        
        lngCourseQty = 0
        objRsSource.Filter = "Name = '" & fixApostrophe(Trim(CStr(mWs.Cells(i, 1).Value))) & "' AND Mobile = " & Trim(CStr(mWs.Cells(i, 2).Value)) & ""
        
        If objRsSource.RecordCount > 0 Then
            
            objRsSource.movefirst
            
            Do While objRsSource.EOF = False
                
                On Error Resume Next
                
                intColumnNumber = mobjutil.WB_GetColumnNumberFromHeader(mWs, Trim(CStr(objRsSource.Fields(2).Value)))
                
                If intColumnNumber > 0 Then lngCourseQty = lngCourseQty + 1
                
                If Err.Number <> 0 Then
                    Debug.Print "Problem " + CStr(objRsSource.Fields(2).Value)
                    GoTo proceed:
                End If
                Err.Clear
                
                On Error GoTo 0
                
                mWs.Cells(i, intColumnNumber).Value = objRsSource.Fields(3)

proceed:
                
                objRsSource.moveNext
                'intCount = intCount + 1
            Loop
        End If
        
        objRsSource.Filter = 0
            'intCount = 0
        If lngCourseQty > 0 Then mWs.Range("BD" + CStr(i)).Value = lngCourseQty
        If lngCourseQty = 0 Then mWs.Range("BE" + CStr(i)).Value = "All rows are empty as no training has been taken by user - but, he is found in Raw RSP Report (User List)"
        
    Next i
    
    
    
    
    
'        Dim objRS As Object
'
'    'Assign : Name, Mobile, Training Attended ,Percentage
'
'
'    objRS.Filter "Name='Gopal Krishna' AND Mobile = 999999999"
'    '
'    'objRS.movefirst
'    'LOOP
'
'    objRS.Filter = 0


End Function

Private Function fixApostrophe(strSource As String) As String

    fixApostrophe = Replace(strSource, "'", "''")
    
End Function


Public Function demolish()
    
    mWs.Rows("1:1").Select
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
    
    mWs.Rows("1:1").Font.Bold = True
    
    mWs.Range("BD1").Value = "Total Trainings"
    
    mWs.UsedRange.EntireColumn.AutoFit
    
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In mWs.Parent.Worksheets
        If ws.Name <> "Output_New" And ws.Name <> "Dashboard" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True
    Set ws = Nothing
    Set mWs = Nothing
    
    
    Set mobjutil = Nothing
    
    
End Function
Public Function initiate(wsReportingSheet As Worksheet)
    
    
    
    Dim wrs As Worksheet
    Application.DisplayAlerts = False
    For Each wrs In wsReportingSheet.Parent.Worksheets
        If wrs.Name = "Output_New" Then wrs.Delete
    Next wrs
    Set wrs = Nothing
    Application.DisplayAlerts = True
    
    
    Set mWs = wsReportingSheet
    mWs.Name = "Output_New"
    Set mobjutil = New C_GEN_MyUtilities
    
    mWs.UsedRange.Clear
    
    
End Function

