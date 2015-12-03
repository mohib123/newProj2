VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_GEN_MyUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'INTRODUCTION:
'This Class is a common Class where the most commonly required functions (utilities)
'are included as methods.

'This class is not project-specific, rather it holds very common functionalities that are
'regularly used in any project coded by me

'The methods here are classified into 4 main categories based on
'what they deal with

'Categories Include:
'1)String --->[function names start with STR_]
'2)Workbooks/Worksheets/Folder/File-path  --->[function names start with WB_]
'3)Arrays  --->[function names start with ARR_]
'4)Recordsets(ADODB) --->[function names start with ADO_]


'AUTHOR:
'Raad Ahmad


Enum FileType
    CSV = 1
    XLSX = 2
    XLS = 3
    TXT = 4
    ALL = 5
    ALLEXCELS = 6
End Enum
Public Function WB_DeleteBlankRows(ws As Worksheet)
    ws.Activate
    ws.Cells(1, 1).Activate
    ws.UsedRange.Select
    Dim lngCounter As Long
    For lngCounter = Selection.Rows.Count To 1 Step -1
        If WorksheetFunction.CountA(Selection.Rows(lngCounter)) = 0 Then
            Selection.Rows(lngCounter).EntireRow.Delete
        End If
    Next lngCounter
    
    ws.Cells(1, 1).Activate
    ws.Cells(1, 1).Select
    
End Function

Public Function WB_IsWorkBookOpen(FileNameWithFullPath As String) As Boolean
    
    'checks if a workbook is already open or not
    
    Dim lngFF As Long
    Dim lngErrNo As Long

    On Error Resume Next
    
    lngFF = FreeFile()
    Open FileNameWithFullPath For Input Lock Read As #lngFF
    Close lngFF
    lngErrNo = Err
    
    On Error GoTo 0

    Select Case lngErrNo
        Case 0:    WB_IsWorkBookOpen = False
        Case 70:   WB_IsWorkBookOpen = True
        Case Else: Error lngErrNo
    End Select
    
End Function

Function WB_FileExists(FullPath As String) As Boolean
    
    If Trim(FullPath) = "" Then GoTo errF
    
    If Dir(FullPath) <> vbNullString Then
        WB_FileExists = True
    Else
        GoTo errF
    End If
    
Exit Function
errF:
    WB_FileExists = False
End Function

Public Function WB_OpenFile(ByVal BoxTitle As String, ByVal FilterType As FileType) As Variant

    Dim strFilterType As String
    Dim vntFilePathGotten As Variant

    Select Case FilterType
        Case CSV
            vntFilePathGotten = Application.GetOpenFilename(title:=BoxTitle, FileFilter:="csv Files (*.csv), *.csv")
        Case XLSX
            vntFilePathGotten = Application.GetOpenFilename(title:=BoxTitle, FileFilter:="Excel Workbooks (*.xlsx), *.xlsx")
        Case XLS
            vntFilePathGotten = Application.GetOpenFilename(title:=BoxTitle, FileFilter:="Excel Workbooks (*.xls), *.xls")
        Case TXT
            vntFilePathGotten = Application.GetOpenFilename(title:=BoxTitle, FileFilter:="Text Files (*.txt), *.txt")
        Case ALL
            vntFilePathGotten = Application.GetOpenFilename(title:=BoxTitle)
        Case ALLEXCELS
            vntFilePathGotten = Application.GetOpenFilename(title:=BoxTitle, FileFilter:="Excel Workbooks (*.xls*), *.xls*")
        Case Else
            Exit Function
    End Select

    WB_OpenFile = vntFilePathGotten

End Function

Public Function WB_GetOpenFolderName(BoxTitle As String) As String
    
    Set myDlg = Application.FileDialog(msoFileDialogFolderPicker)
    myDlg.AllowMultiSelect = False
    myDlg.title = BoxTitle
    
    If myDlg.Show = -1 Then 'folder is chosen
        WB_GetOpenFolderName = myDlg.SelectedItems(1) & "\"
    Else
        WB_GetOpenFolderName = "False"
    End If
    
    Set myDlg = Nothing
    
End Function

Public Function WB_GetLastUsedColumnNumber(ws As Worksheet) As Long
    
    WB_GetLastUsedColumnNumber = 0
    
    On Error Resume Next
    
    WB_GetLastUsedColumnNumber = ws.Cells.Find(What:="*", _
        After:=ws.Range("A1"), _
        Lookat:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Column
        
    On Error GoTo 0
    
End Function
Public Function WB_TrimColumnHeaders(ws As Worksheet)
    
    Dim lngTtlSearch As Long
    lngTtlSearch = WB_GetLastUsedColumnNumber(ws)
    Dim i As Integer
    
    For i = 1 To lngTtlSearch
        ws.Cells(1, i).Value = Trim(CStr(ws.Cells(1, i)))
    Next i
    
End Function

Public Function WB_GetSelectionRangeAddress() As String
    WB_GetSelectionRangeAddress = Selection.Address(False, False)
End Function

Public Function WB_GetColumnAlphabetName(ws As Worksheet, colNum As Long) As String
    WB_GetColumnAlphabetName = CStr(Split(ws.Cells(1, colNum).Address, "$")(1))
End Function

Public Function WB_GetColumnNumberFromHeader(ws As Worksheet, colHeaderTitle As String) As Long

    Dim intTtlColumns As Integer
    Dim blnColumnFound As Boolean
    Dim intCounter As Integer
    Dim lngHolder As Long
    
    
    intCounter = 0
    lngHolder = 0
    
    intTtlColumns = WB_GetLastUsedColumnNumber(ws)
      
         Do
            If (UCase(Trim(ws.Cells(1, intCounter + 1).FormulaR1C1)) = UCase(Trim(colHeaderTitle))) Then
                lngHolder = intCounter + 1
                Exit Do
            End If
            intCounter = intCounter + 1
        Loop Until intCounter = intTtlColumns + 2
    
    WB_GetColumnNumberFromHeader = lngHolder
    
End Function
Public Function WB_GetLastUsedRowNumberInAColumn(colNum As Long, ws As Worksheet) As Long
    
    Dim strColHead As String
    strColHead = CStr(Split(ws.Cells(1, colNum).Address, "$")(1))
    
    Err.Clear
    On Error Resume Next
    WB_GetLastUsedRowNumberInAColumn = ws.Columns(strColHead + ":" + strColHead).Find("*", , , , xlByRows, xlPrevious).Row
    If Err.Number = 91 Then WB_GetLastUsedRowNumberInAColumn = 0
    Err.Clear
    
End Function

Public Function WB_GetLastUsedRowNumber(ws As Worksheet) As Long

    On Error GoTo errHandler
    
    If ActiveWorkbook Is Nothing Then
        WB_GetLastUsedRowNumber = 0
        Exit Function
    ElseIf WB_IsSheetEmpty(ws) Then
        WB_GetLastUsedRowNumber = 0
        Exit Function
    Else
        With ws
            WB_GetLastUsedRowNumber = .Cells.Find(What:="*", _
            After:=.Range("A1"), _
            Lookat:=xlPart, _
            LookIn:=xlFormulas, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False).Row
        End With
        
        'the old version:=
        'WB_GetLastUsedRowNumber = ws.Cells.Find("*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
    End If
    
    
    
       
    
    
    
    Exit Function
    
errHandler:
    WB_GetLastUsedRowNumber = 0

End Function

Public Function WB_IsSheetEmpty(ws As Worksheet) As Boolean

    If ws Is Nothing Then
        WB_IsSheetEmpty = True
    Else
        If (Application.WorksheetFunction.CountA(ws.Cells) = 0) Then
            WB_IsSheetEmpty = True
        Else
            WB_IsSheetEmpty = False
        End If
    End If
        
End Function

Public Function STR_GiveMeAllLettersAfterLast(LastWhat As String, FullString As String) As String

    Dim intCounter As Integer
    Dim strHolder As String
    
    intCounter = Len(FullString)
    strHolder = ""
    If intCounter = 0 Then Exit Function
    
    Do
        If mid(FullString, intCounter, 1) = LastWhat Then
            Exit Do
        Else
            strHolder = mid(FullString, intCounter, 1) + strHolder
        End If
        intCounter = intCounter - 1
    
    Loop Until intCounter = 0
    
    STR_GiveMeAllLettersAfterLast = strHolder

End Function


Public Function STR_SplitAndAffixStringAfterEveryParamReturnArray(ByVal SourceString As String, SplitParam As String, AddPrefix As String, AddPostfix As String) As Variant

    Dim intPos1 As Integer
    
    Dim strMyArray() As String
    Dim i As Integer
    i = 0

    Dim strTempHold As String
    Dim strFullHold As String
    
    strTempHold = ""
    strFullHold = ""
    
    Dim intTtlLength As Integer
    intTtlLength = Len(SourceString)
    Dim intCounter As Integer
    intCounter = 1
    
    Do
        'On Error Resume Next
        
        intPos1 = InStr(intCounter, SourceString, SplitParam)
        
        If intPos1 = 0 Then
            strTempHold = Trim(mid(SourceString, intCounter, intTtlLength - intCounter + 1))
        Else
            strTempHold = Trim(mid(SourceString, intCounter, intPos1 - intCounter))
        End If
        
        intCounter = intPos1 + 1
        
        If Trim(strTempHold) <> "" Then
            ReDim Preserve strMyArray(i)
            strMyArray(i) = Trim(AddPrefix + strTempHold + AddPostfix)
            i = i + 1
        End If
        
    
    Loop Until intPos1 = 0
    
    STR_SplitAndAffixStringAfterEveryParamReturnArray = strMyArray

End Function

Public Function STR_ToUppercaseFirstLetters(ByVal SourceString As String) As String
    
    Dim lngTtlLength As Long
    lngTtlLength = Len(SourceString)
    
    
    SourceString = LCase(Trim(SourceString))
    SourceString = Replace(Replace(SourceString, "  ", " "), "   ", " ")
    SourceString = UCase(Left(SourceString, 1)) + Right(SourceString, Len(SourceString) - 1)
    
    Dim blnSpaceFound As Boolean
    blnSpaceFound = False
    
    Dim lngCounter As Long
    lngCounter = 1
    
    Dim strHolder As String
    
    Do
        If blnSpaceFound Then
            strHolder = strHolder + UCase(mid(SourceString, lngCounter, 1))
            blnSpaceFound = False
        Else
            strHolder = strHolder + mid(SourceString, lngCounter, 1)
        End If
        If mid(SourceString, lngCounter, 1) = " " Then blnSpaceFound = True
        'shall catch the blnResult in the nextiteration
        
        lngCounter = lngCounter + 1
    Loop Until lngCounter = lngTtlLength + 1
    
    STR_ToUppercaseFirstLetters = strHolder

End Function

Public Function STR_GiveMeAllLettersAfterFirst(ByVal FirstWhat As String, ByVal FullString As String) As String
    
    Dim strHold As String
    FullString = Trim(FullString)
    strHold = Right(FullString, Len(FullString) - InStr(1, FullString, FirstWhat))
    STR_GiveMeAllLettersAfterFirst = strHold

End Function



Function ADO_GiveMeRecordsetFromExternalWorkbook(ByVal ExcelFileFullPathName As String, ByVal FullSQLQuery As String) As Object
    
    Dim cnx As Object
    Dim cmd As Object
    Dim rst As Object
    
    
    'Late Binding to the ADODB objects
    Set cnx = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    Set rst = CreateObject("ADODB.Recordset")
    
    'Signifying the nature of the recordset object (rst)
    rst.CursorLocation = 3
    rst.CursorType = 2
    rst.LockType = 3
    
    'Cursor Location Enums
    'adUseClient            3 Uses client-side cursors supplied by a local cursor library. Local cursor services often will allow many features that driver-supplied cursors may not, so using this setting may provide an advantage with respect to features that will be enabled. For backward compatibility, the synonym adUseClientBatch is also supported.
    'adUseNone              1 Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.)
    'adUseServer            2 Default. Uses cursors supplied by the data provider or driver. These cursors are sometimes very flexible and allow for additional sensitivity to changes others make to the data source. However, some features of the The Microsoft Cursor Service for OLE DB, such as disassociated. Recordset objects, cannot be simulated with server-side cursors and these features will be unavailable with this setting.
    
    'Cursor Type Enums
    'adOpenDynamic          2 Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
    'adOpenForwardOnly      0 Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
    'adOpenKeyset           1 Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
    'adOpenStatic           3 Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
    'adOpenUnspecified      -1 Does not specify the type of cursor.
    
    'Lock Type Enums
    'adLockBatchOptimistic  4 ----> Indicates optimistic batch updates. Required for batch update mode.
    'adLockOptimistic       3 ----> Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method.
    'adLockPessimistic      2 ----> Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing.
    'adLockReadOnly         1 ----> Indicates read-only records. You cannot alter the data.
    'adLockUnspecified      -1 ----> Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.

    
    
    'opening the connection ...
    With cnx
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .connectionstring = "Data Source='" & ExcelFileFullPathName & "'; " & "Extended Properties='Excel 8.0;HDR=Yes'"
        .Open
    End With
    
    Set cmd.activeconnection = cnx
    'cmd.CommandText = "SELECT DISTINCT[Brand] FROM [" + strSheetName + "$]"
    cmd.CommandText = FullSQLQuery
    rst.Open cmd
    
    Dim strm As Object
    Set strm = CreateObject("ADODB.Stream")
    rst.Save strm
    rst.Close
    
    Set ADO_GiveMeRecordsetFromExternalWorkbook = CreateObject("ADODB.Recordset")
    ADO_GiveMeRecordsetFromExternalWorkbook.Open strm
    
    'Set ADO_GiveMeRecordsetFromExternalWorkbook = rst
    
    'now closing and nullifying all the objects used in this procedure
    
    Set strm = Nothing
    
    Set rst = Nothing
    
    
    Set cmd = Nothing
    
    cnx.Close
    Set cnx = Nothing
        
End Function


Function ADO_GiveMeRecordsetFromActiveWorkbook(ByVal FullSQLQuery As String) As Object
    'The FullSQLQuery that is passed to this function should be in the following format
    '"SELECT DISTINCT[Brand] FROM [" + SheetName + "$]" , or
    '"SELECT * FROM [" + SheetName + "$]"
    
    Dim cnx As Object
    Dim cmd As Object
    Dim rst As Object
    
    
    'Late Binding to the ADODB objects
    Set cnx = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    Set rst = CreateObject("ADODB.Recordset")
    
    'Signifying the nature of the recordset object (rst)
    rst.CursorLocation = 3
    rst.CursorType = 2
    rst.LockType = 3

    'Cursor Location Enums
    'adUseClient            3 Uses client-side cursors supplied by a local cursor library. Local cursor services often will allow many features that driver-supplied cursors may not, so using this setting may provide an advantage with respect to features that will be enabled. For backward compatibility, the synonym adUseClientBatch is also supported.
    'adUseNone              1 Does not use cursor services. (This constant is obsolete and appears solely for the sake of backward compatibility.)
    'adUseServer            2 Default. Uses cursors supplied by the data provider or driver. These cursors are sometimes very flexible and allow for additional sensitivity to changes others make to the data source. However, some features of the The Microsoft Cursor Service for OLE DB, such as disassociated. Recordset objects, cannot be simulated with server-side cursors and these features will be unavailable with this setting.
    
    'Cursor Type Enums
    'adOpenDynamic          2 Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
    'adOpenForwardOnly      0 Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
    'adOpenKeyset           1 Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
    'adOpenStatic           3 Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
    'adOpenUnspecified      -1 Does not specify the type of cursor.
    
    'Lock Type Enums
    'adLockBatchOptimistic  4 ----> Indicates optimistic batch updates. Required for batch update mode.
    'adLockOptimistic       3 ----> Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method.
    'adLockPessimistic      2 ----> Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing.
    'adLockReadOnly         1 ----> Indicates read-only records. You cannot alter the data.
    'adLockUnspecified      -1 ----> Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.

    
    'Now we are opening the connection
    With cnx
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .connectionstring = "Data Source='" & ThisWorkbook.path & "\" & ThisWorkbook.Name & "'; " & "Extended Properties='Excel 12.0;HDR=Yes'"
        .Open
    End With
    
    Set cmd.activeconnection = cnx
    cmd.CommandText = FullSQLQuery
    rst.Open cmd
    
    
    Dim strm As Object
    Set strm = CreateObject("ADODB.Stream")
    rst.Save strm
    
    rst.Close
    
    Set ADO_GiveMeRecordsetFromActiveWorkbook = CreateObject("ADODB.Recordset")
    ADO_GiveMeRecordsetFromActiveWorkbook.Open strm
    
    'now closing and nullifying all the objects used in this procedure
    
    Set strm = Nothing
    Set rst = Nothing
    
    
    Set cmd = Nothing
    
    cnx.Close
    Set cnx = Nothing
        
End Function

Public Function ARR_PutRangeDataIntoArray(ws As Worksheet, AddressOfRange As String) As Variant
    ARR_PutRangeDataIntoArray = ws.Range(AddressOfRange).Value
End Function

Public Function ARR_PutArrayDataIntoRange(ws As Worksheet, PasteInWhichCell As String, ByVal TheArray As Variant)
    
    ws.Range(PasteInWhichCell).Resize(UBound(TheArray, 1) + 1, UBound(TheArray, 2) + 1) = TheArray
    'UBound(TheArray, 2) : by placing '2' I am providing the dimension for which the upperbound is needed
    
End Function

Public Function ARR_IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    ARR_IsArrayAllocated = IsArray(arr) And Not IsError(LBound(arr, 1)) And LBound(arr, 1) <= UBound(arr, 1)
    On Error GoTo 0
End Function

Public Function ARR_Get2DArrayFromACollectionOfArray(ByVal arrCollection As Variant) As Variant
    
    Dim lngColumnCount As Long
    Dim lngRowCount As Long
    Dim vntInnerArray As Variant
    Dim vnt2DArray As Variant
    
    
    vntInnerArray = arrCollection(0)
    lngColumnCount = UBound(vntInnerArray)
    lngRowCount = UBound(arrCollection)
    
    ReDim vnt2DArray(lngRowCount, lngColumnCount) As Variant
    
    Dim i As Long
    Dim j As Long
    
    For i = 0 To lngRowCount
        vntInnerArray = arrCollection(i)
        For j = 0 To lngColumnCount
            vnt2DArray(i, j) = vntInnerArray(j)
        Next j
    Next i
    
    On Error Resume Next
    Erase arrCollection
    Erase vntInnerArray
    On Error GoTo 0
    
    ARR_Get2DArrayFromACollectionOfArray = vnt2DArray
    
End Function

Public Function WB_SortSheet(ws As Worksheet, ColumnNumberToSort As Long, HasHeader As Boolean)
    
    'First We Sort The sheet
    Dim TotalRecs As Long
    Dim LastColName As String
    TotalRecs = WB_GetLastUsedRowNumber(ws)
    LastColName = WB_GetColumnAlphabetName(ws, WB_GetLastUsedColumnNumber(ws))
    
    Dim strSortColName As String
    strSortColName = Trim(WB_GetColumnAlphabetName(ws, ColumnNumberToSort))
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range( _
        strSortColName + "1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ws.Sort
        .SetRange Range("A1:" + Trim(LastColName) + CStr(TotalRecs))
        If HasHeader = True Then
            .Header = xlYes
        Else
            .Header = xlNo
        End If
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'next we prepare the two arrays
    
End Function


