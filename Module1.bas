Attribute VB_Name = "Module1"
Option Explicit

Public myTaskFile As String
Public lNumbTasks  As Long
Public iWhichRec As Integer
Public strLastSort As String

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
  "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias _
  "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
  ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString _
  As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub InitVars()
    Dim strTemp As String, strArray() As String, ii As Integer
    myTaskFile = App.Path & "\Tasks.txt"
    
    strTemp = ReadINI_String("ValidStatus", "Groups", myTaskFile)
    SeparateString strArray, strTemp, ","
    For ii = 0 To UBound(strArray)
        frmMain.cboStatus.AddItem strArray(ii)
    Next
    strTemp = ReadINI_String("ValidPriority", "Groups", myTaskFile)
    SeparateString strArray, strTemp, ","
    For ii = 0 To UBound(strArray)
        frmMain.cboPriority.AddItem strArray(ii)
    Next
    strTemp = ReadINI_String("ValidRequest", "Groups", myTaskFile)
    SeparateString strArray, strTemp, ","
    For ii = 0 To UBound(strArray)
        frmMain.cboRequest.AddItem strArray(ii)
    Next
    
    strLastSort = "0:ASC"
    GetWindowPosition
End Sub

'Writes the specified value to the Ini file
Public Function WriteINI(strAppName As String, strKeyname As String, strValue As String, strININame As String)
    WriteINI = WritePrivateProfileString(strAppName, strKeyname, strValue, strININame)
End Function

  
'Reads the specified Ini string value
Public Function ReadINI_String(strAppName As String, strKeyname As String, strININame As String, Optional varDefault) As String
    Const c_intLargest As Integer = 1023
    
    Dim strReturn     As String
    Dim intSizeReturn As Integer
    Dim strDefault    As String
    
    If IsMissing(varDefault) Then
        strDefault = ""
    Else
        strDefault = CStr(varDefault)
    End If
    
    strReturn = Space$(c_intLargest)
        
    'Read from the INI File
    intSizeReturn = GetPrivateProfileString(strAppName, strKeyname, strDefault, strReturn, c_intLargest, strININame)
    
    'trim Read value and removed NULL
    strReturn = Trim$(Left$(strReturn, intSizeReturn))
    
    ReadINI_String = strReturn
End Function

Function DeleteINI_Section(ByVal INIFileLoc As String, ByVal Section As String)
    WritePrivateProfileString Section, vbNullString, vbNullString, INIFileLoc
End Function

Public Sub SeparateString(strResult() As String, strWorking As String, strDelimiter As String)
    'Cuts one string into an array of smaller strings according to specified delimiter.
    'The separated strings are placed into an array.

    Dim lngLastLoc As Long
    Dim lngDelimLoc As Long
    Dim lngFieldCount As Long

    'clean out string
    ReDim strResult(0 To 0) As String

    If strWorking <> "" Then
        'initialize last location to 0
        lngLastLoc = 1
        lngFieldCount = 0
        lngDelimLoc = InStr(lngLastLoc, strWorking, strDelimiter)

        'Continue until no more delimiters are found
        Do Until lngDelimLoc = 0
            'Set array size
            If lngFieldCount > 0 Then
                ReDim Preserve strResult(0 To lngFieldCount) As String
            End If

            'Set string and locations to find next delim
            strResult(lngFieldCount) = Trim(Mid$(strWorking, lngLastLoc, lngDelimLoc - lngLastLoc))
            lngLastLoc = lngDelimLoc + Len(strDelimiter)
            lngDelimLoc = InStr(lngLastLoc, strWorking, strDelimiter)
            lngFieldCount = lngFieldCount + 1
        Loop

        'add last field (after delim)
        If lngFieldCount > 0 Then
            ReDim Preserve strResult(0 To lngFieldCount) As String
        End If
        strResult(lngFieldCount) = Trim(Mid$(strWorking, lngLastLoc))
    End If
End Sub

Public Sub LoadData(frm As Form, ByRef myGrid As MSFlexGrid)
    ' This sub is used to load the data from the definitions in the Tasks.txt file
    Dim TaskName As String, jj As Integer

    myGrid.Cols = 7
    myGrid.TextMatrix(0, 1) = "Status"
    myGrid.TextMatrix(0, 2) = "Priority"
    myGrid.TextMatrix(0, 3) = "Description"
    myGrid.TextMatrix(0, 4) = "Due Date"
    myGrid.TextMatrix(0, 5) = "Requested By"
    myGrid.TextMatrix(0, 6) = "Comment"
    
    lNumbTasks = CLng(ReadINI_String("TaskList", "Num", myTaskFile))
    If lNumbTasks > 0 Then
        For jj = 1 To lNumbTasks
            TaskName = "TASK" & jj
            myGrid.AddItem TaskName
            myGrid.TextMatrix(jj, 1) = ReadINI_String(TaskName, "Status", myTaskFile)
            myGrid.TextMatrix(jj, 2) = ReadINI_String(TaskName, "Priority", myTaskFile)
            myGrid.TextMatrix(jj, 3) = ReadINI_String(TaskName, "Description", myTaskFile)
            myGrid.TextMatrix(jj, 4) = ReadINI_String(TaskName, "DueDate", myTaskFile)
            myGrid.TextMatrix(jj, 5) = ReadINI_String(TaskName, "RequestedBy", myTaskFile)
            myGrid.TextMatrix(jj, 6) = ReadINI_String(TaskName, "Comment", myTaskFile)
            myGrid.Row = jj
            HighPriorityColor myGrid
        Next jj
        AutosizeGridColumns myGrid, 15, 2000, frm
    End If
End Sub

Public Sub SaveData(frm As Form, ByRef myGrid As MSFlexGrid)
    'This write data from the grid back to the Tasks.txt file.
    Dim jj As Integer, TaskName As String
    
    WriteINI "TaskList", "Num", CStr(lNumbTasks), myTaskFile
    For jj = 1 To lNumbTasks
        WriteINI myGrid.TextMatrix(jj, 0), "Status", myGrid.TextMatrix(jj, 1), myTaskFile
        WriteINI myGrid.TextMatrix(jj, 0), "Priority", myGrid.TextMatrix(jj, 2), myTaskFile
        WriteINI myGrid.TextMatrix(jj, 0), "Description", myGrid.TextMatrix(jj, 3), myTaskFile
        WriteINI myGrid.TextMatrix(jj, 0), "DueDate", myGrid.TextMatrix(jj, 4), myTaskFile
        WriteINI myGrid.TextMatrix(jj, 0), "RequestedBy", myGrid.TextMatrix(jj, 5), myTaskFile
        WriteINI myGrid.TextMatrix(jj, 0), "Comment", myGrid.TextMatrix(jj, 6), myTaskFile
    Next jj
End Sub

Public Sub RenumberTasks(ByRef myGrid As MSFlexGrid)
    Dim jj As Integer
    For jj = 1 To lNumbTasks
        myGrid.TextMatrix(jj, 0) = "TASK" & jj
    Next jj
End Sub

Public Sub LoadRequestedBy(frm As Form, ByRef myGrid As MSFlexGrid)
    Dim jj As Integer, kk As Integer
    Dim temp As String, found As Boolean
    
    For jj = 1 To lNumbTasks
        found = False
        temp = myGrid.TextMatrix(jj, 5)
        For kk = 1 To frmMain.cboRequest.ListCount
            If frmMain.cboRequest.List(kk) = temp Then
                found = True
            End If
        Next
        If found = False Then
            frmMain.cboRequest.AddItem temp
        End If
    Next
End Sub

Public Sub AutosizeGridColumns(ByRef msFG As MSFlexGrid, ByVal MaxRowsToParse As Integer, ByVal MaxColWidth As Integer, frm As Form)
    ' Auto resize flexgrid column widths. ***CUSTOM VERSION OF ROUTINE FOR THIS APP ***
    Dim I, J As Integer
    Dim txtString As String
    Dim intTempWidth, intBiggestWidth As Integer
    Dim intRows As Integer
    Const intPadding = 150

    With msFG
        .ColWidth(0) = 0 'HARD CODE first Column to width = 0 (HIDDEN)
        For I = 1 To .Cols - 1
            ' Loops through every column
            .Col = I
            ' Set the active colunm
            intRows = .Rows
            ' Set the number of rows
            If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
            ' If there are more rows of data, reset intRows to the MaxRowsToParse constant
            
            intBiggestWidth = 0
            ' Reset some values to 0

            For J = 0 To intRows - 1
                ' check up to MaxRowsToParse # of rows and obtain the greatest width of the cell contents
                .Row = J
                
                txtString = .Text
                intTempWidth = frm.TextWidth(txtString) + intPadding
                ' The intPadding constant compensates for text insets. You can adjust this value above as desired.
                
                If intTempWidth > intBiggestWidth Then intBiggestWidth = intTempWidth
                ' Reset intBiggestWidth to the intMaxCol Width value if necessary
            Next J
            .ColWidth(I) = intBiggestWidth
        Next I
        
        ' Now check to see if the columns aren't as wide as the grid itself.
        ' If not, expand the last column to fill the grid
        intTempWidth = 0

        For I = 1 To .Cols - 1
            intTempWidth = intTempWidth + .ColWidth(I)
        Next I

        If intTempWidth < msFG.Width Then
            .ColWidth(6) = .ColWidth(6) + intTempWidth
        End If

        .Col = 1
        .Row = 1
    End With
End Sub

Public Sub GridDeleteRow(ByRef myGrid As MSFlexGrid)
    Dim ii As Integer
    ' In order to delete the "last record" we have to change .FixedRows to zero
    ii = myGrid.Row
    If myGrid.Rows = 2 Then
        myGrid.FixedRows = 0
        myGrid.Row = ii
    End If
    If myGrid.Row > 0 Then
        myGrid.RemoveItem myGrid.Row
    End If
End Sub

Public Sub SortGrid(ByRef myGrid As MSFlexGrid, SortCol As Long)
    Dim oldSort() As String, jj As Integer
    
    'Remove any sort arrows icons first....
    myGrid.Row = 0
    For jj = 1 To 6
        myGrid.Col = jj
        Set frmMain.MSFlexGrid1.CellPicture = Nothing
    Next
    
    SeparateString oldSort, strLastSort, ":"
    myGrid.Col = SortCol
    If CLng(oldSort(0)) <> SortCol Then
        myGrid.Sort = flexSortGenericAscending
        strLastSort = CStr(SortCol) & ":ASC"
        Set frmMain.MSFlexGrid1.CellPicture = LoadPicture(App.Path & "\up.bmp")
        frmMain.MSFlexGrid1.CellPictureAlignment = flexAlignRightCenter
    Else
        If oldSort(1) = "ASC" Then
            myGrid.Sort = flexSortGenericDescending
            strLastSort = CStr(SortCol) & ":DES"
            Set frmMain.MSFlexGrid1.CellPicture = LoadPicture(App.Path & "\down.bmp")
            frmMain.MSFlexGrid1.CellPictureAlignment = flexAlignRightCenter
        Else
            myGrid.Sort = flexSortGenericAscending
            strLastSort = CStr(SortCol) & ":ASC"
            Set frmMain.MSFlexGrid1.CellPicture = LoadPicture(App.Path & "\up.bmp")
            frmMain.MSFlexGrid1.CellPictureAlignment = flexAlignRightCenter
        End If
    End If
End Sub

Public Sub FilterReport(ByRef myGrid As MSFlexGrid)
    Dim strPri As String, strSta As String, strReq As String
    Dim jj As Integer, kk As Integer
    
    strPri = frmMain.cboPriority.Text
    strSta = frmMain.cboStatus.Text
    strReq = frmMain.cboRequest.Text
    
    For jj = 1 To lNumbTasks ' Temporarily, make ALL rows visible
        myGrid.RowHeight(jj) = -1  ' -1 resets to default height
    Next
    
    If frmMain.chkHideComplete = Checked Then
        For jj = 1 To lNumbTasks
            If myGrid.TextMatrix(jj, 1) = "Completed" Then
                myGrid.RowHeight(jj) = 0 ' Hide Row
            End If
        Next
    End If
    For jj = 1 To lNumbTasks
        If strPri <> "(All)" And strPri <> myGrid.TextMatrix(jj, 2) Then
            myGrid.RowHeight(jj) = 0 ' Hide row
        End If
    Next
    For jj = 1 To lNumbTasks
        If strSta <> "(All)" And strSta <> myGrid.TextMatrix(jj, 1) Then
            myGrid.RowHeight(jj) = 0 ' Hide row
        End If
    Next
    For jj = 1 To lNumbTasks
        If strReq <> "(All)" And strReq <> myGrid.TextMatrix(jj, 5) Then
            myGrid.RowHeight(jj) = 0 ' Hide row
        End If
    Next
End Sub

Public Sub HighPriorityColor(ByRef myGrid As MSFlexGrid)
    Dim jj As Integer, myColor As Long
    
    If myGrid.TextMatrix(myGrid.Row, 2) = "High" Then
        myColor = vbRed
    ElseIf myGrid.TextMatrix(myGrid.Row, 2) = "Medium" Then
        myColor = vbBlue
    Else
        myColor = vbBlack
    End If
    
    For jj = 1 To 6
        myGrid.Col = jj
        myGrid.CellForeColor = myColor
    Next
End Sub

Public Sub SaveWindowPosition()
    WriteINI "Position", "Left", frmMain.Left, myTaskFile
    WriteINI "Position", "Top", frmMain.Top, myTaskFile
    WriteINI "Position", "Width", frmMain.Width, myTaskFile
    WriteINI "Position", "Height", frmMain.Height, myTaskFile
End Sub

'Gets the last saved position of the specified window.
Public Sub GetWindowPosition()
    Dim strLeft     As String
    Dim strTop      As String
    Dim strHeight   As String
    Dim strWidth    As String
    Dim intScreenWid As Integer
    Dim intScreenHei As Integer
    
    strLeft = ReadINI_String("Position", "Left", myTaskFile)
    If strLeft <> "" Then
        frmMain.Left = CLng(strLeft)
    End If
    
    strTop = ReadINI_String("Position", "Top", myTaskFile)
    If strTop <> "" Then
        frmMain.Top = CLng(strTop)
    End If
    
    strHeight = ReadINI_String("Position", "Height", myTaskFile)
    If strHeight <> "" Then
        frmMain.Height = CInt(strHeight)
    End If
    
    strWidth = ReadINI_String("Position", "Width", myTaskFile)
    If strWidth <> "" Then
        frmMain.Width = CInt(strWidth)
    End If
    
    'Make sure form is smaller than screen....
    intScreenWid = Screen.Width
    intScreenHei = Screen.Height
    If frmMain.Left < 0 Or frmMain.Left > 10000 Or frmMain.Top < 0 Or frmMain.Top > 10000 Then
        frmMain.Width = 11000
        frmMain.Left = 2800
        frmMain.Height = 8200
        frmMain.Top = 2200
    End If
        
End Sub

