Option Compare Database

Private objCurDB As DAO.Database

Public Function CurDb(Optional bolRefresh As Boolean = False) As DAO.Database
    If objCurDB Is Nothing Or bolRefresh = True Then
        Set objCurDB = CurrentDb()
    End If
 
    Set CurDb = objCurDB
End Function

Public Sub CurDBClear()
    Set objCurDB = Nothing
End Sub

' Called from external script (i.e. powershell, batch, etc.)
Public Sub UpdateDatabase()
    DbUpdate signalCompletion:=False
End Sub


Function TableClear(tableName As String)
    Dim clearSQL As String
    Set db = CurDb
    
    clearSQL = "DELETE [" & tableName & "].* " & _
                "FROM [" & tableName & "];"
    
    db.Execute clearSQL
End Function

Function ImportFilesExcel(tableName As String, path As String, fName As String, tabName As String)
    Dim i As Integer
    Dim file As String
    Dim str As String
    Set db = CurDb
    
    str = ""
    
    ' Turn off warnings
    DoCmd.SetWarnings False
    
    ' Search 'dir' for fName
    file = Dir(path & fName)
    i = 1
    Do While Len(file) > 0
        ImportFileExcel tableName, path & file, tabName
        str = str & file & vbCrLf
        file = Dir
        i = i + 1
    Loop
    
    ' Debug-Output imported file names
    MsgBox "Finished importing files to [" & tableName & "]: " & vbCrLf & str
    
    ' Turn warnings back on
    DoCmd.SetWarnings True
End Function


Function ImportFileExcel(tableName As String, fPath As String, tabName As String)
    On Error GoTo ErrEnter
    
    DoCmd.TransferSpreadsheet acImport, , tableName, fPath, True, tabName
    
ExitFunc:
    
    Exit Function

ErrEnter:
    
    Select Case Err.Number
        Case 0
            'no error
        Case 3011
            ' couldn't find file/tab
            MsgBox "The file """ & fPath & """ or tab """ & tabName & """ could not be found!" & vbCrLf & vbCrLf & _
                "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Source: " & Err.Source
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & "Source: " & Err.Source
    End Select
    Err.Clear
    Resume ExitFunc
End Function

Function ExportExcel(tblName As String, Optional exportPath As String, Optional excelFileName As String)
    On Error GoTo Err_Enter
    
    If exportPath = "" Then
        exportPath = CurrentProject.path & "\"
    End If
    If excelFileName = "" Then
        excelFileName = tblName & ".xlsx"
    End If
    
    DoCmd.TransferSpreadsheet _
        TransferType:=acExport, _
        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
        tableName:=tblName, _
        FileName:=exportPath & excelFileName, _
        HasFieldNames:=True
        
Err_Exit:
    On Error GoTo 0
    Exit Function
        
Err_Enter:
    MsgBox "The error " & Err.Number & " has been thrown." & vbCrLf & vbCrLf & _
            "Error Source: GetFileExt" & vbCrLf & _
            "Error Description: " & Err.Description, _
            vbCritical, "An Error has Occured!"
            
    Err.Clear
    Resume Err_Exit
End Function


Function GetWorksheetNames(ByVal fName As String) As Variant
    Dim objXl As Object
    Dim objWb As Object
    Dim objWs As Object
    Dim names() As String
    Dim namesCount As Integer
    
    Set objXl = CreateObject("Excel.Application")
    Set objWb = objXl.Workbooks.Open(fName)
    
    namesCount = 0
    For Each objWs In objWb.Worksheets
        ReDim Preserve names(namesCount)
        names(namesCount) = objWs.Name
        namesCount = namesCount + 1
    Next
    GetWorksheetNames = names
    
    
    Set objWs = Nothing
    objWb.Close False
    Set objWb = Nothing
    objXl.Quit
    Set objXl = Nothing
End Function

Function CheckWorksheetName(ByVal fName As String, ByVal wsName As String) As Boolean
    Dim wsNames() As String
    Dim ws As Variant
    
    wsNames = GetWorksheetNames(fName)
    
    If (UBound(wsNames) = 0) And LBound(wsNames) > 0 Then
        CheckWorksheetName = False
        Exit Function
    End If
    
    For Each ws In wsNames
        If wsName = ws Then
            CheckWorksheetName = True
            Exit Function
        End If
    Next
    
    CheckWorksheetName = False
End Function
