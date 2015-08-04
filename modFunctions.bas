Attribute VB_Name = "modFunctions"
Option Compare Database

Function Get_Directory(ByRef strMessage As String) As String
    'Function allows users to select a directory
    'This function is specifically designed to select directory on I:\ drive
    'Returns the directory path as a string
    
    On Error GoTo BadDirections
    
    Dim objFolderRef As Object
    Set objFolderRef = CreateObject("Shell.Application").BrowseForFolder _
    (0, strMessage, &H4000, "C:\")
    If Not objFolderRef Is Nothing Then
        Get_Directory = objFolderRef.items.Item.path
    Else
        Get_Directory = vbNullString
    End If
    
    Set objFoderRef = Nothing
    Exit Function
    
BadDirections:
    Set objFolderRef = Nothing
    Get_Directory = "Error Selecting a Folder"

End Function


Public Function GetFileName()
    'Function allows users to select a directory and a file
    'Requires following references
    'Visual Basic for Applications
    'Microsoft Access 12.0 Object Library
    'OLE Automation
    'Microsoft Visual Basic for Applications Extensibility
    'Microsoft ActiveX Data Objects 2.1 Library
    'Microsoft Office 12 Object Library
    
    Dim result As Integer
    Dim fileName As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select file"
        .Filters.Add "All Files", "*.*"
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = False
        '.InitialFileName = CurrentProject.Path
        
        result = .Show
        If (result <> 0) Then
            fileName = Trim(.SelectedItems.Item(1))
        End If
    End With
    
    GetFileName = fileName
    
End Function


Function TodayDate()
    'Function pulls the current day's date
    'Returns the date as a String in MMDDYY format

    Dim today As String
    Dim mnth As String
    Dim yr As String
    
    today = Day(Date)
    mnth = Month(Date)
    yr = Year(Date)
    
    If Len(mnth) = 1 Then
        mnth = "0" & mnth
    End If
    If Len(today) = 1 Then
        today = "0" & today
    End If
    
    TodayDate = mnth + today + yr

End Function


Public Function TableExists(sTable As String) As Boolean
    Dim tdf As TableDef
    
    On Error Resume Next
    
    Set tdf = CurrentDb.TableDefs(sTable)
    
    If Err.Number = 0 Then
        TableExists = True
    Else
        TableExists = False
    End If
    
End Function

Function FolderCreate(ByVal path As String) As Boolean
'FileSystemObject requires the Microsoft Scripting Runtime reference

    If Len(Dir(path, vbDirectory)) = 0 Then
        MkDir path
    End If

End Function




