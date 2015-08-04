Attribute VB_Name = "modExportSourceFiles"
Option Compare Database

Public Sub ExportSourceFiles() 'destPath As String)
'Script Name:       ExportSourceFiles
'Author:            Steve O'Neal
'Created:           5/29/2015
'Last Modified:
'Version:           1.0
'Dependency:        NONE
'
'Script used to export modules and class files as source code files
'Files are exported to 'Source' folder on H:\ drive by database name
    
    Dim db As Database
    Dim dbName As String
    Dim component As VBComponent
    Dim destPath As String
    
    dbName = Application.CurrentProject.Name
    
    destPath = "D:\Documents\Work_New\Source\" & dbName & "\"
    
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            component.Export destPath & component.Name & ToFileExtension(component.Type)
        End If
    Next
     
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
     
End Function
 
