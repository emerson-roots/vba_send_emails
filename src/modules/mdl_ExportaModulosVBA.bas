Attribute VB_Name = "mdl_ExportaModulosVBA"
Option Explicit
'
' fonte:
'https://www.rondebruin.nl/win/s9/win002.htm
'
' modified: 17/11/2020
'

'=================================================================================
'
'REFERENCE REQUIRED:
'
' Microsoft Visual Basic For Applications Extensibility 5.3
'
' Microsoft Scripting Runtime
'
'
' IMPORTANTE ATIVAR A OP��O CONFIAR EM MACROS ACESSANDO A ROTA
'
'ARQUIVO>OP��ES>CENTRAL DE CONFIABILIDADE>CONFIGURA��ES DE CENTRAL DE CONFIABILIDADE>CONFIAR NO ACESSO AO MODELO DE OBJETO VBA
'
'=================================================================================

Const cNameFolderSaveModules As String = "src"
Const cFolderModules As String = "modules"
Const cFolderForms As String = "forms"
Const cFolderClassModules As String = "class_modules"

Public Sub ExportModules()

    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    Dim folderVBAProjectFiles As String
    
    folderVBAProjectFiles = FolderWithVBAProjectFiles

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If folderVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
    Kill folderVBAProjectFiles & "\*.*"
    Kill folderVBAProjectFiles & "\" & cFolderForms & "\*.*"
    Kill folderVBAProjectFiles & "\" & cFolderModules & "\*.*"
    Kill folderVBAProjectFiles & "\" & cFolderClassModules & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
               "not possible to export the code"
        Exit Sub
    End If
    
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szExportPath = folderVBAProjectFiles & "\"
        szFileName = cmpComponent.name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
        Case vbext_ct_ClassModule
            szFileName = szFileName & ".cls"
            szExportPath = szExportPath & cFolderClassModules
        Case vbext_ct_MSForm
            szFileName = szFileName & ".frm"
            szExportPath = szExportPath & cFolderForms
        Case vbext_ct_StdModule
            szFileName = szFileName & ".bas"
            szExportPath = szExportPath & cFolderModules
        Case vbext_ct_Document
            ''' This is a worksheet or workbook object.
            ''' Don't try to export.
            bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & "\" & szFileName
            
            ''' remove it from the project if you want
            '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
    
End Sub

Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.name = ThisWorkbook.name Then
        MsgBox "Select another destination workbook" & _
               "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
               "not possible to Import the code"
        Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
        MsgBox "There are no files to import"
        Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.name) = "cls") Or _
                                                           (objFSO.GetExtensionName(objFile.name) = "frm") Or _
                                                           (objFSO.GetExtensionName(objFile.name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    'Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    'Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = ThisWorkbook.Path              'WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    
    If FSO.FolderExists(SpecialPath & cNameFolderSaveModules) = False Then
        On Error Resume Next
        MkDir SpecialPath & cNameFolderSaveModules
        MkDir SpecialPath & cNameFolderSaveModules & "\" & cFolderForms
        MkDir SpecialPath & cNameFolderSaveModules & "\" & cFolderModules
        MkDir SpecialPath & cNameFolderSaveModules & "\" & cFolderClassModules
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & cNameFolderSaveModules) = True Then
        FolderWithVBAProjectFiles = SpecialPath & cNameFolderSaveModules
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
        
    Set VBProj = ActiveWorkbook.VBProject
        
    For Each VBComp In VBProj.VBComponents
        If VBComp.Type = vbext_ct_Document Then
            'Thisworkbook or worksheet module
            'We do nothing
        Else
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
End Function


