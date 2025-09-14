' ==============================================
' Module: Rename Files Utility
' Description: This module provides functionality to rename files in a specified folder
'              based on names listed in an Excel worksheet. It includes functions to
'              select a folder, list files, and rename them according to user input.
' Author: Fernando Chavarria
' Last update: 12 - Sept - 2025
' ==============================================

Option Explicit

Private m_strFolderPath As String


' ==============================================
' Main Functions
' ==============================================

Sub Files_List_Names()
    ' Lists all files from selected folder into Excel worksheet

   Call Files_Select_Folder

    ' Save the names of all the files from folderPath
    Dim objFSO As Object 'do I really need this object?
    Dim objFolder As Object
    Dim objFile As Object
    Dim wsActive As Worksheet 'change to your specific sheet 
    Dim lngRow As Long
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(m_strFolderPath)
    Set wsActive = ActiveSheet
    lngRow = 5
    
    For Each objFile In objFolder.Files
        wsActive.Cells(lngRow, 10).Value = objFile.Name
        wsActive.Cells(lngRow, 11).Value = objFile.Type
        wsActive.Cells(lngRow, 12).Value = objFile.Size
  
        With wsActive.Shapes.Item(wsActive.Shapes.Count)
            wsActive.Cells(lngRow, 13).Value = .Width & " x " & .Height
        End With

        wsActive.Cells(lngRow, 15).Select
        
        Selection.InsertPictureInCell (objFile.Path)
        With wsActive.Shapes.Item(wsActive.Shapes.Count)
            wsActive.Cells(lngRow, 13).Value = .Width
            wsActive.Cells(lngRow, 14).Value = .Height
        End With
        
        
        lngRow = lngRow + 1
    Next file
    
    ' Clean up
    Set objFolder = Nothing
    Set objFSO = Nothing
End Sub


Sub Files_Rename_FromSheet()
    Dim wsActive As Worksheet
    Dim lngRow As Long
    Dim strOldName As String
    Dim strNewName As String
    Dim strOldFilePath As String
    Dim strNewFilePath As String
    Dim objFSO As Object
    
    Application.ScreenUpdating = False
    Set wsActive = ActiveSheet
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    lngRow = 5
    
    ' Loop through the rows and rename files
    Do While wsActive.Cells(lngRow, 10).Value <> ""
        strOldName = wsActive.Cells(lngRow, 10).Value
        strNewName = wsActive.Cells(lngRow, 19).Value
        
        If strNewName <> "" Then
            strOldFilePath = m_strFolderPath & "\" & strOldName
            strNewFilePath = m_strFolderPath & "\" & strNewName
            
            If objFSO.FileExists(strOldFilePath) Then
                objFSO.MoveFile strOldFilePath, strNewFilePath
                wsActive.Cells(lngRow, 20).Value = "done"
            End If
        End If
        
        lngRow = lngRow + 1
    Loop
    
    Set objFSO = Nothing
    Application.ScreenUpdating = True
End Sub





' ==============================================
' Functions
' ==============================================

Sub Files_Select_Folder()
    ' Opens folder picker dialog and stores selected path
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        .Show
        If .SelectedItems.Count > 0 Then
            m_strFolderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected!", vbExclamation
            Exit Sub
        End If
    End With

End Sub




' ==============================================
' Utility Functions
' ==============================================

Sub Sheet_Clear_Data()
    Dim wsActive As Worksheet
    Dim lngRow As Long
    Dim strOldName As String
    Dim strNewName As String
    

    Set wsActive = ActiveSheet
    lngRow = 5
    
    ' Loop through the rows and rename files
    Do While wsActive.Cells(lngRow, 10).Value <> ""
    
    
    wsActive.Range(Cells(lngRow, 10), Cells(lngRow, 20)).Clear

        
        
        lngRow = lngRow + 1
    Loop
End Sub



