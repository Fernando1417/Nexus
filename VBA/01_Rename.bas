' ==============================================
' Module: Rename Files Utility
' Description: This module provides functionality to rename files in a specified folder
'              based on names listed in an Excel worksheet. It includes functions to
'              select a folder, list files, and rename them according to user input.
' Author: Fernando Chavarria
' Last update: 12 - Sept - 2025
' ==============================================



Option Explicit

Dim folderPath As String




' ==============================================
' Main Function
' ==============================================

Sub PrintNames()
    ' Print the names of all files in a folder in excel

   Call pick_a_folder

    ' Save the names of all the files from folderPath
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim row As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set ws = ActiveSheet
    row = 5
    
    For Each file In folder.Files
        ws.Cells(row, 10).Value = file.Name
        ws.Cells(row, 11).Value = file.Type
        ws.Cells(row, 12).Value = file.Size
  
        With ws.Shapes.Item(ws.Shapes.Count)
            ws.Cells(row, 13).Value = .Width & " x " & .Height
        End With

        ws.Cells(row, 15).Select
        
        Selection.InsertPictureInCell (file.Path)
        With ws.Shapes.Item(ws.Shapes.Count)
            ws.Cells(row, 13).Value = .Width
            ws.Cells(row, 14).Value = .Height
        End With
        
        
        row = row + 1
    Next file
    
    ' Clean up
    Set folder = Nothing
    Set fso = Nothing
End Sub


' This subroutine renames files based on the current name in column J and the new name in column N.
Sub RenameFilesFromSheet()
    Dim ws As Worksheet
    Dim row As Long
    Dim oldName As String
    Dim newName As String
    Dim fso As Object
    
    Application.ScreenUpdating = False
    Set ws = ActiveSheet
    Set fso = CreateObject("Scripting.FileSystemObject")
    row = 5
    
    ' Loop through the rows and rename files
    Do While ws.Cells(row, 10).Value <> ""
        oldName = ws.Cells(row, 10).Value
        newName = ws.Cells(row, 19).Value
        
        If newName <> "" Then
            oldFilePath = folderPath & "\" & oldName
            newFilePath = folderPath & "\" & newName
            
            If fso.FileExists(oldFilePath) Then
                fso.MoveFile oldFilePath, newFilePath
                ws.Cells(row, 20).Value = "done"
            End If
        End If
        
        row = row + 1
    Loop
    
    Set fso = Nothing
    Application.ScreenUpdating = True
End Sub





' ==============================================
' suporting Function
' ==============================================


Sub pick_a_folder()

    ' Ask user to pick a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        .Show
        If .SelectedItems.Count > 0 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected!", vbExclamation
            Exit Sub
        End If
    End With

End Sub




' ==============================================
' reset sheet
' ==============================================


Sub ReSet()
    Dim ws As Worksheet
    Dim row As Long
    Dim oldName As String
    Dim newName As String
    

    Set ws = ActiveSheet
    row = 5
    
    ' Loop through the rows and rename files
    Do While ws.Cells(row, 10).Value <> ""
    
    
    ws.Range(Cells(row, 10), Cells(row, 20)).Clear

        
        
        row = row + 1
    Loop
End Sub



