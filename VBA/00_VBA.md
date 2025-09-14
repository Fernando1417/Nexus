# general guid for writing VBA for MS Excel

## sections of subrutine

### First section: header + signature

Start every module with `Option Explicit`, then give each Sub a clear header comment and a strongly‑typed signature. This “first section” sets intent, inputs, and error policy.

Template (copy/paste at the top of a Sub):

' ---\[ Sub Header \]-----------------------------------------------------------  
' Name: SubName  
' Purpose: One‑line summary of what this subroutine does.  
' Inputs: arg1 As Long – meaning; arg2 As String – meaning  
' Outputs: None (writes to sheet/range X) ' or: Returns via ByRef/out params  
' Side effects:Changes selection? Writes to sheet? Shows messages?  
' Errors: Raises on invalid input / shows MsgBox / returns early  
' Preconditions:ActiveSheet exists; named range "Data" exists (example)  
' ---------------------------------------------------------------------------  
Public Sub SubName(ByVal arg1 As Long, ByVal arg2 As String, \_  
Optional ByVal verbose As Boolean = False)  
Const APP\_NAME As String = "MyWorkbook"  
On Error GoTo ErrHandler

## names recomenadtion

## Option Explicit

Option Explicit tells VBA to require explicit variable declarations.

Purpose: Forces you to Dim/Private/Public/Static every variable before use.  
Benefit: Catches typos and scope mistakes at compile time instead of failing silently.  
Where: Put it at the very top of a module; it applies to that module only.  
Declarations: Use Dim x As Long (or appropriate type) before using x.  
IDE setting: In the VB Editor, Tools → Options → Editor → check “Require Variable Declaration” to auto‑insert it in new modules.  
Example: Without it, total = totla + 1 creates a new variant totla; with it, you get a compile error instead.

## Get File Path

```
Function Utilities_Get_File_Path(fileName As String) As String
    ' This function displays a file dialog box to allow the user to select a file and
    ' returns the selected file path.
    ' This function works on both Windows and macOS


    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
        Title:="Select a File for " & fileName, _
        ButtonText:="Select a File for " & fileName)
    
    Utilities_Get_File_Path = IIf(filePath <> False, filePath, "")
End Function
```