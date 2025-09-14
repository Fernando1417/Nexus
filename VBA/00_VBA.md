# general guid for writing VBA for MS Excel

# sections of subrutine

### First section: header + signature

Start every module with `Option Explicit`, then give each Sub a clear header comment and a strongly‑typed signature. This “first section” sets intent, inputs, and error policy.

Template (copy/paste at the top of a Sub):

## Option Explicit

Option Explicit tells VBA to require explicit variable declarations.

Purpose: Forces you to Dim/Private/Public/Static every variable before use.  
Benefit: Catches typos and scope mistakes at compile time instead of failing silently.  
Where: Put it at the very top of a module; it applies to that module only.  
Declarations: Use Dim x As Long (or appropriate type) before using x.  
IDE setting: In the VB Editor, Tools → Options → Editor → check “Require Variable Declaration” to auto‑insert it in new modules.  
Example: Without it, total = totla + 1 creates a new variant totla; with it, you get a compile error instead.

# Naming Conventions

## Main Sub Procedures

**Convention:** Use `PascalCase` with descriptive action verbs

### Format

```
Sub [Module/Category]_[Action]_[Object]()
```

### Examples

```
' Good examples
Sub Report_Generate_MonthlySales()
Sub Data_Import_CustomerList()
Sub Chart_Update_Dashboard()
Sub Email_Send_WeeklyReport()
Sub File_Export_InventoryData()

' Avoid these
Sub DoStuff()
Sub Main()
Sub Process()
```

### Guidelines

*   Start with module/category prefix for organization
*   Use action verbs (Generate, Import, Export, Update, Send, Create, Delete)
*   Be specific about what the sub does
*   Maximum 3-4 words to maintain readability

## Functions

**Convention:** Use `PascalCase` with descriptive names indicating return value

### Format

```
Function [Module/Category]_[Action/Get]_[ReturnType]() As [DataType]
```

### Examples

```
' Utility functions
Function Utilities_Get_FilePath(fileName As String) As String
Function Utilities_Check_FileExists(filePath As String) As Boolean
Function Utilities_Format_Currency(amount As Double) As String

' Data processing functions
Function Data_Calculate_TotalSales(dataRange As Range) As Double
Function Data_Find_LastRow(ws As Worksheet) As Long
Function Data_Validate_Email(email As String) As Boolean

' Business logic functions
Function Sales_Calculate_Commission(sales As Double, rate As Double) As Double
Function Inventory_Get_StockLevel(productId As String) As Integer
```

### Guidelines

*   Prefix with module/category
*   Use "Get", "Calculate", "Check", "Validate", "Find" as common action words
*   Clearly indicate what the function returns
*   Always specify return type with `As [DataType]`

## Local Variables

**Convention:** Use `camelCase` with Hungarian notation prefix (optional but recommended)

```
' Basic data types
Dim strCustomerName As String
Dim intRowCount As Integer
Dim lngLastRow As Long
Dim dblTotalAmount As Double
Dim blnIsValid As Boolean
Dim dtStartDate As Date

' Objects
Dim wsData As Worksheet
Dim wbSource As Workbook
Dim rngData As Range
Dim chartSales As Chart

' Collections and arrays
Dim arrProductList() As String
Dim colCustomers As Collection
Dim dictLookup As Dictionary
```

## Module-Level Variables

**Convention:** Use `m_` prefix + `camelCase`

```
' At the top of the module
Private m_strConfigPath As String
Private m_wsMainData As Worksheet
Private m_blnIsInitialized As Boolean
Private m_dictSettings As Dictionary
```

# Validation and Error Handling

The **Validation-First Approach** means checking all your preconditions and requirements **before** attempting any risky operations. Think of it as "measure twice, cut once" for programming.

## Core Concept

Instead of trying operations and catching errors, you **prevent** errors by validating everything upfront:

```
Sub RecommendedPattern()    
	 ' Step 1: Validate all inputs     
	 	If Not ValidateInputs() Then Exit Sub     
    ' Step 2: Validate all resources     
    	If Not ValidateResources() Then Exit Sub     
    ' Step 3: Execute main logic (now safe)     
    	ProcessData
End Sub
```