Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

' ===================================================================
' Access Database Documentation Exporter
' ===================================================================
' Instructions:
' 1. Open your Access database
' 2. Press Alt+F11 to open VBA Editor
' 3. Insert > Module
' 4. Paste this code
' 5. Update the OUTPUT_PATH constant below
' 6. Run ExportAllDocumentation from the Immediate window or a button
' ===================================================================

' CONFIGURATION - Change this to your desired output folder
Const OUTPUT_PATH As String = "C:\AccessExport\"

' Main export function - exports everything
Sub ExportAllDocumentation()
    On Error Resume Next
    
    ' Create output directory if it doesn't exist
    MkDir OUTPUT_PATH
    MkDir OUTPUT_PATH & "Queries\"
    MkDir OUTPUT_PATH & "Forms\"
    MkDir OUTPUT_PATH & "Reports\"
    MkDir OUTPUT_PATH & "Modules\"
    MkDir OUTPUT_PATH & "Tables\"
    
    Debug.Print "Starting export to: " & OUTPUT_PATH
    Debug.Print String(50, "=")
    
    ' Export each component type
    Call ExportQueries
    Call ExportForms
    Call ExportReports
    Call ExportModules
    Call ExportTableStructures
    Call CreateInventoryReport
    
    Debug.Print String(50, "=")
    Debug.Print "Export complete! Check: " & OUTPUT_PATH
    MsgBox "Export complete! Files saved to: " & OUTPUT_PATH, vbInformation
End Sub

' Export all queries (with error handling for broken ones)
Sub ExportQueries()
    Dim qdf As QueryDef
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    Dim successCount As Integer
    Dim errorCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    successCount = 0
    errorCount = 0
    
    Debug.Print "Exporting Queries..."
    
    For Each qdf In CurrentDb.QueryDefs
        ' Skip system queries
        If Left(qdf.Name, 4) <> "~sq_" And Left(qdf.Name, 4) <> "MSys" Then
            On Error Resume Next
            
            fileName = OUTPUT_PATH & "Queries\" & CleanFileName(qdf.Name) & ".sql"
            Set ts = fso.CreateTextFile(fileName, True)
            
            If Err.Number = 0 Then
                ts.WriteLine "-- Query Name: " & qdf.Name
                ts.WriteLine "-- Type: " & GetQueryType(qdf.Type)
                ts.WriteLine "-- Date Exported: " & Now()
                ts.WriteLine "-- " & String(70, "-")
                ts.WriteLine ""
                ts.WriteLine qdf.SQL
                ts.Close
                
                successCount = successCount + 1
                Debug.Print "  ? " & qdf.Name
            Else
                errorCount = errorCount + 1
                Debug.Print "  ? " & qdf.Name & " (Error: " & Err.Description & ")"
                Err.Clear
            End If
            
            On Error GoTo 0
        End If
    Next qdf
    
    Debug.Print "Queries: " & successCount & " exported, " & errorCount & " errors"
    Set fso = Nothing
End Sub

' Export form information and VBA code
Sub ExportForms()
    Dim obj As AccessObject
    Dim frm As Form
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    Dim successCount As Integer
    Dim errorCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    successCount = 0
    errorCount = 0
    
    Debug.Print "Exporting Forms..."
    
    For Each obj In CurrentProject.AllForms
        On Error Resume Next
        
        fileName = OUTPUT_PATH & "Forms\" & CleanFileName(obj.Name) & ".txt"
        Set ts = fso.CreateTextFile(fileName, True)
        
        If Err.Number = 0 Then
            DoCmd.OpenForm obj.Name, acDesign, , , , acHidden
            Set frm = Forms(obj.Name)
            
            ts.WriteLine "Form Name: " & obj.Name
            ts.WriteLine "Date Exported: " & Now()
            ts.WriteLine String(70, "=")
            ts.WriteLine ""
            
            ' Basic properties
            ts.WriteLine "PROPERTIES:"
            ts.WriteLine "  Record Source: " & Nz(frm.RecordSource, "(none)")
            ts.WriteLine "  Has Module: " & frm.HasModule
            ts.WriteLine "  Default View: " & GetViewType(frm.DefaultView)
            ts.WriteLine ""
            
            ' Controls
            ts.WriteLine "CONTROLS:"
            Call ExportFormControls(frm, ts)
            ts.WriteLine ""
            
            ' VBA Code
            If frm.HasModule Then
                ts.WriteLine "VBA CODE:"
                ts.WriteLine String(70, "-")
                Call ExportVBACode(obj.Name, 2, ts) ' 2 = Form module
            End If
            
            ts.Close
            DoCmd.Close acForm, obj.Name, acSaveNo
            
            successCount = successCount + 1
            Debug.Print "  ? " & obj.Name
        Else
            errorCount = errorCount + 1
            Debug.Print "  ? " & obj.Name & " (Error: " & Err.Description & ")"
            Err.Clear
        End If
        
        On Error GoTo 0
    Next obj
    
    Debug.Print "Forms: " & successCount & " exported, " & errorCount & " errors"
    Set fso = Nothing
End Sub

' Export report information
Sub ExportReports()
    Dim obj As AccessObject
    Dim rpt As Report
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    Dim successCount As Integer
    Dim errorCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    successCount = 0
    errorCount = 0
    
    Debug.Print "Exporting Reports..."
    
    For Each obj In CurrentProject.AllReports
        On Error Resume Next
        
        fileName = OUTPUT_PATH & "Reports\" & CleanFileName(obj.Name) & ".txt"
        Set ts = fso.CreateTextFile(fileName, True)
        
        If Err.Number = 0 Then
            DoCmd.OpenReport obj.Name, acViewDesign, , , acHidden
            Set rpt = Reports(obj.Name)
            
            ts.WriteLine "Report Name: " & obj.Name
            ts.WriteLine "Date Exported: " & Now()
            ts.WriteLine String(70, "=")
            ts.WriteLine ""
            
            ts.WriteLine "PROPERTIES:"
            ts.WriteLine "  Record Source: " & Nz(rpt.RecordSource, "(none)")
            ts.WriteLine "  Has Module: " & rpt.HasModule
            ts.WriteLine ""
            
            ' VBA Code if present
            If rpt.HasModule Then
                ts.WriteLine "VBA CODE:"
                ts.WriteLine String(70, "-")
                Call ExportVBACode(obj.Name, 3, ts) ' 3 = Report module
            End If
            
            ts.Close
            DoCmd.Close acReport, obj.Name, acSaveNo
            
            successCount = successCount + 1
            Debug.Print "  ? " & obj.Name
        Else
            errorCount = errorCount + 1
            Debug.Print "  ? " & obj.Name & " (Error: " & Err.Description & ")"
            Err.Clear
        End If
        
        On Error GoTo 0
    Next obj
    
    Debug.Print "Reports: " & successCount & " exported, " & errorCount & " errors"
    Set fso = Nothing
End Sub

' Export VBA modules
Sub ExportModules()
    Dim vbComp As Object
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    Dim successCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    successCount = 0
    
    Debug.Print "Exporting Modules..."
    
    On Error Resume Next
    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        If vbComp.Type = 1 Then ' Standard module
            fileName = OUTPUT_PATH & "Modules\" & CleanFileName(vbComp.Name) & ".bas"
            vbComp.Export fileName
            
            If Err.Number = 0 Then
                successCount = successCount + 1
                Debug.Print "  ? " & vbComp.Name
            Else
                Debug.Print "  ? " & vbComp.Name & " (Error: " & Err.Description & ")"
                Err.Clear
            End If
        End If
    Next vbComp
    On Error GoTo 0
    
    Debug.Print "Modules: " & successCount & " exported"
    Set fso = Nothing
End Sub

' Export table structures (not data, just schema)
Sub ExportTableStructures()
    Dim tdf As TableDef
    Dim fld As Field
    Dim idx As Index
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    Dim successCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    successCount = 0
    
    Debug.Print "Exporting Table Structures..."
    
    fileName = OUTPUT_PATH & "Tables\TableStructures.txt"
    Set ts = fso.CreateTextFile(fileName, True)
    
    ts.WriteLine "TABLE STRUCTURES"
    ts.WriteLine "Date Exported: " & Now()
    ts.WriteLine String(70, "=")
    ts.WriteLine ""
    
    For Each tdf In CurrentDb.TableDefs
        ' Skip system and temporary tables
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            On Error Resume Next
            
            ts.WriteLine "TABLE: " & tdf.Name
            
            ' Check if linked by examining Connect property
            If Len(tdf.Connect) > 0 Then
                ts.WriteLine "  Type: LINKED"
                ts.WriteLine "  Connect: " & tdf.Connect
                If Len(tdf.SourceTableName) > 0 Then
                    ts.WriteLine "  Source Table: " & tdf.SourceTableName
                End If
            Else
                ts.WriteLine "  Type: LOCAL"
                ts.WriteLine "  Records: " & tdf.RecordCount
            End If
            
            ts.WriteLine ""
            ts.WriteLine "  FIELDS:"
            
            For Each fld In tdf.Fields
                ts.WriteLine "    - " & fld.Name & " (" & GetFieldType(fld.Type) & ")" & _
                            IIf(fld.Required, " NOT NULL", "") & _
                            IIf(Len(Nz(fld.DefaultValue, "")) > 0, " DEFAULT " & fld.DefaultValue, "")
            Next fld
            
            ts.WriteLine ""
            ts.WriteLine "  INDEXES:"
            For Each idx In tdf.Indexes
                ts.WriteLine "    - " & idx.Name & IIf(idx.Primary, " (PRIMARY KEY)", "") & _
                            IIf(idx.Unique, " (UNIQUE)", "")
            Next idx
            
            ts.WriteLine ""
            ts.WriteLine String(70, "-")
            ts.WriteLine ""
            
            successCount = successCount + 1
            Debug.Print "  ? " & tdf.Name
            
            Err.Clear
            On Error GoTo 0
        End If
    Next tdf
    
    ts.Close
    Debug.Print "Tables: " & successCount & " documented"
    Set fso = Nothing
End Sub

' Create an inventory/summary report
Sub CreateInventoryReport()
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = OUTPUT_PATH & "INVENTORY.txt"
    Set ts = fso.CreateTextFile(fileName, True)
    
    ts.WriteLine "ACCESS DATABASE INVENTORY"
    ts.WriteLine "Database: " & CurrentDb.Name
    ts.WriteLine "Date: " & Now()
    ts.WriteLine String(70, "=")
    ts.WriteLine ""
    
    ts.WriteLine "SUMMARY:"
    ts.WriteLine "  Tables: " & CountObjects(acTable)
    ts.WriteLine "  Queries: " & CountObjects(acQuery)
    ts.WriteLine "  Forms: " & CountObjects(acForm)
    ts.WriteLine "  Reports: " & CountObjects(acReport)
    ts.WriteLine "  Modules: " & CountObjects(acModule)
    ts.WriteLine ""
    
    ts.WriteLine "LOCAL TABLES (stored in Access):"
    Call ListLocalTables(ts)
    ts.WriteLine ""
    
    ts.WriteLine "LINKED TABLES (SQL Server or other):"
    Call ListLinkedTables(ts)
    
    ts.Close
    Set fso = Nothing
End Sub

' Helper: Export form controls
Sub ExportFormControls(frm As Form, ts As Object)
    Dim ctl As Control
    
    On Error Resume Next
    For Each ctl In frm.Controls
        ts.WriteLine "  - " & ctl.Name & " (" & TypeName(ctl) & ")"
        
        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
            ts.WriteLine "      Control Source: " & Nz(ctl.ControlSource, "(none)")
            If ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
                ts.WriteLine "      Row Source: " & Nz(ctl.RowSource, "(none)")
            End If
        End If
        Err.Clear
    Next ctl
    On Error GoTo 0
End Sub

' Helper: Export VBA code from form/report module
Sub ExportVBACode(objName As String, objType As Integer, ts As Object)
    Dim vbComp As Object
    Dim i As Long
    
    On Error Resume Next
    Set vbComp = Application.VBE.ActiveVBProject.VBComponents(objName)
    
    If Not vbComp Is Nothing Then
        With vbComp.CodeModule
            For i = 1 To .CountOfLines
                ts.WriteLine .Lines(i, 1)
            Next i
        End With
    End If
    On Error GoTo 0
End Sub

' Helper: Clean filename for saving
Function CleanFileName(fileName As String) As String
    Dim result As String
    result = fileName
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    CleanFileName = result
End Function

' Helper: Get query type name
Function GetQueryType(qType As Integer) As String
    Select Case qType
        Case 0: GetQueryType = "Select"
        Case 48: GetQueryType = "Union"
        Case 80: GetQueryType = "MakeTable"
        Case 81: GetQueryType = "Update"
        Case 112: GetQueryType = "Delete"
        Case 96: GetQueryType = "Append"
        Case 240: GetQueryType = "DDL"
        Case 224: GetQueryType = "Crosstab"
        Case Else: GetQueryType = "Unknown (" & qType & ")"
    End Select
End Function

' Helper: Get field type name
Function GetFieldType(fType As Integer) As String
    Select Case fType
        Case 1: GetFieldType = "Boolean"
        Case 2: GetFieldType = "Byte"
        Case 3: GetFieldType = "Integer"
        Case 4: GetFieldType = "Long"
        Case 5: GetFieldType = "Currency"
        Case 6: GetFieldType = "Single"
        Case 7: GetFieldType = "Double"
        Case 8: GetFieldType = "Date"
        Case 10: GetFieldType = "Text"
        Case 11: GetFieldType = "OLE"
        Case 12: GetFieldType = "Memo"
        Case 16: GetFieldType = "BigInt"
        Case Else: GetFieldType = "Unknown (" & fType & ")"
    End Select
End Function

' Helper: Get view type
Function GetViewType(vType As Integer) As String
    Select Case vType
        Case 0: GetViewType = "Single Form"
        Case 1: GetViewType = "Continuous Forms"
        Case 2: GetViewType = "Datasheet"
        Case Else: GetViewType = "Unknown"
    End Select
End Function

' Helper: Count objects
Function CountObjects(objType As AcObjectType) As Integer
    Dim count As Integer
    Dim obj As AccessObject
    
    count = 0
    On Error Resume Next
    
    Select Case objType
        Case acTable
            For Each obj In CurrentData.AllTables
                If Left(obj.Name, 4) <> "MSys" Then count = count + 1
            Next
        Case acQuery
            For Each obj In CurrentData.AllQueries
                count = count + 1
            Next
        Case acForm
            For Each obj In CurrentProject.AllForms
                count = count + 1
            Next
        Case acReport
            For Each obj In CurrentProject.AllReports
                count = count + 1
            Next
        Case acModule
            For Each obj In CurrentProject.AllModules
                count = count + 1
            Next
    End Select
    
    CountObjects = count
End Function

' Helper: List local tables
Sub ListLocalTables(ts As Object)
    Dim tdf As TableDef
    Dim isLinked As Boolean
    
    For Each tdf In CurrentDb.TableDefs
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            On Error Resume Next
            isLinked = (Len(tdf.Connect) > 0)
            On Error GoTo 0
            
            If Not isLinked Then
                ts.WriteLine "  - " & tdf.Name & " (Records: " & tdf.RecordCount & ")"
            End If
        End If
    Next tdf
End Sub

' Helper: List linked tables
Sub ListLinkedTables(ts As Object)
    Dim tdf As TableDef
    Dim isLinked As Boolean
    Dim connectType As String
    
    For Each tdf In CurrentDb.TableDefs
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            On Error Resume Next
            isLinked = (Len(tdf.Connect) > 0)
            On Error GoTo 0
            
            If isLinked Then
                ' Determine connection type
                If InStr(1, tdf.Connect, "ODBC", vbTextCompare) > 0 Then
                    connectType = "SQL Server/ODBC"
                ElseIf InStr(1, tdf.Connect, "DATABASE=", vbTextCompare) > 0 Then
                    connectType = "Access Database"
                Else
                    connectType = "Other"
                End If
                
                ts.WriteLine "  - " & tdf.Name
                ts.WriteLine "      Type: " & connectType
                ts.WriteLine "      Connect: " & tdf.Connect
                If Len(tdf.SourceTableName) > 0 Then
                    ts.WriteLine "      Source Table: " & tdf.SourceTableName
                End If
                ts.WriteLine ""
            End If
        End If
    Next tdf
End Sub

