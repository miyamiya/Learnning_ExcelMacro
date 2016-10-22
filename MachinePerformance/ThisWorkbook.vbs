' used disk-sheet
Const DISK_NAME = 1
Const DISK_FREE = 2
Const DISK_TOTAL = 3
Const DISK_CREATEDAT = 4

' used memory-sheet
Const MEM_NAME = 1
Const MEM_FREE = 2
Const MEM_TOTAL = 3
Const MEM_CREATEDAT = 4

' used cpu-sheet
Const CPU_NAME = 1
Const CPU_PERSENT = 2
Const CPU_CREATEDAT = 3

' WMI Object (connect to localhost)
Dim Wmi As Object

' Now datetime
Dim CreatedAt As Date

'
' auto running
'
Sub Workbook_Open()
    Application.WindowState = xlMinimized
    Call Logging
End Sub

'
' main
'
Public Sub Logging()
    Call Initialize
    Call SetInfoForCpu
    Call SetInfoForMemory
    Call SetInfoForDiskdrive
    Call Terminate
End Sub

'
' Initialize process
'
Public Sub Initialize()
    Application.DisplayAlerts = False

    ' Set WMI Object
    Set Wmi = GetObject( _
        "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2" _
    )
    
    ' Set Now datatime
    CreatedAt = Format(Now, "yyyy/mm/dd hh:nn:ss")
End Sub

'
' Terminate process
'
Public Sub Terminate()
    
    Set Wmi = Nothing
    ThisWorkbook.Save
    If Workbooks.Count = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close
    End If
End Sub

'
' Set cpu information to Datasheet
'
Public Sub SetInfoForCpu()
    Dim Sheet As Worksheet: Set Sheet = ThisWorkbook.Worksheets("cpu")
    Dim Row As Long: Row = Sheet.Cells(Rows.Count, CPU_NAME).End(xlUp).Row + 1

    Dim Items As Object: Set Items = Wmi.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor")
    Dim Item As Object

    For Each Item In Items
        If Item.Name <> "" Then
            Sheet.Cells(Row, CPU_NAME).Value = Item.Name
            Sheet.Cells(Row, CPU_PERSENT).Value = Item.PercentProcessorTime
            Sheet.Cells(Row, CPU_CREATEDAT).Value = CreatedAt
            Row = Row + 1
        End If
    Next

    Set Item = Nothing
    Set Items = Nothing
    Set Sheet = Nothing
End Sub

'
' Set memory information to Datasheet
'
Public Sub SetInfoForMemory()
    Dim Sheet As Worksheet: Set Sheet = ThisWorkbook.Worksheets("memory")
    Dim Row As Long: Row = Sheet.Cells(Rows.Count, MEM_NAME).End(xlUp).Row + 1

    Dim Items As Object: Set Items = Wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    Dim Item As Object

    For Each Item In Items
        If Item.FreePhysicalMemory <> "" Then
            Sheet.Cells(Row, MEM_NAME).Value = "PhysicalMemory"
            Sheet.Cells(Row, MEM_FREE).Value = Item.FreePhysicalMemory
            Sheet.Cells(Row, MEM_TOTAL).Value = Item.TotalVisibleMemorySize
            Sheet.Cells(Row, MEM_CREATEDAT).Value = CreatedAt
            Row = Row + 1
        End If
    
        If Item.FreeVirtualMemory <> "" Then
            Sheet.Cells(Row, MEM_NAME).Value = "VirtualMemory"
            Sheet.Cells(Row, MEM_FREE).Value = Item.FreeVirtualMemory
            Sheet.Cells(Row, MEM_TOTAL).Value = Item.TotalVirtualMemorySize
            Sheet.Cells(Row, MEM_CREATEDAT).Value = CreatedAt
            Row = Row + 1
        End If
    Next

    Set Item = Nothing
    Set Items = Nothing
    Set Sheet = Nothing
End Sub

'
' Set diskdrive information to Datasheet
'
Public Sub SetInfoForDiskdrive()
    
    Dim Sheet As Worksheet: Set Sheet = ThisWorkbook.Worksheets("disk")
    Dim Row As Long: Row = Sheet.Cells(Rows.Count, DISK_NAME).End(xlUp).Row + 1
    
    Dim Items As Object: Set Items = Wmi.ExecQuery("SELECT * FROM Win32_LogicalDisk")
    Dim Item As Object

    For Each Item In Items
        If Item.DeviceID <> "" Then
            Sheet.Cells(Row, DISK_NAME).Value = Item.DeviceID
            Sheet.Cells(Row, DISK_FREE).Value = Item.FreeSpace
            Sheet.Cells(Row, DISK_TOTAL).Value = Item.Size
            Sheet.Cells(Row, DISK_CREATEDAT).Value = CreatedAt
            Row = Row + 1
        End If
    Next

    Set Item = Nothing
    Set Items = Nothing
    Set Sheet = Nothing
End Sub