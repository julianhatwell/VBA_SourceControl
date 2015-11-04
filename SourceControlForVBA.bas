Attribute VB_Name = "Module1"
Option Explicit

Sub ExtractFromWorkBook()
    ' container to manage the extract tasks
    
    ' prepare to release control to the process
    Dim ControlFile As String
    ControlFile = ActiveWorkbook.Name
    
    ' get the workbook to extract
    Dim wkb As Workbook
    Set wkb = Workbooks.Open(GetExistingFilePath("xls"))
    
    'extract the files
    Call SaveCodeModules(wkb)
    
    'tidy up
    wkb.Close (False)
    Set wkb = Nothing
    
    ' resume control in the original workbook
    Windows(ControlFile).Activate
    
End Sub


Sub SaveCodeModules(ByRef wkb As Workbook)

    Dim ExtractProject As VBProject
    Set ExtractProject = wkb.VBProject
    
    Dim ExportPath As String
    ExportPath = GetPath()
        
    Dim vbc As VBComponent
    Dim ModuleName As String
    
    ' here the work is done to extract all the code module files
    For Each vbc In wkb.VBProject.VBComponents
        If vbc.CodeModule.CountOfLines > 0 Then
            ModuleName = vbc.CodeModule.Name
            vbc.Export ExportPath & "\" & ModuleName & ".vba"
        End If
    Next vbc
    
    'tidy up
    Set ExtractProject = Nothing
    Set vbc = Nothing
    
End Sub


Sub ImportCodeModules()
    
    ' give user option to pick a file
    Dim UserFile As Integer
    UserFile = MsgBox("Do you wish to create a new file?" & vbCrLf & vbCrLf & _
    "YES for new, NO to select existing file or CANCEL to exit.", vbYesNoCancel, "Select File")
       
    Dim wkb As Workbook
    
    Select Case UserFile
    Case Is = vbYes
        Set wkb = Workbooks(GetNewWorkbook())
    Case Is = vbNo
        Set wkb = Workbooks.Open(GetExistingFilePath("xls"))
    Case Else
        Exit Sub
    End Select
        
    ' remove any existing VBA components, but warn first
    Dim UserOK As Integer
    UserOK = MsgBox("WARNING: This will delete existing VB components from the file. Are you sure?", vbOKCancel, "Warning")
    
    If UserOK <> vbOK Then
        Exit Sub
    End If
    
    Call RemoveComponents(wkb)
    
    'import files
    Dim fso As New FileSystemObject
    Dim file As Variant
    Dim SourceFolder As String
    
    SourceFolder = GetPath()
    
    With wkb.VBProject
        For Each file In fso.GetFolder(SourceFolder).Files
            If InStr(1, file.Name, ".vba") Then
                .VBComponents.Import file
            End If
        Next file
    End With
    
    Set fso = Nothing
    Set wkb = Nothing
End Sub

Sub RemoveComponents(ByRef wkb As Workbook)

    Dim vbc As VBComponent
    With wkb.VBProject
        For Each vbc In .VBComponents
            If vbc.Type = vbext_ct_MSForm Or vbc.Type = vbext_ct_StdModule Then
                .VBComponents.Remove .VBComponents(vbc.Name)
            End If
        Next vbc
    End With
    
    Set vbc = Nothing
End Sub

Function GetPath()

    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetPath = sItem
    Set fldr = Nothing
  
End Function

Function GetNewWorkbook() As String

Dim NewBook As Workbook
Set NewBook = Workbooks.Add

GetNewWorkbook = NewBook.Name

End Function


Function GetExistingFilePath(ByRef FileType As String) As String
Dim UserChoice As Integer
Dim FilePath As String

    FileType = Left(FileType, 3)

    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    
    Call Application.FileDialog(msoFileDialogOpen).Filters.Clear
    
    Call Application.FileDialog(msoFileDialogOpen).Filters.Add( _
        FileType & " only", "*." & FileType)
        
    Call Application.FileDialog(msoFileDialogOpen).Filters.Add( _
        "All files", "*.*")
        
    UserChoice = Application.FileDialog(msoFileDialogOpen).Show

    If UserChoice <> 0 Then
        FilePath = Application.FileDialog( _
            msoFileDialogOpen).SelectedItems(1)
    
    End If
    
    GetExistingFilePath = FilePath

End Function

Sub SetHeaders()
    
    Dim i As Integer, j As Integer
    Dim HeaderValues(4) As String

    HeaderValues(0) = "Booked Rooms"
    HeaderValues(1) = "Available Rooms"
    HeaderValues(2) = "Booked Seats"
    HeaderValues(3) = "Booked Capacity"
    HeaderValues(4) = "Available Capacity"

    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    
    Range("B1:F1").Select
    Selection.Merge
    
    Range("G1:K1").Select
    Selection.Merge
    
    Range("B1:K1").Select
    Selection.AutoFill Destination:=Range("B1:BI1"), Type:=xlFillDefault
    
    With ActiveSheet
        For j = 0 To 11
            For i = 0 To 4
                .Cells(2, i + 2 + (j * 5)).Value = HeaderValues(i)
            Next i
        Next j
           
    End With
    
    Range("B1:BI2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

    Rows("3:3").Select
    ActiveWindow.FreezePanes = True
End Sub

