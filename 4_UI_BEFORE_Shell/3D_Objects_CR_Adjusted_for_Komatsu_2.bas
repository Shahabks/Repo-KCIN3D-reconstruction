Attribute VB_Name = "Module1"
'Declaring the API Sleep subroutine.
#If VBA7 And Win64 Then
    'For 64 bit Excel.
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    'For 32 bit Excel.
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

Sub GetmyDatatoAutoCADC()

      
    'Declaring the necessary variables.
    Dim acadApp     As Object
    Dim acadDoc     As Object
    Dim acadCmd     As String
    Dim sht         As Worksheet
    Dim LastRow     As Long
    Dim LastColumn  As Integer
    Dim i           As Long
    Dim j           As Integer
    
    'Set the sheet name that contains the commands.
    Set sht = ThisWorkbook.Sheets("Sheet")
    
    'Activate the Send AutoCAD Commands sheet and find the last row.
    With sht
        .Activate
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
        
    'Check if there is at least one command to send.
    If LastRow < 3 Then
        MsgBox "There are no commands to send!", vbCritical, "No Commands Error"
        sht.Range("C13").Select
        Exit Sub
    End If
    
    'Check if AutoCAD application is open. If it is not opened create a new instance and make it visible.
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("AutoCAD.Application")
        acadApp.Visible = True
    End If
        
    'Check (again) if there is an AutoCAD object.
    If acadApp Is Nothing Then
        MsgBox "Sorry, it was impossible to start AutoCAD!", vbCritical, "AutoCAD Error"
        Exit Sub
    End If
    
    'Maximize AutoCAD window.
    acadApp.WindowState = 3 '3 = acMax  in early binding
    On Error GoTo 0
    
    'If there is no active drawing create a new one.
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
    End If
    On Error GoTo 0

    'Check if the active space is paper space and change it to model space.
    If acadDoc.ActiveSpace = 0 Then '0 = acPaperSpace in early binding
        acadDoc.ActiveSpace = 1     '1 = acModelSpace in early binding
    End If
        
    With sht
    
        'Loop through all the rows of the sheet that contain commands.
        For i = 3 To LastRow
            
            'Find the last column.
            LastColumn = .Cells(i, .Columns.Count).End(xlToLeft).Column
            
            'Check if there is at least on command in each row.
            If LastColumn > 2 Then
                
                'Create a string that incorporates all the commands that exist in each row.
                acadCmd = ""
                For j = 1 To LastColumn
                    If Not IsEmpty(.Cells(i, j).Value) Then
                        acadCmd = acadCmd & .Cells(i, j).Value & vbCr
                    End If
                Next j
                 
                'Check AutoCAD version.
                If Val(acadApp.Version) < 20 Then
                    'Prior to AutoCAD 2015, in Select and Select All commands (AI_SELALL) the carriage-return
                    'character 'vbCr' is used, since another command should be applied in the selected items.
                    'In all other commands the Enter character 'Chr$(27)' is used in order to denote that the command finished.
                    If InStr(1, acadCmd, "SELECT", vbTextCompare) > 0 Or InStr(1, acadCmd, "AI_SELALL", vbTextCompare) Then
                       acadDoc.SendCommand acadCmd & vbCr
                    Else
                       acadDoc.SendCommand acadCmd & Chr$(27)
                    End If
                Else
                    'In the newest version of AutoCAD (2015) the carriage-return
                    'character 'vbCr' is applied in all commands.
                    acadDoc.SendCommand acadCmd & vbCr
                End If
            
            End If
            
            'Pause a few milliseconds  before proceed to the next command. The next line is probably optional.
            'However, I suggest to not remove it in order to give AutoCAD the necessary time to execute the command.
            Sleep 20
            
        Next i
        
    End With
    
    'Inform the user about the process.
    MsgBox "The user commands were successfully sent to AutoCAD!", vbInformation, "Done"
      
End Sub

Sub ClearAll()
    
    Dim LastRow As Long
    
    'Find the last row and clear all the input data from the sheet.
    With Sheets("Sheet")
        .Activate
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        If LastRow > 12 Then
            .Range("A3:BA" & LastRow).ClearContents
        End If
        .Range("A3").Select
    End With
    
End Sub
