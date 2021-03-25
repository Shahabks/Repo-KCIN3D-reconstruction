threeD features-Poly# Reconstruction

Sub Generate3DFPoly()

           
    'Declaring the necessary variables.
    Dim acadApp                 As Object
    Dim acadDoc                 As Object
    Dim LastRow                 As Long
    Dim acad3DPol               As Object
    Dim dblCoordinates()        As Double
    Dim i                       As Long
    Dim j                       As Long
    Dim k                       As Long
    Dim objCircle(0 To 0)       As Object
    Dim CircleCenter(0 To 2)    As Double
    Dim CircleRadius            As Double
    Dim RotPoint1(2)            As Double
    Dim RotPoint2(2)            As Double
    Dim Regions                 As Variant
    Dim objSolidPol             As Object
    Dim FinalPosition(0 To 2)   As Double
    
    
    'Activate the coordinates sheet and find the last row.
    With Sheets("Coordinates")
        .Activate
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
        
    'Check if there are at least two points.
    If LastRow < 3 Then
        MsgBox "There are not enough points to draw the 3D polyline!", vbCritical, "Points Error"
        Exit Sub
    End If
    
    'Check if AutoCAD application is open. If not, create a new instance and make it visible.
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then
        Set acadApp = CreateObject("AutoCAD.Application")
        acadApp.Visible = True
    End If
    
    'Check if there is an AutoCAD object.
    If acadApp Is Nothing Then
        MsgBox "Sorry, it was impossible to start AutoCAD!", vbCritical, "AutoCAD Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    'Check if there is an active drawing. If no active drawing is found, create a new one.
    On Error Resume Next
    Set acadDoc = acadApp.ActiveDocument
    If acadDoc Is Nothing Then
        Set acadDoc = acadApp.Documents.Add
    End If
    On Error GoTo 0
        
    'Get the one dimensional array size (= 3 * number of coordinates (x,y,z)).
    ReDim dblCoordinates(1 To 3 * (LastRow - 1))
    
    'Pass the coordinates to the one dimensional array.
    k = 1
    For i = 2 To LastRow
        For j = 1 To 3
            dblCoordinates(k) = Sheets("Coordinates").Cells(i, j)
            k = k + 1
        Next j
    Next i
    
    'Check if the active space is paper space and change it to model space.
    If acadDoc.ActiveSpace = 0 Then '0 = acPaperSpace in early binding
        acadDoc.ActiveSpace = 1 '1 = acModelSpace in early binding
    End If
    
    'Draw the 3D polyline at model space.
    Set acad3DPol = acadDoc.ModelSpace.Add3DPoly(dblCoordinates)
    
    'Leave the 3D polyline open (the last point is not connected to the first one).
    'Set the next line to True if you need to close the polyline.
    acad3DPol.Closed = False
    acad3DPol.Update
    
    'Get the circle radius.
    CircleRadius = Sheets("Coordinates").Range("E1").Value
    
    If CircleRadius > 0 Then

        'Set the circle center at the (0,0,0) point.
        CircleCenter(0) = 0: CircleCenter(1) = 0: CircleCenter(2) = 0
        
        'Draw the circle.
        Set objCircle(0) = acadDoc.ModelSpace.AddCircle(CircleCenter, CircleRadius)
        
        'Initialize the rotational axis.
        RotPoint1(0) = 0: RotPoint1(1) = 0: RotPoint1(2) = 0
        RotPoint2(0) = 0: RotPoint2(1) = 10: RotPoint2(2) = 0
        
        'Rotate the circle in order to avoid errors with AddExtrudedSolidAlongPath method.
        objCircle(0).Rotate3D RotPoint1, RotPoint2, 0.785398163 '45 degrees

        'Create a region from the circle.
        Regions = acadDoc.ModelSpace.AddRegion(objCircle)
    
        'Create the "solid polyline".
        Set objSolidPol = acadDoc.ModelSpace.AddExtrudedSolidAlongPath(Regions(0), acad3DPol)
                
        'Set the position where the solid should be transfered after its design (its original position).
        With Sheets("Coordinates")
            FinalPosition(0) = .Range("A2").Value
            FinalPosition(1) = .Range("B2").Value
            FinalPosition(2) = .Range("C2").Value
        End With
        
        'Move the solid to its final position.
        objSolidPol.Move CircleCenter, FinalPosition
           
        'Delete the circle.
        objCircle(0).Delete
        
        'Delete the region.
        Regions(0).Delete
                      
        'If the "solid polyline" was created successfully delete the initial polyline.
        If Err.Number = 0 Then
            acad3DPol.Delete
        End If
        
    End If

    'Zooming in to the drawing area.
    acadApp.ZoomExtents
    
    'Release the objects.
    Set objCircle(0) = Nothing
    Set objSolidPol = Nothing
    Set acad3DPol = Nothing
    Set acadDoc = Nothing
    Set acadApp = Nothing
    
    'Inform the user that the 3D polyline was created.
    MsgBox "The 3D polyline was successfully created in AutoCAD!", vbInformation, "Finished"

End Sub

Sub ClearCoordinates()
    
    Dim LastRow As Long
    
    Sheets("Coordinates").Activate
    
    'Find the last row and clear all the input data..
    With Sheets("Coordinates")
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A2:C" & LastRow).ClearContents
        .Range("E1").ClearContents
        .Range("A2").Select
    End With
    
End Sub
