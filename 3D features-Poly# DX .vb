threeD features-Poly# DX 


Sub P3DDX()

Dim XLnCADObject As Acad3DPolyline
Dim i As Integer
Dim x As Double, y As Double, z As Double
Open "C:\Users\CINPC0075\Desktop\Komatsu\Client_DWGs\Demo-up-FS\REDlineDX.csv" For Output As #1
Dim basePnt As Variant
ThisDrawing.Utility.GetEntity XLnCADObject, basePnt, "Select a 3D Polyline"
Dim Cords As Variant
Cords = XLnCADObject.Coordinates
For i = LBound(Cords) To UBound(Cords) Step 3
x = Format(XLnCADObject.Coordinates(i), "0.000")
y = Format(XLnCADObject.Coordinates(i + 1), "0.000")
z = Format(XLnCADObject.Coordinates(i + 2), "0.000")
Print #1, x; y; z
Next i
Dim length As Double
length = Format(XLnCADObject.length, "0.000")
Print #1, vbCrLf; "Length; of; the; selected; LW; polyline Is"; length
Print #1, vbCrLf; "Number; of; Vertices; of; the; selected; polyline Is"; (UBound(Cords) + 1) / 2
Close (1)

End Sub
