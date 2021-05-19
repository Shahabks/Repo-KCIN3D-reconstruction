'Code should be placed in a .vbs file
Set objExcel = CreateObject("Excel.Application")
objExcel.Application.Run "'C:\Users\CINPC0075\Desktop\Repo-KCIN3D-reconstruction\For_UI_FS\Initial_Prototype_UI\Temp\generatorCAD.xlsm'!Module1.GetmyDatatoAutoCADC"
objExcel.DisplayAlerts = False
objExcel.Application.Quit
Set objExcel = Nothing