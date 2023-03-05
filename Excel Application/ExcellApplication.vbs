Set objExcel = CreateObject ("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.Cells(1,1).Value = "Hello World"
objExcel.Cells(1,2).Value = "Hello World"
objExcel.Cells(1,3).Value = "Hello World"