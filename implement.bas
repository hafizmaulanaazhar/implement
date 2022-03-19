
Sub Main

End Sub

Dim oDialog As Object
Dim oLabeb As Object

Sub OpenDGL_AddMhs()
	BasicLibraries.LoadLibrary("Tools")
	oDialog = LoadDialog("Standard", "AddMhs", DialogLibraries)
	oDialog.Execute()
End Sub
'===============================================================
Sub ReadNSaveDLG_AddMhs()
	nim = oDialog.GetControl("nim")
	nama = oDialog.GetControl("nama")
	nilai = oDialog.GetControl("nilai")
	lastrow = LastRowNumber()
	ThisComponent.Sheet(0).getCellByPosition(2,lastrow).setValue(nim.Text)
	ThisComponent.Sheet(0).getCellByPosition(3,lastrow).setValue(nama.Text)
	ThisComponent.Sheet(0).getCellByPosition(4,lastrow).setValue(nilai.Text)
	nim.Text = ""
	nama.Text = ""
	nilai.Text = ""
End Sub
'==================================================================
Function LastRowNumber() as long
	Dim oDoc As Object
	Dim lastRow as Long
	Dim oSheet as object
	Dim oCol,rd,find,aray
	
	oDoc = ThisComponent
	oSheet = oDoc.Sheets().getByName("DataMhs")
	oCol = oSheet.getColums().getByIndex(2)
	rd = oCol.createReplaceDescriptor
	rd.searchRegularExpression = true
	rd.setSearchString(".")
	find = oCol.FindAll(rd)
	aray = Split(find.AbsoluteName. "$")
	LastRowNumber = aray(unbound(Aray))
End Function
'===================================================================
Sub DeleteRow()
	Dim document as object
	Dim dispatcher as object
	Dim returnvalue as string
	
	document = ThisComponent.CurrentController.frame
	dispatcher = createUnoService("com.sun.star.frame.DispatcHerlper")
	
	Dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = get_range_address()
	returnvalue = MsgBox ("Apakah yakin anda ingin menghapus data?", 1, "Konfirmasi")
	if returnvalaue = 1 Then
		dispatcher.executeDispatch(document, ".uno.GoToCell", "", 0, args1())
		dispatcher.executeDispatch(document, ".uno.ClearContens", "", 0, Array())
	End If
End Sub
'=======================================================================
Function get_range_address() as string

	oActivateCell = ThisComponent.getCurrentSelection()
	oConv = ThisComponent.createInstance("com.sun.star.table.CellRangeAddressConversion")
	oConv.Address = oActiveCell.getRangeAddress
	get_range_address = oConv.PersistentRepresentation
End Function
'=======================================================================
	
		
