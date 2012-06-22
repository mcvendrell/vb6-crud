To work with this CRUD system for your own forms, you will only need to:

1. Define table name in line:
		Const Tabla = "CLIENTES"
	 
2. Define initial order in line:
		mstrOrden = "APELLIDOS, NOMBRE"
	 
3. Include the mandatory fields in the INSERT and UPDATE statements, with quotes for string fields (Grabar Function):
		Sql = "insert into " & Tabla _
			  & "       (" & TxtOb(0).Tag & "," & TxtOb(1).Tag & "," & TxtOb(2).Tag & "," & TxtOb(3).Tag & ") " _
			  & "values (" & TxtOb(0) & ",'" & TxtOb(1) & "','" & TxtOb(2) & "','" & TxtOb(3) & "')"

		Sql = "update " & Tabla _
			  & "   set " & TxtOb(1).Tag & " = '" & TxtOb(1) & "'" _
			  & "     , " & TxtOb(2).Tag & " = '" & TxtOb(2) & "'" _
			  & "     , " & TxtOb(3).Tag & " = '" & TxtOb(3) & "'" _
			  & " where " & TxtOb(0).Tag & " = " & TxtOb(0)
	 
4. Indicate the index for the optional fields for each type (Grabar Function):
		If TxtPres(I) <> "" Then
			Select Case I
				Case 100
					'Campos numéricos (no hay)
					Sql = "update " & Tabla _
							& "   set " & TxtPres(I).Tag & " = " & TxtPres(I) _
							& " where " & TxtOb(0).Tag & " = " & TxtOb(0)
				Case Else
					'Campos fecha o caracter
					Sql = "update " & Tabla _
							& "   set " & TxtPres(I).Tag & " = '" & TxtPres(I) & "'" _
							& " where " & TxtOb(0).Tag & " = " & TxtOb(0)
			End Select
		Else

5. Indicate the name for combos and checkboxes to save (Grabar Function):
		'Combos
		Sql = "update " & Tabla _
				& "   set " & CmbLocalidad.Tag & " = " & IIf(CmbLocalidad = "", "Null", "'" & CmbLocalidad & "'") _
				& " where " & TxtOb(0).Tag & " = " & TxtOb(0)
		Conexion.Execute Sql

		'Chks
		Sql = "update " & Tabla _
				& "   set " & ChkBaja.Tag & " = '" & IIf(ChkBaja.Value = vbChecked, "S", "N") & "'" _
				& " where " & TxtOb(0).Tag & " = " & TxtOb(0)
		Conexion.Execute Sql

6. Indicate the type for textboxes checks and validations (TxtOb_KeyPress, TxtPres_KeyPress and TxtOb_Validate, TxtPres_Validate Functions):

7. Go to the form and assign the right names into the TAG property, on each field corresponding to the database name field.

8. Edit the DataGrid to show the fields you want (searches are based on this grid to give you the field searching options).

And this is ALL.