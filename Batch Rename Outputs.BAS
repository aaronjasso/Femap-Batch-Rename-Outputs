Sub Main
    Dim App As femap.model
    Set App = feFemap()
	'=========================================
	'Program to batch rename output sets. Allows for a common
	'prefix shared by all and individual names for each result.
	'
	'Copyright (c) 2019 Aaron Jasso
	'
	'This program is free software: you can redistribute it and/or modify
	'it under the terms of the GNU General Public License as published by
	'the Free Software Foundation, either version 3 of the License, or
	'(at your option) any later version.
	'
	'This program is distributed in the hope that it will be useful,
	'but WITHOUT ANY WARRANTY; without even the implied warranty of
	'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	'GNU General Public License for more details.
	'
	'You should have received a copy of the GNU General Public License
	'along with this program.  If not, see <https://www.gnu.org/licenses/>.
	'
	'----------------------------------------------------------------------------------------------------
	'Please comment or contribute at https://github.com/aaronjasso/Femap-Batch-Rename-Outputs
	'
	'Version 3 changes (5.17.2019):
	'-Changed output set selection method to open to the list view
	'
	'Version 2 changes:
	'-Added a suffix Input For things Like "(nonlinear)"
	'
	'Written by AMJ, 11/17/2017
	'=========================================
	Dim outSet As femap.Set
	Set outSet = App.feSet
	Dim rc As Variant

	'get outputs to rename
	'rc = outSet.Select(FT_OUT_CASE,True,"Select Output Sets to rename")
	rc = outSet.SelectMultiID(FT_OUT_CASE,1,"Select Output Sets to Rename")

	'hande Cancel button
	If rc = 2 Then
		Exit Sub
	End If

	If outSet.Count = 0 Then
		Exit Sub
	End If

	'initiate popup box
	Begin Dialog UserDialog 600,399,"Rename Outout Sets" ' %GRID:10,7,1,1
		GroupBox 150,35,440,84,"",.GroupBox3
		GroupBox 10,0,580,35,"",.GroupBox4
		Text 160,49,420,63,"The Prefix will be used at the beginning of each new name and the Suffix appended to the end, include any desired seperators like dashes or spaces. The Individual Set Names will be applied in numerical order. If the end is reached, it will start again at the top. No fields are required.",.Text1
		Text 40,14,520,14,outSet.Count & " Output Sets Selected",.Text2,2
		GroupBox 10,35,140,42,"Prefix",.GroupBox1
		GroupBox 10,77,140,42,"Suffix",.GroupBox5
		TextBox 30,49,100,21,.Prefix
		TextBox 30,91,100,21,.Suffix
		GroupBox 10,119,580,231,"Individual Set Names",.GroupBox2
		TextBox 30,140,540,196,.SetNames,1
		OKButton 100,357,150,35
		PushButton 310,357,150,35,"Cancel",.PushButton1
	End Dialog
	Dim dlg As UserDialog
	'dlg.Prefix = "Run:  "
	'Dialog dlg

	'hande Cancel button
	If Dialog(dlg) =1 Then
	'If rc = 1 Then
		Exit Sub
	End If

	'rename output sets
	Dim currentResults As femap.OutputSet
	Set currentResults = App.feOutputSet
	Dim names() As String
	Dim i As Integer

	names = Split(dlg.setnames,vbNewLine)
	i = 0

	'don't use last line if it's blank
	If names(UBound(names)) = "" Then
		ReDim Preserve names(0 To UBound(names) - 1)
	End If

	While outSet.Next
		currentResults.Get(outSet.CurrentID)
		currentResults.title = dlg.Prefix & names(i) & dlg.Suffix
		currentResults.Put(outSet.CurrentID)
		If i < UBound(names) Then
			i = i + 1
		Else
			i = 0
		End If
	Wend
    
End Sub
