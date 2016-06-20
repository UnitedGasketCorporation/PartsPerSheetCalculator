MessageToWindow(text, window)
{
msgbox %text%
global MessageToWindow1 = %window%
}

#SingleInstance off
#notrayicon
SetWindow = 0
global MessageToWindow1 = 0
UnitMeasuere =
PartW =
PartL =
PartSpacing =
EdgeSpace =
MatW =
MatL =
TotalW =
TotalL =
Total =
TotalWasteW =
TotalWasteL =
PercentWaste =
TotalW2 =
TotalL2 =
Total2 =
TotalWasteW2 =
TotalWasteL2 =
PercentWaste2 =

Window1:
Gui Destroy
SetWindow = 1
Gui, Add, Text,, Unit of Measure:
Gui, Add, Text,, Part Width:
Gui, Add, Text,, Part Length:
Gui, Add, Text,, Space Between Parts:
Gui, Add, Text,, Space on Edges:
Gui, Add, Text,, Material Width:
Gui, Add, Text,, Material Length:
Gui, Add, Button,, Help
Gui, Add, Edit, VUnitMeasure ym
Gui, Add, Edit, VPartW
Gui, Add, Edit, VPartL
Gui, Add, Edit, vPartSpacing
Gui, Add, Edit, vEdgeSpace
Gui, Add, Edit, VMatW
Gui, Add, Edit, VMatL
GuiControl,, UnitMeasure, %UnitMeasure%
GuiControl,, PartW, %PartW%
GuiControl,, PartL, %PartL%
Guicontrol,, PartSpacing, %PartSpacing%
GuiControl,, EdgeSpace, %EdgeSpace%
GuiControl,, MatW, %MatW%
GuiControl,, MatL, %MatL%
Gui, Add, Button, default, Figure It Out For Me
Gui, Show,, Cut Direction Calculator
return

ButtonFigureItOutForMe:
Gui, Submit
if PartW =
{
Msgbox Part Width can not be blank!
goto Window1
}
if PartW <= 0
{
Msgbox Part Width can not be zero or negative!
goto Window1
}
if PartL =
{
msgbox Part Length can not be blank!
goto Window1
}
if PartL <= 0
{
msgbox Part Length can not be zero or negative!
goto Window1
}
if PartSpacing =
{
msgbox Part Spacing can not be blank!
goto Window1
}
if PartSpacing < 0
{
msgbox Part Spacing can not be negative!
goto Window1
}
if EdgeSpace =
{
msgbox Edge Spacing can not be blank!
goto Window1
}
if EdgeSpace < 0
{
msgbox Edge Spacing can not be negative!
goto Window1
}
if MatW =
{
msgbox Material Width can not be blank!
goto Window1
}
if MatW <= 0
{
msgbox Material Width can not be zero or negative!
goto Window1
}
if MatL =
{
msgbox Material Length can not be blank!
goto Window1
}
if MatL <= 0
{
msgbox Material Length can not be zero or negative!
goto Window1
}
if UnitMeasure = asdf
{
MsgBox You do realize that asdf is not a proper measurement abbreviation, right?
}

PartW := PartW+PartSpacing
PartL := PartL+PartSpacing
MatlW := MatlW-EdgeSpace-EdgeSpace+PartSpacing
MatlL := MatlL-EdgeSpace-EdgeSpace+PartSpacing
TotalW := Floor(MatW/PartW)
TotalL := Floor(MatL/PartL)
Total := TotalW*TotalL
If TotalW = 0
TotalL = 0
If TotalL = 0
TotalW = 0
PartW := PartW-PartSpacing
PartL := PartL-PartSpacing
MatlW := MatlW+EdgeSpace+EdgeSpace-PartSpacing
MatlL := MatlL+EdgeSpace+EdgeSpace-PartSpacing
TotalWasteW := MatW-(TotalW*PartW)
TotalWasteL := MatL-(TotalL*PartL)
TotalWaste := (TotalWasteW*MatL)+(TotalWasteL*MatW)-(TotalWasteW*TotalWasteL)
PercentWaste := (TotalWaste/(MatL*MatW))*100
PartW := PartW+PartSpacing
PartL := PartL+PartSpacing
MatlW := MatlW-EdgeSpace-EdgeSpace+PartSpacing
MatlL := MatlL-EdgeSpace-EdgeSpace+PartSpacing
TotalW2 := Floor(MatW/PartL)
TotalL2 := Floor(MatL/PartW)
Total2 := TotalW2*TotalL2
If TotalW2 = 0
TotalL2 = 0
If TotalL2 = 0
TotalW2 = 0
PartW := PartW-PartSpacing
PartL := PartL-PartSpacing
MatlW := MatlW+EdgeSpace+EdgeSpace-PartSpacing
MatlL := MatlL+EdgeSpace+EdgeSpace-PartSpacing
TotalWasteW2 := MatW-(TotalW2*PartL)
TotalWasteL2 := MatL-(TotalL2*PartW)
TotalWaste2 := (TotalWasteW2*MatL)+(TotalWasteL2*MatW)-(TotalWasteW2*TotalWasteL2)
PercentWaste2 := (TotalWaste2/(MatL*MatW))*100

Window2:
Gui Destroy
SetWindow = 2
if Total > %Total2%
{
Gui, Font, bold
Gui, Add, Text,, > Non-Rotated Results <
Gui, Font, norm
}
else
Gui, Add, Text,, Non-Rotated Results
Gui, Add, Text,, Total Output:
Gui, Add, Text,, Parts Wide:
Gui, Add, Text,, Parts High:
Gui, Add, Text,, Waste matl Vertically:
Gui, Add, Text,, Waste matl Horizontally:
Gui, Add, Text,, Total waste matl:
Gui, Add, Text,, Percent waste matl:
Gui, Add, Button, default, Help
Gui, Add, Text, ym,
Gui, Add, Edit, ReadOnly vTotal
Gui, Add, Edit, ReadOnly vTotalW
Gui, Add, Edit, ReadOnly vTotalL
Gui, Add, Edit, ReadOnly vTotalWasteL
Gui, Add, Edit, ReadOnly vTotalWasteW
Gui, Add, Edit, Readonly vTotalWaste
Gui, Add, Edit, Readonly vPercentWaste
Gui, Add, Button,, Export Excel
if Total2 > %Total%
{
Gui, Font, bold
Gui, Add, Text, ym, > Rotated Results <
Gui, Font, norm
}
else
Gui, Add, Text, ym, Rotated Results
Gui, Add, Text,, Total Output:
Gui, Add, Text,, Parts Wide:
Gui, Add, Text,, Parts High:
Gui, Add, Text,, Waste matl Vertically:
Gui, Add, Text,, Waste matl Horizontally:
Gui, Add, Text,, Total waste matl:
Gui, Add, Text,, Percent waste matl:
Gui, Add, Button,, Go Back
Gui, Add, Text, ym,
Gui, Add, Edit, ReadOnly vTotal2
Gui, Add, Edit, ReadOnly vTotalW2
Gui, Add, Edit, ReadOnly vTotalL2
Gui, Add, Edit, ReadOnly vTotalWasteL2
Gui, Add, Edit, ReadOnly vTotalWasteW2
Gui, Add, Edit, Readonly vTotalWaste2
Gui, Add, Edit, Readonly vPercentWaste2
Gui, Add, Button,, Done
Gui, Add, Text, ym, You Entered:
Gui, Add, Text,, Unit of Measure:
Gui, Add, Text,, Part Cut Width:
Gui, Add, Text,, Part Cut Length:
Gui, Add, Text,, Part Spacing:
Gui, Add, Text,, Edge Spacing:
Gui, Add, Text,, Material Width:
Gui, Add, Text,, Material Length:
Gui, Add, text, ym,
Gui, Add, Edit, ReadOnly VUnitMeasure
Gui, Add, Edit, ReadOnly VPartW
Gui, Add, Edit, ReadOnly VPartL
Gui, Add, Edit, ReadOnly vPartSpacing
Gui, Add, Edit, ReadOnly vEdgeSpace
Gui, Add, Edit, ReadOnly VMatW
Gui, Add, Edit, ReadOnly VMatL
GuiControl,, Total, %Total%
GuiControl,, TotalW, %TotalW%
GuiControl,, TotalL, %TotalL%
GuiControl,, TotalWasteL, %TotalWasteL% %UnitMeasure%
GuiControl,, TotalWasteW, %TotalWasteW% %UnitMeasure%
GuiControl,, TotalWaste, %TotalWaste% sq %UnitMeasure%
GuiControl,, PercentWaste, %PercentWaste%
GuiControl,, Total2, %Total2%
GuiControl,, TotalW2, %TotalW2%
GuiControl,, TotalL2, %TotalL2%
GuiControl,, TotalWasteL2, %TotalWasteL2% %UnitMeasure%
GuiControl,, TotalWasteW2, %TotalWasteW2% %UnitMeasure%
GuiControl,, TotalWaste2, %TotalWaste2% sq %UnitMeasure%
GuiControl,, PercentWaste2, %PercentWaste2%
GuiControl,, UnitMeasure, %UnitMeasure%
GuiControl,, PartW, %PartW%
GuiControl,, PartL, %PartL%
Guicontrol,, PartSpacing, %PartSpacing%
GuiControl,, EdgeSpace, %EdgeSpace%
GuiControl,, MatW, %MatW%
GuiControl,, MatL, %MatL%
showGUI:
Gui, Show,, Cut Direction Calculator
return
ButtonGoBack:
goto window1
ButtonHelp:
Gui Destroy
Gui, Add, Text,,
Gui, Add, Text,, If you leave PART SPACING or EDGE SPACE blank, it will assume zero.
Gui, Add, Text,, If you leave MATERIAL LENGTH blank or zero, it will assume a value 1,000 times that of the part length.
Gui, Add, Text,, Waste Matl vertically, horizontally are total, including the web. So is Total Waste Matl and Percent Waste Matl.
Gui, Add, Text,,
Gui, Add, Button,, Stop Helping Me
Gui, Show,,
return
ButtonStopHelpingMe:
goto Window%SetWindow%
ButtonExportExcel:
ErrorLevel = 0
FileSelectFile, FileLocation, S18, , , Excel File (*.csv)
if ErrorLevel  = 1
goto Window2
StringRight, ExtTest, FileLocation, 4
if ExtTest <> .csv
FileLocation = %FileLocation%.csv
FileAppend,
(
Non-Rotated Results,,Rotated Results,,You Entered,
Total Output,%Total%,Total Output,%Total2%,Unit of Measure,%UnitMeasure%
Parts Wide,%TotalW%,Parts Wide,%TotalW2%,Part Cut Width,%PartW%
Parts High,%TotalL%,Parts High,%TotalL2%,Part Cut Length,%PartL%
Waste Matl at top/bottom edges,%TotalWasteL% %UnitMeasure%,Waste Matl at top/bottom edges,%TotalWasteL2% %UnitMeasure%,Part Spacing,%PartSpacing%
Waste matl at Left/right edges,%TotalWasteW% %UnitMeasure%,Waste matl at Left/right edges,%TotalWasteW2% %UnitMeasure%,Edge Space,%EdgeSpace%
Total Waste Matl,%TotalWaste% sq %UnitMeasure%,Total Waste Matl,%TotalWaste% sq %UnitMeasure%,Material Width,%MatW%
Percent Waste Matl,%PercentWaste%,Percent Waste Matl,%PercentWaste2%,Material Length,%MatL%
), %FileLocation%
MsgBox, 4, Open File, Open the saved file now?
IfMsgBox yes
Run %FileLocation%
goto Window2
GuiClose:
GuiEscape:
ButtonDone:
ExitApp
trimmer(num)
{
Loop, % StrLen(num)
{
stringright, tester, num, 1
If (tester = "0")
stringtrimright, num, num, 1
}
Loop, % StrLen(num)
{
stringright, tester, num, 1
If (tester = ".")
stringtrimright, num, num, 1
}
return num
}