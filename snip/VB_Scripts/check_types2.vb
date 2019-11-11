Option Explicit


' ***********************************************************************
' Purpose: Interrogates Type, and TypeName of SelectedElement and SelectedElement(1).Value
' Assumptions: Works only at part level. Loops until Cancel or Escape
' Author: Danny Alcini
' Languages: VBA
' Locales: English
' CATIA Level: V5R14 & 16
' ***********************************************************************

Sub CATMain()

Dim CATIA As Object
Dim partDocument1 As PartDocument
Dim part1 As Part
Dim anyObjFilter(0)
Dim oMySelection As Object
Dim selMySelection As Selection
Dim Status As String
Dim oSelectedElement As SelectedElement
Dim oElement As Object
Dim strSelectionName As String
Dim strSelectedElementName As String
Dim strSelectionTypeName As String
Dim strSelectedElementTypeName As String
Dim strSelectionType As String
Dim iPopup As Integer

Set CATIA = GetObject(, "CATIA.Application")
Set partDocument1 = CATIA.ActiveDocument
Set part1 = partDocument1.Part
Set oMySelection = partDocument1.Selection
Set selMySelection = oMySelection
anyObjFilter(0) = "AnyObject"

iPopup = 1

Do Until iPopup = 2

oMySelection.Clear

Status = oMySelection.SelectElement3(anyObjFilter, "Select an Element", True, _
CATMultiSelTriggWhenSelPerf, False)

Set oSelectedElement = oMySelection.Item(1)
Set oElement = oMySelection.Item(1).Value

strSelectionName = oMySelection.Item(1).Name
strSelectionType = oMySelection.Item(1).Type
strSelectionTypeName = TypeName(oMySelection.Item(1))
strSelectedElementName = oElement.Name
strSelectedElementTypeName = TypeName(oElement)

iPopup = MsgBox( _
"Selection.Name = " + strSelectionName + vbCrLf + _
"Selection.Type = " + strSelectionType + vbCrLf + _
"Selection TypeName = " + strSelectionTypeName + vbCrLf + _
"SelectedElement.Name = " + strSelectedElementName + vbCrLf + _
"SelectedElement TypeName = " + strSelectedElementTypeName, _
vbOKCancel, "VB Type Tester")

Loop

End Sub

