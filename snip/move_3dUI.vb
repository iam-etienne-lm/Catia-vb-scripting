Option Explicit

Sub CATMain()
'On Error Resume Next
'Dim specsAndGeomWindow1 As SpecsAndGeomWindow
'Set specsAndGeomWindow1 = CATIA.ActiveWindow

Dim productdocument1 As ProductDocument, product1 As Product, products1 As Products, product2 As Product

Set productdocument1 = CATIA.ActiveDocument
Set product1 = productdocument1.Product
Set products1 = product1.Products
Set product2 = products1.Item(2)

Dim ang_degtorad As Double, Xrot As Double, Yrot As Double, Zrot As Double
ang_degtorad = 0.01745 'convert deg to rad

'Dim Xrot
'Dim Yrot
'Dim Zrot

Xrot = InputBox("Rotation around X axis", "X Rotation", "0")

If Xrot = 0 Then
    Xrot = 1

Else

    Xrot = Xrot * ang_degtorad
    Xrot = Sin(Xrot)

End If

'MsgBox Xrot
Yrot = InputBox("Rotation around Y axis", "Y Rotation", "0")

If Yrot = 0 Then
    Yrot = 1

Else

    Yrot = Yrot * ang_degtorad
    Yrot = Sin(Yrot)

End If

'MsgBox Yrot
Zrot = InputBox("Rotation around Z axis", "Z Rotation", "0")

If Zrot = 0 Then
    Zrot = 1

Else

    Zrot = Zrot * ang_degtorad
    Zrot = Sin(Zrot)

End If

'MsgBox Zrot
Dim move1 As Move, move2 As Move, move3 As Move
Dim arrayx(11) As Double, arrayy(11) As Double, arrayz(11) As Double
''Dim arrayy(11)
''Dim arrayz(11)
'
''========================================
'
'If (Xrot <> 1 And Xrot > 0) Then
'
'    MsgBox "In X positive rotation"
'
''    Set move1 = product2.Move
''    Set move1 = move1.MovableObject
'
'    arrayx(0) = 1
'    arrayx(1) = 0
'    arrayx(2) = 0
'
'    arrayx(3) = 0
'    arrayx(4) = Xrot
'    arrayx(5) = Xrot
'
'    arrayx(6) = 0
'    arrayx(7) = -Xrot
'    arrayx(8) = Xrot
'
'    arrayx(9) = 0#
'    arrayx(10) = 0#
'    arrayx(11) = 0#
''    move1.Apply arrayx
'    product2.Move.Apply arrayx
'
'End If
'
'If Xrot <> 1 And Xrot < 0 Then
'
'    MsgBox "In X negative rotation"
'
''    Set move1 = product2.Move
''    Set move1 = move1.MovableObject
'
'    arrayx(0) = 1
'    arrayx(1) = 0
'    arrayx(2) = 0
'
'    arrayx(3) = 0
'    arrayx(4) = -Xrot
'    arrayx(5) = Xrot
'
'    arrayx(6) = 0
'    arrayx(7) = -Xrot
'    arrayx(8) = -Xrot
'
'    arrayx(9) = 0#
'    arrayx(10) = 0#
'    arrayx(11) = 0#
''    move1.Apply arrayx
'    product2.Move.Apply arrayx
'End If
'
''========================================
'
'If Yrot <> 1 And Yrot > 0 Then
'
'    MsgBox "In Y positive rotation"
'
''    Set move2 = product2.Move
''    Set move2 = move2.MovableObject
'
'    arrayy(0) = Yrot
'    arrayy(1) = 0
'    arrayy(2) = -Yrot
'
'    arrayy(3) = 0
'    arrayy(4) = 1
'    arrayy(5) = 0
'
'    arrayy(6) = Yrot
'    arrayy(7) = 0
'    arrayy(8) = Yrot
'
'    arrayy(9) = 0#
'    arrayy(10) = 0#
'    arrayy(11) = 0#
''    move2.Apply arrayy
'    product2.Move.Apply arrayy
'
'End If
'
'If Yrot <> 1 And Yrot < 0 Then
'
'    MsgBox "In Y negative rotation"
'
''    Set move2 = product2.Move
''    Set move2 = move2.MovableObject
'
'
'    arrayy(0) = -Yrot
'    arrayy(1) = 0
'    arrayy(2) = -Yrot
'
'    arrayy(3) = 0
'    arrayy(4) = 1
'    arrayy(5) = 0
'
'    arrayy(6) = Yrot
'    arrayy(7) = 0
'    arrayy(8) = -Yrot
'
'    arrayy(9) = 0#
'    arrayy(10) = 0#
'    arrayy(11) = 0#
''    move2.Apply arrayy
'    product2.Move.Apply arrayy
'
'End If
'
''========================================
'
'If Zrot <> 1 And Zrot > 0 Then
'
'    MsgBox "In Z positive rotation"
'
'    'Set move3 = product2.Move
'    'Set move3 = move3.MovableObject
'
'
'    arrayz(0) = Zrot
'    arrayz(1) = Zrot
'    arrayz(2) = 0
'
'    arrayz(3) = -Zrot
'    arrayz(4) = Zrot
'    arrayz(5) = 0
'
'    arrayz(6) = 0
'    arrayz(7) = 0
'    arrayz(8) = 1
'
'    arrayz(9) = 0#
'    arrayz(10) = 0#
'    arrayz(11) = 0#
'    'move3.Apply arrayz
'    product2.Move.Apply arrayz
'
'End If
'
'If Zrot <> 1 And Zrot < 0 Then
'
'    MsgBox "In Z negative rotation"
'
'    ' Set move3 = product2.Move
'    ' Set move3 = move3.MovableObject
'
'
'    arrayz(0) = -Zrot
'    arrayz(1) = Zrot
'    arrayz(2) = 0
'
'    arrayz(3) = -Zrot
'    arrayz(4) = -Zrot
'    arrayz(5) = 0
'
'    arrayz(6) = 0
'    arrayz(7) = 0
'    arrayz(8) = 1
'
'    arrayz(9) = 0#
'    arrayz(10) = 0#
'    arrayz(11) = 0#
'
''     move3.Apply arrayz
'    product2.Move.Apply arrayz
'
'End If


End Sub

