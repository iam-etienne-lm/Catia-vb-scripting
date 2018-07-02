Sub CATMain()

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

Dim products1 As Products
Set products1 = product1.Products

Dim arrayOfVariantOfBSTR1(0)
arrayOfVariantOfBSTR1(0) = "C:\Users\ng72ae5\Desktop\CATvb\mockupz\testground\yahole\Catalog  8mm a 20mm\Support Holes 14mm.CATPart"
Set products1Variant = products1
products1Variant.AddComponentsFromFiles arrayOfVariantOfBSTR1, "All"

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim productDocument2 As ProductDocument
Set productDocument2 = documents1.Add("Product")

Dim str_indexer As String
str_indexer = RandomString(3)

Dim int_countindex As Integer
int_countindex = products1.Count

Dim str_lastpart As String
str_lastpart = products1.Item(int_countindex).PartNumber

Dim product2 As Product
'need to get the name of the latest inserted part
'Set product2 = products1.Item("Support Holes 14mm." & str_indexer)
'Set product2 = products1.Item("Support Holes 14mm.4")
Set product2 = products1.Item(str_lastpart)

Dim move1 As Move
Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble1(11)
arrayOfVariantOfDouble1(0) = 1#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = 1#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
arrayOfVariantOfDouble1(9) = 0#
arrayOfVariantOfDouble1(10) = -11.271208
arrayOfVariantOfDouble1(11) = 8#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble1

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble2(11)
arrayOfVariantOfDouble2(0) = 1#
arrayOfVariantOfDouble2(1) = 0#
arrayOfVariantOfDouble2(2) = 0#
arrayOfVariantOfDouble2(3) = 0#
arrayOfVariantOfDouble2(4) = 1#
arrayOfVariantOfDouble2(5) = 0#
arrayOfVariantOfDouble2(6) = 0#
arrayOfVariantOfDouble2(7) = 0#
arrayOfVariantOfDouble2(8) = 1#
arrayOfVariantOfDouble2(9) = 5#
arrayOfVariantOfDouble2(10) = 0#
arrayOfVariantOfDouble2(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble2

Set product1 = product1.ReferenceProduct

Dim constraints1 As Constraints
Set constraints1 = product1.Connections("CATIAConstraints")


'cette ligne
'dim ref_2bicreate1 as string
'ref_2bicreate1 = "Product1/Support Holes 14mm."& str_indexer"/!Selection_RSur:(Face:(Brp:(AxisSystem.1;2);None:();Cf11:());AxisSystem.1;InSameTool;Z0;G4252)"

Dim reference1 As Reference
Set reference1 = product1.CreateReferenceFromName("Product1/Support Holes 14mm." & str_indexer & "/!Selection_RSur:(Face:(Brp:(AxisSystem.1;2);None:();Cf11:());AxisSystem.1;InSameTool;Z0;G4252)")

Dim reference2 As Reference
Set reference2 = product1.CreateReferenceFromName("")

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeDistance, reference1, reference2)

Dim length1 As Length
Set length1 = constraint1.Dimension

length1.Value = 3#

constraint1.Orientation = catCstOrientOpposite

product1.Update

Set constraints1 = product1.Connections("CATIAConstraints")

'Dim str_ref3_content
'str_ref3_content = "Product1/Support Holes 14mm." & str_indexer & "/!Selection_REdge:(Edge:(Face:(Brp:(AxisSystem.1;1);None:();Cf11:());Face:(Brp:(AxisSystem.1;3);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());AxisSystem.1;InSameTool;Z0;G4252)"

Dim reference3 As Reference
'edited naming
'Set reference3 = product1.CreateReferenceFromName(str_ref3_content)
Set reference3 = product1.CreateReferenceFromName("Product1/Support Holes 14mm." & str_indexer & "/!Selection_REdge:(Edge:(Face:(Brp:(AxisSystem.1;1);None:();Cf11:());Face:(Brp:(AxisSystem.1;3);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());AxisSystem.1;InSameTool;Z0;G4252)")

Dim reference4 As Reference
Set reference4 = product1.CreateReferenceFromName("")

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeOn, reference3, reference4)

product1.Update

constraint1.Orientation = catCstOrientSame

product1.Update

End Sub

Function RandomString(Length As Integer)
'PURPOSE: Create a Randomized String of Characters
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'edited

Dim CharacterBank As Variant
Dim x As Long
Dim str As String

'Test Length Input
  If Length < 1 Then
    MsgBox "Length variable must be greater than 0"
    Exit Function
  End If

CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
  "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", _
  "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
  "W", "X", "Y", "Z")
  

'Randomly Select Characters One-by-One
  For x = 1 To Length
    Randomize
    str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
  Next x

'Output Randomly Generated String
  RandomString = str

End Function



