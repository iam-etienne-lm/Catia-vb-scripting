Sub CATMain()

'Naming instanciation required

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

Dim product2 As Product
Set product2 = products1.Item("Support Holes 14mm.1")

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

Dim reference1 As Reference
'cette ligne
Set reference1 = product1.CreateReferenceFromName("Product1/Support Holes 14mm.1/!Selection_RSur:(Face:(Brp:(AxisSystem.1;2);None:();Cf11:());AxisSystem.1;InSameTool;Z0;G4252)")

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

Dim reference3 As Reference
'cette ligne
Set reference3 = product1.CreateReferenceFromName("Product1/Support Holes 14mm.1/!Selection_REdge:(Edge:(Face:(Brp:(AxisSystem.1;1);None:();Cf11:());Face:(Brp:(AxisSystem.1;3);None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());AxisSystem.1;InSameTool;Z0;G4252)")

Dim reference4 As Reference
Set reference4 = product1.CreateReferenceFromName("")

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeOn, reference3, reference4)

product1.Update

constraint1.Orientation = catCstOrientSame

product1.Update

End Sub

