Sub CATMain()

' Case 0 in product

' Dim productDocument1 As ProductDocument
' Set productDocument1 = CATIA.ActiveDocument

' Dim product1 As Product
' Set product1 = productDocument1.Product

' Dim products1 As Products
' Set products1 = product1.Products

' Dim product2 As Product
' Set product2 = products1.AddNewComponent("Part", "y001")

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

Dim products1 As Products
Set products1 = product1.Products

        'KNOWN UNRESOLVED ERROR ADDNEWCOMPONENT
Dim product2 As Product
On Error Resume Next
Set product2 = products1.AddNewComponent("Part", "y002")
MsgBox Err.Description
On Error GoTo 0

'CASE 1 = Creates a new part
' Set documents1 = CATIA.Documents
' Set partDocument1 = documents1.Add("Part")

'CASE2 = Run in Existing part
' Dim partDocument1 As PartDocument
' Set partDocument1 = CATIA.ActiveDocument

' Set part1 = partDocument1.Part

Set bodies1 = part1.Bodies

Set body1 = bodies1.Item("PartBody")

Set sketches1 = body1.Sketches

Set reference1 = part1.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(AxisSystem.1;2);None:();Cf11:());AxisSystem.1;Z0;G4252)")

Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 0#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = 1#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
sketch1.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

' ===========================================================================================
' -------------------------------------------------------------------------------------------
										' SKETCH ON
' -------------------------------------------------------------------------------------------
' ===========================================================================================

	Set factory2D1 = sketch1.OpenEdition()

	Set geometricElements1 = sketch1.GeometricElements

	Set axis2D1 = geometricElements1.Item("AbsoluteAxis")

	Set line2D1 = axis2D1.GetItem("HDirection")

	line2D1.ReportName = 1

	Set line2D2 = axis2D1.GetItem("VDirection")

	line2D2.ReportName = 2

	Set circle2D1 = factory2D1.CreateClosedCircle(0#, 0#, 10#)

	Set point2D1 = axis2D1.GetItem("Origin")

	circle2D1.CenterPoint = point2D1

	circle2D1.ReportName = 3

	circle2D1.Construction = True

	Set constraints1 = sketch1.Constraints

	Set reference2 = part1.CreateReferenceFromObject(circle2D1)

	Set constraint1 = constraints1.AddMonoEltCst(catCstTypeRadius, reference2)

	constraint1.Mode = catCstModeDrivingDimension

	Set length1 = constraint1.Dimension

	length1.Value = 10#
	
	' =============================================
	' =====				radius1				=====
	' =============================================

	Set point2D2 = factory2D1.CreatePoint(-12.160855, 1.064495)

	point2D2.ReportName = 4

	Set point2D3 = factory2D1.CreatePoint(-2.457235, 12.628819)

	point2D3.ReportName = 5

	Set line2D3 = factory2D1.CreateLine(-12.160855, 1.064495, -2.457235, 12.628819)

	line2D3.ReportName = 6

	line2D3.Construction = True

	line2D3.StartPoint = point2D2

	line2D3.EndPoint = point2D3

	Set reference3 = part1.CreateReferenceFromObject(line2D3)

	Set reference4 = part1.CreateReferenceFromObject(circle2D1)

	Set constraint2 = constraints1.AddBiEltCst(catCstTypeTangency, reference3, reference4)

	constraint2.Mode = catCstModeDrivingDimension

	Set point2D4 = factory2D1.CreatePoint(0#, -10#)

	point2D4.ReportName = 7

	point2D4.Construction = False

	Set point2D5 = factory2D1.CreatePoint(0#, 10.5)

	point2D5.ReportName = 8

	point2D5.Construction = False

	Set line2D4 = factory2D1.CreateLine(0#, -10#, 0#, 10.5)

	line2D4.ReportName = 9

	line2D4.Construction = True

	line2D4.StartPoint = point2D4

	line2D4.EndPoint = point2D5

	Set reference5 = part1.CreateReferenceFromObject(point2D4)

	Set reference6 = part1.CreateReferenceFromObject(circle2D1)

	Set constraint3 = constraints1.AddBiEltCst(catCstTypeOn, reference5, reference6)

	constraint3.Mode = catCstModeDrivingDimension

	Set reference7 = part1.CreateReferenceFromObject(point2D4)

	Set reference8 = part1.CreateReferenceFromObject(line2D2)

	Set constraint4 = constraints1.AddBiEltCst(catCstTypeOn, reference7, reference8)

	constraint4.Mode = catCstModeDrivingDimension

	Set reference9 = part1.CreateReferenceFromObject(line2D4)

	Set reference10 = part1.CreateReferenceFromObject(line2D2)

	Set constraint5 = constraints1.AddBiEltCst(catCstTypeVerticality, reference9, reference10)

	constraint5.Mode = catCstModeDrivingDimension

	Set point2D6 = factory2D1.CreatePoint(1.853346, 13.348507)

	point2D6.ReportName = 10

	Set point2D7 = factory2D1.CreatePoint(14.379785, -1.579922)

	point2D7.ReportName = 11

	Set line2D5 = factory2D1.CreateLine(1.853346, 13.348507, 14.379785, -1.579922)

	line2D5.ReportName = 12

	line2D5.Construction = True

	line2D5.StartPoint = point2D6

	line2D5.EndPoint = point2D7

	Set reference11 = part1.CreateReferenceFromObject(line2D5)

	Set reference12 = part1.CreateReferenceFromObject(circle2D1)

	Set constraint6 = constraints1.AddBiEltCst(catCstTypeTangency, reference11, reference12)

	constraint6.Mode = catCstModeDrivingDimension

	Set reference13 = part1.CreateReferenceFromObject(line2D5)

	Set reference14 = part1.CreateReferenceFromObject(line2D4)

	Set constraint7 = constraints1.AddBiEltCst(catCstTypeAngle, reference13, reference14)

	constraint7.Mode = catCstModeDrivingDimension

	constraint7.AngleSector = catCstAngleSector2

	Set angle1 = constraint7.Dimension

	angle1.Value = 40#
	
	' =============================================
	' =====				angle1				=====
	' =============================================

	Set point2D8 = factory2D1.CreatePoint(7.660444, 6.427876)

	point2D8.ReportName = 13

		point2D8.Construction = False

		Set reference15 = part1.CreateReferenceFromObject(circle2D1)

		Set reference16 = part1.CreateReferenceFromObject(point2D8)

		Set constraint8 = constraints1.AddBiEltCst(catCstTypeOn, reference15, reference16)

		constraint8.Mode = catCstModeDrivingDimension

		Set reference17 = part1.CreateReferenceFromObject(line2D5)

		Set reference18 = part1.CreateReferenceFromObject(point2D8)

		Set constraint9 = constraints1.AddBiEltCst(catCstTypeOn, reference17, reference18)

		constraint9.Mode = catCstModeDrivingDimension

		Set point2D9 = factory2D1.CreatePoint(-7.660444, 6.427876)

		point2D9.ReportName = 14

		point2D9.Construction = False

		Set reference19 = part1.CreateReferenceFromObject(circle2D1)

		Set reference20 = part1.CreateReferenceFromObject(point2D9)

		Set constraint10 = constraints1.AddBiEltCst(catCstTypeOn, reference19, reference20)

		constraint10.Mode = catCstModeDrivingDimension

		Set reference21 = part1.CreateReferenceFromObject(line2D3)

		Set reference22 = part1.CreateReferenceFromObject(point2D9)

		Set constraint11 = constraints1.AddBiEltCst(catCstTypeOn, reference21, reference22)

		constraint11.Mode = catCstModeDrivingDimension

		Set line2D6 = factory2D1.CreateLine(-7.660444, 6.427876, 7.660444, 6.427876)

	line2D6.ReportName = 15

	line2D6.Construction = True

	line2D6.StartPoint = point2D9

	line2D6.EndPoint = point2D8

	Set reference23 = part1.CreateReferenceFromObject(line2D6)

	Set reference24 = part1.CreateReferenceFromObject(line2D1)

	Set constraint12 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference23, reference24)

	constraint12.Mode = catCstModeDrivingDimension

	Set point2D10 = factory2D1.CreatePoint(-5.745333, 6.427876)

	point2D10.ReportName = 16

	Set point2D11 = factory2D1.CreatePoint(-3.830222, 6.427876)

	point2D11.ReportName = 17

	Set point2D12 = factory2D1.CreatePoint(-1.915111, 6.427876)

	point2D12.ReportName = 18

	Set point2D13 = factory2D1.CreatePoint(-0#, 6.427876)

	point2D13.ReportName = 19

	Set point2D14 = factory2D1.CreatePoint(1.915111, 6.427876)

	point2D14.ReportName = 20

	Set point2D15 = factory2D1.CreatePoint(3.830222, 6.427876)

	point2D15.ReportName = 21

	Set point2D16 = factory2D1.CreatePoint(5.745333, 6.427876)

	point2D16.ReportName = 22

	Set reference25 = part1.CreateReferenceFromObject(point2D9)

	Set reference26 = part1.CreateReferenceFromObject(point2D10)

	Set constraint13 = constraints1.AddBiEltCst(catCstTypeDistance, reference25, reference26)

	constraint13.Mode = catCstModeDrivingDimension

	Set length2 = constraint13.Dimension
	
	' =============================================
	' =====				thread				=====
	' =============================================

	length2.Value = 1.915111

		Set reference27 = part1.CreateReferenceFromObject(line2D6)

		Set reference28 = part1.CreateReferenceFromObject(point2D10)

		Set constraint14 = constraints1.AddBiEltCst(catCstTypeOn, reference27, reference28)

		constraint14.Mode = catCstModeDrivingDimension

		Set reference29 = part1.CreateReferenceFromObject(line2D6)

		Set reference30 = part1.CreateReferenceFromObject(point2D11)

		Set constraint15 = constraints1.AddBiEltCst(catCstTypeOn, reference29, reference30)

		constraint15.Mode = catCstModeDrivingDimension

		Set reference31 = part1.CreateReferenceFromObject(point2D10)

		Set reference32 = part1.CreateReferenceFromObject(point2D11)

		Set constraint16 = constraints1.AddBiEltCst(catCstTypeDistance, reference31, reference32)

		constraint16.Mode = catCstModeDrivingDimension

		Set reference33 = part1.CreateReferenceFromObject(line2D6)

		Set reference34 = part1.CreateReferenceFromObject(point2D12)

		Set constraint17 = constraints1.AddBiEltCst(catCstTypeOn, reference33, reference34)

		constraint17.Mode = catCstModeDrivingDimension

		Set reference35 = part1.CreateReferenceFromObject(point2D11)

		Set reference36 = part1.CreateReferenceFromObject(point2D12)

		Set constraint18 = constraints1.AddBiEltCst(catCstTypeDistance, reference35, reference36)

		constraint18.Mode = catCstModeDrivingDimension

		Set reference37 = part1.CreateReferenceFromObject(line2D6)

		Set reference38 = part1.CreateReferenceFromObject(point2D13)

		Set constraint19 = constraints1.AddBiEltCst(catCstTypeOn, reference37, reference38)

		constraint19.Mode = catCstModeDrivingDimension

		Set reference39 = part1.CreateReferenceFromObject(point2D12)

		Set reference40 = part1.CreateReferenceFromObject(point2D13)

		Set constraint20 = constraints1.AddBiEltCst(catCstTypeDistance, reference39, reference40)

		constraint20.Mode = catCstModeDrivingDimension

		Set reference41 = part1.CreateReferenceFromObject(line2D6)

		Set reference42 = part1.CreateReferenceFromObject(point2D14)

		Set constraint21 = constraints1.AddBiEltCst(catCstTypeOn, reference41, reference42)

		constraint21.Mode = catCstModeDrivingDimension

		Set reference43 = part1.CreateReferenceFromObject(point2D13)

		Set reference44 = part1.CreateReferenceFromObject(point2D14)

		Set constraint22 = constraints1.AddBiEltCst(catCstTypeDistance, reference43, reference44)

		constraint22.Mode = catCstModeDrivingDimension

		Set reference45 = part1.CreateReferenceFromObject(line2D6)

		Set reference46 = part1.CreateReferenceFromObject(point2D15)

		Set constraint23 = constraints1.AddBiEltCst(catCstTypeOn, reference45, reference46)

		constraint23.Mode = catCstModeDrivingDimension

		Set reference47 = part1.CreateReferenceFromObject(point2D14)

		Set reference48 = part1.CreateReferenceFromObject(point2D15)

		Set constraint24 = constraints1.AddBiEltCst(catCstTypeDistance, reference47, reference48)

		constraint24.Mode = catCstModeDrivingDimension

		Set reference49 = part1.CreateReferenceFromObject(line2D6)

		Set reference50 = part1.CreateReferenceFromObject(point2D16)

		Set constraint25 = constraints1.AddBiEltCst(catCstTypeOn, reference49, reference50)

		constraint25.Mode = catCstModeDrivingDimension

		Set reference51 = part1.CreateReferenceFromObject(point2D15)

		Set reference52 = part1.CreateReferenceFromObject(point2D16)

		Set constraint26 = constraints1.AddBiEltCst(catCstTypeDistance, reference51, reference52)

		constraint26.Mode = catCstModeDrivingDimension

					Set point2D17 = factory2D1.CreatePoint(5.745333, 8.788694)

					point2D17.ReportName = 23

					point2D17.Construction = False

					Set line2D7 = factory2D1.CreateLine(5.745333, 6.427876, 5.745333, 8.788694)

					line2D7.ReportName = 24

					line2D7.Construction = True

					line2D7.StartPoint = point2D16

					line2D7.EndPoint = point2D17

					Set reference53 = part1.CreateReferenceFromObject(line2D7)

					Set reference54 = part1.CreateReferenceFromObject(line2D2)

					Set constraint27 = constraints1.AddBiEltCst(catCstTypeVerticality, reference53, reference54)

					constraint27.Mode = catCstModeDrivingDimension

					Set point2D18 = factory2D1.CreatePoint(3.830222, 9.776472)

					point2D18.ReportName = 25

					point2D18.Construction = False

					Set line2D8 = factory2D1.CreateLine(3.830222, 6.427876, 3.830222, 9.776472)

					line2D8.ReportName = 26

					line2D8.Construction = True

					line2D8.StartPoint = point2D15

					line2D8.EndPoint = point2D18

					Set reference55 = part1.CreateReferenceFromObject(line2D8)

					Set reference56 = part1.CreateReferenceFromObject(line2D2)

					Set constraint28 = constraints1.AddBiEltCst(catCstTypeVerticality, reference55, reference56)

					constraint28.Mode = catCstModeDrivingDimension

					Set point2D19 = factory2D1.CreatePoint(1.915111, 10.323873)

					point2D19.ReportName = 27

					point2D19.Construction = False

					Set line2D9 = factory2D1.CreateLine(1.915111, 6.427876, 1.915111, 10.323873)

					line2D9.ReportName = 28

					line2D9.Construction = True

					line2D9.StartPoint = point2D14

					line2D9.EndPoint = point2D19

					Set reference57 = part1.CreateReferenceFromObject(line2D9)

					Set reference58 = part1.CreateReferenceFromObject(line2D2)

					Set constraint29 = constraints1.AddBiEltCst(catCstTypeVerticality, reference57, reference58)

					constraint29.Mode = catCstModeDrivingDimension

					Set point2D20 = factory2D1.CreatePoint(-1.915111, 10.323873)

					point2D20.ReportName = 29

					point2D20.Construction = False

					Set line2D10 = factory2D1.CreateLine(-1.915111, 6.427876, -1.915111, 10.323873)

					line2D10.ReportName = 30

					line2D10.Construction = True

					line2D10.StartPoint = point2D12

					line2D10.EndPoint = point2D20

					Set reference59 = part1.CreateReferenceFromObject(line2D10)

					Set reference60 = part1.CreateReferenceFromObject(line2D2)

					Set constraint30 = constraints1.AddBiEltCst(catCstTypeVerticality, reference59, reference60)

					constraint30.Mode = catCstModeDrivingDimension

					Set point2D21 = factory2D1.CreatePoint(-3.830222, 9.776472)

					point2D21.ReportName = 31

					point2D21.Construction = False

					Set line2D11 = factory2D1.CreateLine(-3.830222, 6.427876, -3.830222, 9.776472)

					line2D11.ReportName = 32

					line2D11.Construction = True

					line2D11.StartPoint = point2D11

					line2D11.EndPoint = point2D21

					Set reference61 = part1.CreateReferenceFromObject(line2D11)

					Set reference62 = part1.CreateReferenceFromObject(line2D2)

					Set constraint31 = constraints1.AddBiEltCst(catCstTypeVerticality, reference61, reference62)

					constraint31.Mode = catCstModeDrivingDimension

					Set point2D22 = factory2D1.CreatePoint(-5.745333, 8.788694)

					point2D22.ReportName = 33

					point2D22.Construction = False

					Set line2D12 = factory2D1.CreateLine(-5.745333, 6.427876, -5.745333, 8.788694)

					line2D12.ReportName = 34

					line2D12.Construction = True

					line2D12.StartPoint = point2D10

					line2D12.EndPoint = point2D22

					Set reference63 = part1.CreateReferenceFromObject(line2D12)

					Set reference64 = part1.CreateReferenceFromObject(line2D2)

					Set constraint32 = constraints1.AddBiEltCst(catCstTypeVerticality, reference63, reference64)

					constraint32.Mode = catCstModeDrivingDimension

	Set reference65 = part1.CreateReferenceFromObject(point2D22)

	Set reference66 = part1.CreateReferenceFromObject(circle2D1)

	Set constraint33 = constraints1.AddBiEltCst(catCstTypeDistance, reference65, reference66)

	constraint33.Mode = catCstModeDrivingDimension

	Set length3 = constraint33.Dimension

	' =============================================
	' =====				offset2				=====
	' =============================================
	
	length3.Value = 0.5

		Set reference67 = part1.CreateReferenceFromObject(point2D21)

		Set reference68 = part1.CreateReferenceFromObject(circle2D1)

		Set constraint34 = constraints1.AddBiEltCst(catCstTypeDistance, reference67, reference68)

		constraint34.Mode = catCstModeDrivingDimension

		Set reference69 = part1.CreateReferenceFromObject(point2D5)

		Set reference70 = part1.CreateReferenceFromObject(circle2D1)

		Set constraint35 = constraints1.AddBiEltCst(catCstTypeDistance, reference69, reference70)

		constraint35.Mode = catCstModeDrivingDimension

		Set reference71 = part1.CreateReferenceFromObject(point2D20)

		Set reference72 = part1.CreateReferenceFromObject(circle2D1)

		Set constraint36 = constraints1.AddBiEltCst(catCstTypeDistance, reference71, reference72)

		constraint36.Mode = catCstModeDrivingDimension

		Set reference73 = part1.CreateReferenceFromObject(point2D19)

		Set reference74 = part1.CreateReferenceFromObject(circle2D1)

		Set constraint37 = constraints1.AddBiEltCst(catCstTypeDistance, reference73, reference74)

		constraint37.Mode = catCstModeDrivingDimension

		Set reference75 = part1.CreateReferenceFromObject(point2D18)

		Set reference76 = part1.CreateReferenceFromObject(circle2D1)

		Set constraint38 = constraints1.AddBiEltCst(catCstTypeDistance, reference75, reference76)

		constraint38.Mode = catCstModeDrivingDimension

		Set reference77 = part1.CreateReferenceFromObject(point2D17)

		Set reference78 = part1.CreateReferenceFromObject(circle2D1)

		Set constraint39 = constraints1.AddBiEltCst(catCstTypeDistance, reference77, reference78)

		constraint39.Mode = catCstModeDrivingDimension

		sketch1.CloseEdition

'===================================================================================================
'--------------------------------------------------------------------------------------------------
'											SKETCH OFF
'--------------------------------------------------------------------------------------------------
'===================================================================================================

part1.InWorkObject = body1

part1.Update

	Set hybridBodies1 = part1.HybridBodies

	Set hybridBody1 = hybridBodies1.Add()

	Set hybridShapeFactory1 = part1.HybridShapeFactory

	Set reference79 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;7);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference80 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;8);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLinePtPt1 = hybridShapeFactory1.AddNewLinePtPt(reference79, reference80)

	hybridBody1.AppendHybridShape hybridShapeLinePtPt1

	part1.InWorkObject = hybridShapeLinePtPt1

part1.Update

	Set reference81 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	Set reference82 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;13);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle1 = hybridShapeFactory1.AddNewLineAngle(reference81, Nothing, reference82, False, 0#, 0#, 45#, False)

	Set reference83 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	hybridShapeLineAngle1.FirstUptoElem = reference83

	hybridBody1.AppendHybridShape hybridShapeLineAngle1

	part1.InWorkObject = hybridShapeLineAngle1

part1.Update

	Set reference84 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	Set reference85 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;27);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle2 = hybridShapeFactory1.AddNewLineAngle(reference84, Nothing, reference85, False, 0#, 0#, 45#, False)

	Set reference86 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	hybridShapeLineAngle2.FirstUptoElem = reference86

	hybridBody1.AppendHybridShape hybridShapeLineAngle2

	part1.InWorkObject = hybridShapeLineAngle2

part1.Update

	Set reference87 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	Set reference88 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;29);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle3 = hybridShapeFactory1.AddNewLineAngle(reference87, Nothing, reference88, False, 0#, 0#, 45#, False)

	Set reference89 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	hybridShapeLineAngle3.FirstUptoElem = reference89

	hybridBody1.AppendHybridShape hybridShapeLineAngle3

	part1.InWorkObject = hybridShapeLineAngle3

part1.Update

	Set reference90 = part1.CreateReferenceFromObject(hybridShapeLineAngle1)

	Set reference91 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;23);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle4 = hybridShapeFactory1.AddNewLineAngle(reference90, Nothing, reference91, False, 0#, 0#, 45#, False)

	Set reference92 = part1.CreateReferenceFromObject(hybridShapeLineAngle1)

	hybridShapeLineAngle4.FirstUptoElem = reference92

	hybridBody1.AppendHybridShape hybridShapeLineAngle4

	part1.InWorkObject = hybridShapeLineAngle4

part1.Update

	Set reference93 = part1.CreateReferenceFromObject(hybridShapeLineAngle4)

	Set reference94 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;25);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle5 = hybridShapeFactory1.AddNewLineAngle(reference93, Nothing, reference94, False, 0#, 0#, 45#, False)

	Set reference95 = part1.CreateReferenceFromObject(hybridShapeLineAngle4)

	hybridShapeLineAngle5.FirstUptoElem = reference95

	hybridBody1.AppendHybridShape hybridShapeLineAngle5

	part1.InWorkObject = hybridShapeLineAngle5

part1.Update

	Set reference96 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	Set reference97 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;14);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle6 = hybridShapeFactory1.AddNewLineAngle(reference96, Nothing, reference97, False, 0#, 0#, 45#, False)

	Set reference98 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	hybridShapeLineAngle6.FirstUptoElem = reference98

	hybridBody1.AppendHybridShape hybridShapeLineAngle6

	part1.InWorkObject = hybridShapeLineAngle6

part1.Update

	Set reference99 = part1.CreateReferenceFromObject(hybridShapeLineAngle6)

	Set reference100 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;33);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle7 = hybridShapeFactory1.AddNewLineAngle(reference99, Nothing, reference100, False, 0#, 0#, 45#, False)

	Set reference101 = part1.CreateReferenceFromObject(hybridShapeLineAngle6)

	hybridShapeLineAngle7.FirstUptoElem = reference101

	hybridBody1.AppendHybridShape hybridShapeLineAngle7

	part1.InWorkObject = hybridShapeLineAngle7

part1.Update

	Set reference102 = part1.CreateReferenceFromObject(hybridShapeLineAngle7)

	Set reference103 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;31);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapeLineAngle8 = hybridShapeFactory1.AddNewLineAngle(reference102, Nothing, reference103, False, 0#, 0#, 45#, False)

	Set reference104 = part1.CreateReferenceFromObject(hybridShapeLineAngle7)

	hybridShapeLineAngle8.FirstUptoElem = reference104

	hybridBody1.AppendHybridShape hybridShapeLineAngle8

	part1.InWorkObject = hybridShapeLineAngle8

part1.Update

	Set reference105 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	Set hybridShapePlaneNormal1 = hybridShapeFactory1.AddNewPlaneNormal(reference105, Nothing)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal1

	part1.InWorkObject = hybridShapePlaneNormal1

part1.Update

	Set reference106 = part1.CreateReferenceFromObject(hybridShapeLineAngle1)

	Set hybridShapePlaneNormal2 = hybridShapeFactory1.AddNewPlaneNormal(reference106, Nothing)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal2

	part1.InWorkObject = hybridShapePlaneNormal2

part1.Update

	Set reference107 = part1.CreateReferenceFromObject(hybridShapeLineAngle4)

	Set hybridShapePlaneNormal3 = hybridShapeFactory1.AddNewPlaneNormal(reference107, Nothing)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal3

	part1.InWorkObject = hybridShapePlaneNormal3

part1.Update

	Set reference108 = part1.CreateReferenceFromObject(hybridShapeLineAngle5)

	Set hybridShapePlaneNormal4 = hybridShapeFactory1.AddNewPlaneNormal(reference108, Nothing)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal4

	part1.InWorkObject = hybridShapePlaneNormal4

part1.Update

	Set reference109 = part1.CreateReferenceFromObject(hybridShapeLineAngle2)

	Set hybridShapePlaneNormal5 = hybridShapeFactory1.AddNewPlaneNormal(reference109, Nothing)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal5

	part1.InWorkObject = hybridShapePlaneNormal5

part1.Update

	Set reference110 = part1.CreateReferenceFromObject(hybridShapeLineAngle3)

	Set hybridShapePlaneNormal6 = hybridShapeFactory1.AddNewPlaneNormal(reference110, Nothing)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal6

	part1.InWorkObject = hybridShapePlaneNormal6

part1.Update

	Set reference111 = part1.CreateReferenceFromObject(hybridShapeLineAngle8)

	Set hybridShapePlaneNormal7 = hybridShapeFactory1.AddNewPlaneNormal(reference111, Nothing)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal7

	part1.InWorkObject = hybridShapePlaneNormal7

part1.Update

	Set reference112 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	Set reference113 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;7);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal1 = hybridShapeFactory1.AddNewPlaneNormal(reference112, reference113)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal1

	part1.InWorkObject = hybridShapePlaneNormal1

part1.Update

	Set reference114 = part1.CreateReferenceFromObject(hybridShapeLineAngle1)

	Set reference115 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;13);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal2 = hybridShapeFactory1.AddNewPlaneNormal(reference114, reference115)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal2

	part1.InWorkObject = hybridShapePlaneNormal2

part1.Update

	Set reference116 = part1.CreateReferenceFromObject(hybridShapeLineAngle4)

	Set reference117 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;23);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal3 = hybridShapeFactory1.AddNewPlaneNormal(reference116, reference117)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal3

	part1.InWorkObject = hybridShapePlaneNormal3

part1.Update

	Set reference118 = part1.CreateReferenceFromObject(hybridShapeLineAngle5)

	Set reference119 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;25);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal4 = hybridShapeFactory1.AddNewPlaneNormal(reference118, reference119)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal4

	part1.InWorkObject = hybridShapePlaneNormal4

part1.Update

	Set reference120 = part1.CreateReferenceFromObject(hybridShapeLineAngle2)

	Set reference121 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;27);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal5 = hybridShapeFactory1.AddNewPlaneNormal(reference120, reference121)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal5

	part1.InWorkObject = hybridShapePlaneNormal5

part1.Update

	Set reference122 = part1.CreateReferenceFromObject(hybridShapeLineAngle3)

	Set reference123 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;29);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal6 = hybridShapeFactory1.AddNewPlaneNormal(reference122, reference123)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal6

	part1.InWorkObject = hybridShapePlaneNormal6

part1.Update

	Set reference124 = part1.CreateReferenceFromObject(hybridShapeLineAngle8)

	Set reference125 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;31);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal7 = hybridShapeFactory1.AddNewPlaneNormal(reference124, reference125)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal7

	part1.InWorkObject = hybridShapePlaneNormal7

part1.Update

	Set reference126 = part1.CreateReferenceFromObject(hybridShapeLineAngle7)

	Set reference127 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;33);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal8 = hybridShapeFactory1.AddNewPlaneNormal(reference126, reference127)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal8

	part1.InWorkObject = hybridShapePlaneNormal8

part1.Update

	Set reference128 = part1.CreateReferenceFromObject(hybridShapeLineAngle6)

	Set reference129 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;14);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set hybridShapePlaneNormal9 = hybridShapeFactory1.AddNewPlaneNormal(reference128, reference129)

	hybridBody1.AppendHybridShape hybridShapePlaneNormal9

	part1.InWorkObject = hybridShapePlaneNormal9

part1.Update

	Set reference130 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;13);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference131 = part1.CreateReferenceFromObject(hybridShapePlaneNormal2)

	Set hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference130, reference131, False, 1#)

	hybridShapeCircleCtrRad1.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad1

	part1.InWorkObject = hybridShapeCircleCtrRad1

part1.Update

	Set reference132 = part1.CreateReferenceFromObject(hybridShapePlaneNormal2)

	Set hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference130, reference132, False, 0.5)

	hybridShapeCircleCtrRad1.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad1

	part1.InWorkObject = hybridShapeCircleCtrRad1

part1.Update

	Set reference133 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;23);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference134 = part1.CreateReferenceFromObject(hybridShapePlaneNormal3)

	Set hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(reference133, reference134, False, 0.5)

	hybridShapeCircleCtrRad2.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad2

	part1.InWorkObject = hybridShapeCircleCtrRad2

part1.Update

	Set reference135 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;25);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference136 = part1.CreateReferenceFromObject(hybridShapePlaneNormal4)

	Set hybridShapeCircleCtrRad3 = hybridShapeFactory1.AddNewCircleCtrRad(reference135, reference136, False, 0.5)

	hybridShapeCircleCtrRad3.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad3

	part1.InWorkObject = hybridShapeCircleCtrRad3

part1.Update

	Set reference137 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;7);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference138 = part1.CreateReferenceFromObject(hybridShapePlaneNormal1)

	Set hybridShapeCircleCtrRad4 = hybridShapeFactory1.AddNewCircleCtrRad(reference137, reference138, False, 0.5)

	hybridShapeCircleCtrRad4.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad4

	part1.InWorkObject = hybridShapeCircleCtrRad4

part1.Update

	Set reference139 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;29);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference140 = part1.CreateReferenceFromObject(hybridShapePlaneNormal6)

	Set hybridShapeCircleCtrRad5 = hybridShapeFactory1.AddNewCircleCtrRad(reference139, reference140, False, 0.5)

	hybridShapeCircleCtrRad5.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad5

	part1.InWorkObject = hybridShapeCircleCtrRad5

part1.Update

	Set reference141 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;31);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference142 = part1.CreateReferenceFromObject(hybridShapePlaneNormal7)

	Set hybridShapeCircleCtrRad6 = hybridShapeFactory1.AddNewCircleCtrRad(reference141, reference142, False, 0.5)

	hybridShapeCircleCtrRad6.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad6

	part1.InWorkObject = hybridShapeCircleCtrRad6

part1.Update

	Set reference143 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;33);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference144 = part1.CreateReferenceFromObject(hybridShapePlaneNormal8)

	Set hybridShapeCircleCtrRad7 = hybridShapeFactory1.AddNewCircleCtrRad(reference143, reference144, False, 0.5)

	hybridShapeCircleCtrRad7.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad7

	part1.InWorkObject = hybridShapeCircleCtrRad7

part1.Update

	Set reference145 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;14);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference146 = part1.CreateReferenceFromObject(hybridShapePlaneNormal9)

	Set hybridShapeCircleCtrRad8 = hybridShapeFactory1.AddNewCircleCtrRad(reference145, reference146, False, 0.5)

	hybridShapeCircleCtrRad8.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad8

	part1.InWorkObject = hybridShapeCircleCtrRad8

part1.Update

	part1.InWorkObject = body1

	Set shapeFactory1 = part1.ShapeFactory

	Set reference147 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)

	Set reference148 = part1.CreateReferenceFromObject(hybridShapeLineAngle1)

	Set rib1 = shapeFactory1.AddNewRibFromRef(reference147, reference148)

part1.Update

	Set reference149 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)

	Set reference150 = part1.CreateReferenceFromObject(hybridShapeLineAngle4)

	Set rib2 = shapeFactory1.AddNewRibFromRef(reference149, reference150)

part1.Update

	Set reference151 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad3)

	Set reference152 = part1.CreateReferenceFromObject(hybridShapeLineAngle5)

	Set rib3 = shapeFactory1.AddNewRibFromRef(reference151, reference152)

part1.Update

	Set reference153 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad4)

	Set reference154 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

	Set rib4 = shapeFactory1.AddNewRibFromRef(reference153, reference154)

part1.Update

	part1.InWorkObject = hybridBody1

	Set reference155 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;27);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

	Set reference156 = part1.CreateReferenceFromObject(hybridShapePlaneNormal5)

	Set hybridShapeCircleCtrRad9 = hybridShapeFactory1.AddNewCircleCtrRad(reference155, reference156, False, 0.5)

	hybridShapeCircleCtrRad9.SetLimitation 1

	hybridBody1.AppendHybridShape hybridShapeCircleCtrRad9

	part1.InWorkObject = hybridShapeCircleCtrRad9

part1.Update

	part1.InWorkObject = body1

	Set reference157 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad9)

	Set reference158 = part1.CreateReferenceFromObject(hybridShapeLineAngle2)

	Set rib5 = shapeFactory1.AddNewRibFromRef(reference157, reference158)

part1.Update

	Set reference159 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad5)

	Set reference160 = part1.CreateReferenceFromObject(hybridShapeLineAngle3)

	Set rib6 = shapeFactory1.AddNewRibFromRef(reference159, reference160)

part1.Update

	Set reference161 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad6)

	Set reference162 = part1.CreateReferenceFromObject(hybridShapeLineAngle8)

	Set rib7 = shapeFactory1.AddNewRibFromRef(reference161, reference162)

part1.Update

	Set reference163 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad7)

	Set reference164 = part1.CreateReferenceFromObject(hybridShapeLineAngle7)

	Set rib8 = shapeFactory1.AddNewRibFromRef(reference163, reference164)

part1.Update

	Set reference165 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad8)

	Set reference166 = part1.CreateReferenceFromObject(hybridShapeLineAngle6)

	Set rib9 = shapeFactory1.AddNewRibFromRef(reference165, reference166)

part1.Update

End Sub

