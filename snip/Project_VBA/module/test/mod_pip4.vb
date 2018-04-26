Sub CATMain()

Dim windows1 As Windows
Set windows1 = CATIA.Windows

Dim specsAndGeomWindow1 As SpecsAndGeomWindow
Set specsAndGeomWindow1 = windows1.Item("Part1")

specsAndGeomWindow1.Activate

specsAndGeomWindow1.WindowState = catWindowStateMinimized

Dim specsAndGeomWindow2 As SpecsAndGeomWindow
Set specsAndGeomWindow2 = windows1.Item("Product1")

specsAndGeomWindow2.Activate

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

Dim products1 As Products
Set products1 = product1.Products

Dim product2 As Product
Set product2 = products1.AddNewComponent("Part", "y007_1")

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
arrayOfVariantOfDouble1(9) = 10#
arrayOfVariantOfDouble1(10) = 0#
arrayOfVariantOfDouble1(11) = 2.439021
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
arrayOfVariantOfDouble2(9) = 0.118726
arrayOfVariantOfDouble2(10) = 0#
arrayOfVariantOfDouble2(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble2

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble3(11)
arrayOfVariantOfDouble3(0) = 1#
arrayOfVariantOfDouble3(1) = 0#
arrayOfVariantOfDouble3(2) = 0#
arrayOfVariantOfDouble3(3) = 0#
arrayOfVariantOfDouble3(4) = 1#
arrayOfVariantOfDouble3(5) = 0#
arrayOfVariantOfDouble3(6) = 0#
arrayOfVariantOfDouble3(7) = 0#
arrayOfVariantOfDouble3(8) = 1#
arrayOfVariantOfDouble3(9) = 0.01445
arrayOfVariantOfDouble3(10) = 0#
arrayOfVariantOfDouble3(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble3

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble4(11)
arrayOfVariantOfDouble4(0) = 1#
arrayOfVariantOfDouble4(1) = 0#
arrayOfVariantOfDouble4(2) = 0#
arrayOfVariantOfDouble4(3) = 0#
arrayOfVariantOfDouble4(4) = 1#
arrayOfVariantOfDouble4(5) = 0#
arrayOfVariantOfDouble4(6) = 0#
arrayOfVariantOfDouble4(7) = 0#
arrayOfVariantOfDouble4(8) = 1#
arrayOfVariantOfDouble4(9) = 0.093957
arrayOfVariantOfDouble4(10) = 0#
arrayOfVariantOfDouble4(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble4

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble5(11)
arrayOfVariantOfDouble5(0) = 1#
arrayOfVariantOfDouble5(1) = 0#
arrayOfVariantOfDouble5(2) = 0#
arrayOfVariantOfDouble5(3) = 0#
arrayOfVariantOfDouble5(4) = 1#
arrayOfVariantOfDouble5(5) = 0#
arrayOfVariantOfDouble5(6) = 0#
arrayOfVariantOfDouble5(7) = 0#
arrayOfVariantOfDouble5(8) = 1#
arrayOfVariantOfDouble5(9) = 0.065058
arrayOfVariantOfDouble5(10) = 0#
arrayOfVariantOfDouble5(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble5

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble6(11)
arrayOfVariantOfDouble6(0) = 1#
arrayOfVariantOfDouble6(1) = 0#
arrayOfVariantOfDouble6(2) = 0#
arrayOfVariantOfDouble6(3) = 0#
arrayOfVariantOfDouble6(4) = 1#
arrayOfVariantOfDouble6(5) = 0#
arrayOfVariantOfDouble6(6) = 0#
arrayOfVariantOfDouble6(7) = 0#
arrayOfVariantOfDouble6(8) = 1#
arrayOfVariantOfDouble6(9) = 0.068653
arrayOfVariantOfDouble6(10) = 0#
arrayOfVariantOfDouble6(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble6

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble7(11)
arrayOfVariantOfDouble7(0) = 1#
arrayOfVariantOfDouble7(1) = 0#
arrayOfVariantOfDouble7(2) = 0#
arrayOfVariantOfDouble7(3) = 0#
arrayOfVariantOfDouble7(4) = 1#
arrayOfVariantOfDouble7(5) = 0#
arrayOfVariantOfDouble7(6) = 0#
arrayOfVariantOfDouble7(7) = 0#
arrayOfVariantOfDouble7(8) = 1#
arrayOfVariantOfDouble7(9) = 0.065062
arrayOfVariantOfDouble7(10) = 0#
arrayOfVariantOfDouble7(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble7

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble8(11)
arrayOfVariantOfDouble8(0) = 1#
arrayOfVariantOfDouble8(1) = 0#
arrayOfVariantOfDouble8(2) = 0#
arrayOfVariantOfDouble8(3) = 0#
arrayOfVariantOfDouble8(4) = 1#
arrayOfVariantOfDouble8(5) = 0#
arrayOfVariantOfDouble8(6) = 0#
arrayOfVariantOfDouble8(7) = 0#
arrayOfVariantOfDouble8(8) = 1#
arrayOfVariantOfDouble8(9) = 0.079508
arrayOfVariantOfDouble8(10) = 0#
arrayOfVariantOfDouble8(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble8

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble9(11)
arrayOfVariantOfDouble9(0) = 1#
arrayOfVariantOfDouble9(1) = 0#
arrayOfVariantOfDouble9(2) = 0#
arrayOfVariantOfDouble9(3) = 0#
arrayOfVariantOfDouble9(4) = 1#
arrayOfVariantOfDouble9(5) = 0#
arrayOfVariantOfDouble9(6) = 0#
arrayOfVariantOfDouble9(7) = 0#
arrayOfVariantOfDouble9(8) = 1#
arrayOfVariantOfDouble9(9) = 0.054203
arrayOfVariantOfDouble9(10) = 0#
arrayOfVariantOfDouble9(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble9

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble10(11)
arrayOfVariantOfDouble10(0) = 1#
arrayOfVariantOfDouble10(1) = 0#
arrayOfVariantOfDouble10(2) = 0#
arrayOfVariantOfDouble10(3) = 0#
arrayOfVariantOfDouble10(4) = 1#
arrayOfVariantOfDouble10(5) = 0#
arrayOfVariantOfDouble10(6) = 0#
arrayOfVariantOfDouble10(7) = 0#
arrayOfVariantOfDouble10(8) = 1#
arrayOfVariantOfDouble10(9) = 0.065058
arrayOfVariantOfDouble10(10) = 0#
arrayOfVariantOfDouble10(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble10

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble11(11)
arrayOfVariantOfDouble11(0) = 1#
arrayOfVariantOfDouble11(1) = 0#
arrayOfVariantOfDouble11(2) = 0#
arrayOfVariantOfDouble11(3) = 0#
arrayOfVariantOfDouble11(4) = 1#
arrayOfVariantOfDouble11(5) = 0#
arrayOfVariantOfDouble11(6) = 0#
arrayOfVariantOfDouble11(7) = 0#
arrayOfVariantOfDouble11(8) = 1#
arrayOfVariantOfDouble11(9) = 0.144572
arrayOfVariantOfDouble11(10) = 0#
arrayOfVariantOfDouble11(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble11

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble12(11)
arrayOfVariantOfDouble12(0) = 1#
arrayOfVariantOfDouble12(1) = 0#
arrayOfVariantOfDouble12(2) = 0#
arrayOfVariantOfDouble12(3) = 0#
arrayOfVariantOfDouble12(4) = 1#
arrayOfVariantOfDouble12(5) = 0#
arrayOfVariantOfDouble12(6) = 0#
arrayOfVariantOfDouble12(7) = 0#
arrayOfVariantOfDouble12(8) = 1#
arrayOfVariantOfDouble12(9) = 0.039754
arrayOfVariantOfDouble12(10) = 0#
arrayOfVariantOfDouble12(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble12

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble13(11)
arrayOfVariantOfDouble13(0) = 1#
arrayOfVariantOfDouble13(1) = 0#
arrayOfVariantOfDouble13(2) = 0#
arrayOfVariantOfDouble13(3) = 0#
arrayOfVariantOfDouble13(4) = 1#
arrayOfVariantOfDouble13(5) = 0#
arrayOfVariantOfDouble13(6) = 0#
arrayOfVariantOfDouble13(7) = 0#
arrayOfVariantOfDouble13(8) = 1#
arrayOfVariantOfDouble13(9) = 0.025306
arrayOfVariantOfDouble13(10) = 0#
arrayOfVariantOfDouble13(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble13

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble14(11)
arrayOfVariantOfDouble14(0) = 1#
arrayOfVariantOfDouble14(1) = 0#
arrayOfVariantOfDouble14(2) = 0#
arrayOfVariantOfDouble14(3) = 0#
arrayOfVariantOfDouble14(4) = 1#
arrayOfVariantOfDouble14(5) = 0#
arrayOfVariantOfDouble14(6) = 0#
arrayOfVariantOfDouble14(7) = 0#
arrayOfVariantOfDouble14(8) = 1#
arrayOfVariantOfDouble14(9) = 0.079509
arrayOfVariantOfDouble14(10) = 0#
arrayOfVariantOfDouble14(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble14

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble15(11)
arrayOfVariantOfDouble15(0) = 1#
arrayOfVariantOfDouble15(1) = 0#
arrayOfVariantOfDouble15(2) = 0#
arrayOfVariantOfDouble15(3) = 0#
arrayOfVariantOfDouble15(4) = 1#
arrayOfVariantOfDouble15(5) = 0#
arrayOfVariantOfDouble15(6) = 0#
arrayOfVariantOfDouble15(7) = 0#
arrayOfVariantOfDouble15(8) = 1#
arrayOfVariantOfDouble15(9) = 0.039753
arrayOfVariantOfDouble15(10) = 0#
arrayOfVariantOfDouble15(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble15

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble16(11)
arrayOfVariantOfDouble16(0) = 1#
arrayOfVariantOfDouble16(1) = 0#
arrayOfVariantOfDouble16(2) = 0#
arrayOfVariantOfDouble16(3) = 0#
arrayOfVariantOfDouble16(4) = 1#
arrayOfVariantOfDouble16(5) = 0#
arrayOfVariantOfDouble16(6) = 0#
arrayOfVariantOfDouble16(7) = 0#
arrayOfVariantOfDouble16(8) = 1#
arrayOfVariantOfDouble16(9) = 0.039755
arrayOfVariantOfDouble16(10) = 0#
arrayOfVariantOfDouble16(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble16

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble17(11)
arrayOfVariantOfDouble17(0) = 1#
arrayOfVariantOfDouble17(1) = 0#
arrayOfVariantOfDouble17(2) = 0#
arrayOfVariantOfDouble17(3) = 0#
arrayOfVariantOfDouble17(4) = 1#
arrayOfVariantOfDouble17(5) = 0#
arrayOfVariantOfDouble17(6) = 0#
arrayOfVariantOfDouble17(7) = 0#
arrayOfVariantOfDouble17(8) = 1#
arrayOfVariantOfDouble17(9) = 0.039755
arrayOfVariantOfDouble17(10) = 0#
arrayOfVariantOfDouble17(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble17

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble18(11)
arrayOfVariantOfDouble18(0) = 1#
arrayOfVariantOfDouble18(1) = 0#
arrayOfVariantOfDouble18(2) = 0#
arrayOfVariantOfDouble18(3) = 0#
arrayOfVariantOfDouble18(4) = 1#
arrayOfVariantOfDouble18(5) = 0#
arrayOfVariantOfDouble18(6) = 0#
arrayOfVariantOfDouble18(7) = 0#
arrayOfVariantOfDouble18(8) = 1#
arrayOfVariantOfDouble18(9) = 0.065059
arrayOfVariantOfDouble18(10) = 0#
arrayOfVariantOfDouble18(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble18

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble19(11)
arrayOfVariantOfDouble19(0) = 1#
arrayOfVariantOfDouble19(1) = 0#
arrayOfVariantOfDouble19(2) = 0#
arrayOfVariantOfDouble19(3) = 0#
arrayOfVariantOfDouble19(4) = 1#
arrayOfVariantOfDouble19(5) = 0#
arrayOfVariantOfDouble19(6) = 0#
arrayOfVariantOfDouble19(7) = 0#
arrayOfVariantOfDouble19(8) = 1#
arrayOfVariantOfDouble19(9) = 0.104813
arrayOfVariantOfDouble19(10) = 0#
arrayOfVariantOfDouble19(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble19

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble20(11)
arrayOfVariantOfDouble20(0) = 1#
arrayOfVariantOfDouble20(1) = 0#
arrayOfVariantOfDouble20(2) = 0#
arrayOfVariantOfDouble20(3) = 0#
arrayOfVariantOfDouble20(4) = 1#
arrayOfVariantOfDouble20(5) = 0#
arrayOfVariantOfDouble20(6) = 0#
arrayOfVariantOfDouble20(7) = 0#
arrayOfVariantOfDouble20(8) = 1#
arrayOfVariantOfDouble20(9) = 0.07951
arrayOfVariantOfDouble20(10) = 0#
arrayOfVariantOfDouble20(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble20

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble21(11)
arrayOfVariantOfDouble21(0) = 1#
arrayOfVariantOfDouble21(1) = 0#
arrayOfVariantOfDouble21(2) = 0#
arrayOfVariantOfDouble21(3) = 0#
arrayOfVariantOfDouble21(4) = 1#
arrayOfVariantOfDouble21(5) = 0#
arrayOfVariantOfDouble21(6) = 0#
arrayOfVariantOfDouble21(7) = 0#
arrayOfVariantOfDouble21(8) = 1#
arrayOfVariantOfDouble21(9) = 0.025306
arrayOfVariantOfDouble21(10) = 0#
arrayOfVariantOfDouble21(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble21

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble22(11)
arrayOfVariantOfDouble22(0) = 1#
arrayOfVariantOfDouble22(1) = 0#
arrayOfVariantOfDouble22(2) = 0#
arrayOfVariantOfDouble22(3) = 0#
arrayOfVariantOfDouble22(4) = 1#
arrayOfVariantOfDouble22(5) = 0#
arrayOfVariantOfDouble22(6) = 0#
arrayOfVariantOfDouble22(7) = 0#
arrayOfVariantOfDouble22(8) = 1#
arrayOfVariantOfDouble22(9) = 0.014447
arrayOfVariantOfDouble22(10) = 0#
arrayOfVariantOfDouble22(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble22

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble23(11)
arrayOfVariantOfDouble23(0) = 1#
arrayOfVariantOfDouble23(1) = 0#
arrayOfVariantOfDouble23(2) = 0#
arrayOfVariantOfDouble23(3) = 0#
arrayOfVariantOfDouble23(4) = 1#
arrayOfVariantOfDouble23(5) = 0#
arrayOfVariantOfDouble23(6) = 0#
arrayOfVariantOfDouble23(7) = 0#
arrayOfVariantOfDouble23(8) = 1#
arrayOfVariantOfDouble23(9) = 0.065061
arrayOfVariantOfDouble23(10) = 0#
arrayOfVariantOfDouble23(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble23

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble24(11)
arrayOfVariantOfDouble24(0) = 1#
arrayOfVariantOfDouble24(1) = 0#
arrayOfVariantOfDouble24(2) = 0#
arrayOfVariantOfDouble24(3) = 0#
arrayOfVariantOfDouble24(4) = 1#
arrayOfVariantOfDouble24(5) = 0#
arrayOfVariantOfDouble24(6) = 0#
arrayOfVariantOfDouble24(7) = 0#
arrayOfVariantOfDouble24(8) = 1#
arrayOfVariantOfDouble24(9) = 0.014449
arrayOfVariantOfDouble24(10) = 0#
arrayOfVariantOfDouble24(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble24

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble25(11)
arrayOfVariantOfDouble25(0) = 1#
arrayOfVariantOfDouble25(1) = 0#
arrayOfVariantOfDouble25(2) = 0#
arrayOfVariantOfDouble25(3) = 0#
arrayOfVariantOfDouble25(4) = 1#
arrayOfVariantOfDouble25(5) = 0#
arrayOfVariantOfDouble25(6) = 0#
arrayOfVariantOfDouble25(7) = 0#
arrayOfVariantOfDouble25(8) = 1#
arrayOfVariantOfDouble25(9) = 0.025305
arrayOfVariantOfDouble25(10) = 0#
arrayOfVariantOfDouble25(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble25

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble26(11)
arrayOfVariantOfDouble26(0) = 1#
arrayOfVariantOfDouble26(1) = 0#
arrayOfVariantOfDouble26(2) = 0#
arrayOfVariantOfDouble26(3) = 0#
arrayOfVariantOfDouble26(4) = 1#
arrayOfVariantOfDouble26(5) = 0#
arrayOfVariantOfDouble26(6) = 0#
arrayOfVariantOfDouble26(7) = 0#
arrayOfVariantOfDouble26(8) = 1#
arrayOfVariantOfDouble26(9) = 0.039755
arrayOfVariantOfDouble26(10) = 0#
arrayOfVariantOfDouble26(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble26

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble27(11)
arrayOfVariantOfDouble27(0) = 1#
arrayOfVariantOfDouble27(1) = 0#
arrayOfVariantOfDouble27(2) = 0#
arrayOfVariantOfDouble27(3) = 0#
arrayOfVariantOfDouble27(4) = 1#
arrayOfVariantOfDouble27(5) = 0#
arrayOfVariantOfDouble27(6) = 0#
arrayOfVariantOfDouble27(7) = 0#
arrayOfVariantOfDouble27(8) = 1#
arrayOfVariantOfDouble27(9) = 0.025305
arrayOfVariantOfDouble27(10) = 0#
arrayOfVariantOfDouble27(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble27

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble28(11)
arrayOfVariantOfDouble28(0) = 1#
arrayOfVariantOfDouble28(1) = 0#
arrayOfVariantOfDouble28(2) = 0#
arrayOfVariantOfDouble28(3) = 0#
arrayOfVariantOfDouble28(4) = 1#
arrayOfVariantOfDouble28(5) = 0#
arrayOfVariantOfDouble28(6) = 0#
arrayOfVariantOfDouble28(7) = 0#
arrayOfVariantOfDouble28(8) = 1#
arrayOfVariantOfDouble28(9) = 0.039754
arrayOfVariantOfDouble28(10) = 0#
arrayOfVariantOfDouble28(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble28

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble29(11)
arrayOfVariantOfDouble29(0) = 1#
arrayOfVariantOfDouble29(1) = 0#
arrayOfVariantOfDouble29(2) = 0#
arrayOfVariantOfDouble29(3) = 0#
arrayOfVariantOfDouble29(4) = 1#
arrayOfVariantOfDouble29(5) = 0#
arrayOfVariantOfDouble29(6) = 0#
arrayOfVariantOfDouble29(7) = 0#
arrayOfVariantOfDouble29(8) = 1#
arrayOfVariantOfDouble29(9) = 0.01445
arrayOfVariantOfDouble29(10) = 0#
arrayOfVariantOfDouble29(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble29

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble30(11)
arrayOfVariantOfDouble30(0) = 1#
arrayOfVariantOfDouble30(1) = 0#
arrayOfVariantOfDouble30(2) = 0#
arrayOfVariantOfDouble30(3) = 0#
arrayOfVariantOfDouble30(4) = 1#
arrayOfVariantOfDouble30(5) = 0#
arrayOfVariantOfDouble30(6) = 0#
arrayOfVariantOfDouble30(7) = 0#
arrayOfVariantOfDouble30(8) = 1#
arrayOfVariantOfDouble30(9) = 0.050612
arrayOfVariantOfDouble30(10) = 0#
arrayOfVariantOfDouble30(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble30

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble31(11)
arrayOfVariantOfDouble31(0) = 1#
arrayOfVariantOfDouble31(1) = 0#
arrayOfVariantOfDouble31(2) = 0#
arrayOfVariantOfDouble31(3) = 0#
arrayOfVariantOfDouble31(4) = 1#
arrayOfVariantOfDouble31(5) = 0#
arrayOfVariantOfDouble31(6) = 0#
arrayOfVariantOfDouble31(7) = 0#
arrayOfVariantOfDouble31(8) = 1#
arrayOfVariantOfDouble31(9) = 0.014449
arrayOfVariantOfDouble31(10) = 0#
arrayOfVariantOfDouble31(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble31

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble32(11)
arrayOfVariantOfDouble32(0) = 1#
arrayOfVariantOfDouble32(1) = 0#
arrayOfVariantOfDouble32(2) = 0#
arrayOfVariantOfDouble32(3) = 0#
arrayOfVariantOfDouble32(4) = 1#
arrayOfVariantOfDouble32(5) = 0#
arrayOfVariantOfDouble32(6) = 0#
arrayOfVariantOfDouble32(7) = 0#
arrayOfVariantOfDouble32(8) = 1#
arrayOfVariantOfDouble32(9) = 0.025306
arrayOfVariantOfDouble32(10) = 0#
arrayOfVariantOfDouble32(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble32

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble33(11)
arrayOfVariantOfDouble33(0) = 1#
arrayOfVariantOfDouble33(1) = 0#
arrayOfVariantOfDouble33(2) = 0#
arrayOfVariantOfDouble33(3) = 0#
arrayOfVariantOfDouble33(4) = 1#
arrayOfVariantOfDouble33(5) = 0#
arrayOfVariantOfDouble33(6) = 0#
arrayOfVariantOfDouble33(7) = 0#
arrayOfVariantOfDouble33(8) = 1#
arrayOfVariantOfDouble33(9) = 0.014448
arrayOfVariantOfDouble33(10) = 0#
arrayOfVariantOfDouble33(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble33

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble34(11)
arrayOfVariantOfDouble34(0) = 1#
arrayOfVariantOfDouble34(1) = 0#
arrayOfVariantOfDouble34(2) = 0#
arrayOfVariantOfDouble34(3) = 0#
arrayOfVariantOfDouble34(4) = 1#
arrayOfVariantOfDouble34(5) = 0#
arrayOfVariantOfDouble34(6) = 0#
arrayOfVariantOfDouble34(7) = 0#
arrayOfVariantOfDouble34(8) = 1#
arrayOfVariantOfDouble34(9) = 0.06506
arrayOfVariantOfDouble34(10) = 0#
arrayOfVariantOfDouble34(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble34

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble35(11)
arrayOfVariantOfDouble35(0) = 1#
arrayOfVariantOfDouble35(1) = 0#
arrayOfVariantOfDouble35(2) = 0#
arrayOfVariantOfDouble35(3) = 0#
arrayOfVariantOfDouble35(4) = 1#
arrayOfVariantOfDouble35(5) = 0#
arrayOfVariantOfDouble35(6) = 0#
arrayOfVariantOfDouble35(7) = 0#
arrayOfVariantOfDouble35(8) = 1#
arrayOfVariantOfDouble35(9) = 0.039754
arrayOfVariantOfDouble35(10) = 0#
arrayOfVariantOfDouble35(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble35

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble36(11)
arrayOfVariantOfDouble36(0) = 1#
arrayOfVariantOfDouble36(1) = 0#
arrayOfVariantOfDouble36(2) = 0#
arrayOfVariantOfDouble36(3) = 0#
arrayOfVariantOfDouble36(4) = 1#
arrayOfVariantOfDouble36(5) = 0#
arrayOfVariantOfDouble36(6) = 0#
arrayOfVariantOfDouble36(7) = 0#
arrayOfVariantOfDouble36(8) = 1#
arrayOfVariantOfDouble36(9) = 0.104816
arrayOfVariantOfDouble36(10) = 0#
arrayOfVariantOfDouble36(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble36

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble37(11)
arrayOfVariantOfDouble37(0) = 1#
arrayOfVariantOfDouble37(1) = 0#
arrayOfVariantOfDouble37(2) = 0#
arrayOfVariantOfDouble37(3) = 0#
arrayOfVariantOfDouble37(4) = 1#
arrayOfVariantOfDouble37(5) = 0#
arrayOfVariantOfDouble37(6) = 0#
arrayOfVariantOfDouble37(7) = 0#
arrayOfVariantOfDouble37(8) = 1#
arrayOfVariantOfDouble37(9) = 0.025305
arrayOfVariantOfDouble37(10) = 0#
arrayOfVariantOfDouble37(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble37

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble38(11)
arrayOfVariantOfDouble38(0) = 1#
arrayOfVariantOfDouble38(1) = 0#
arrayOfVariantOfDouble38(2) = 0#
arrayOfVariantOfDouble38(3) = 0#
arrayOfVariantOfDouble38(4) = 1#
arrayOfVariantOfDouble38(5) = 0#
arrayOfVariantOfDouble38(6) = 0#
arrayOfVariantOfDouble38(7) = 0#
arrayOfVariantOfDouble38(8) = 1#
arrayOfVariantOfDouble38(9) = 0.014448
arrayOfVariantOfDouble38(10) = 0#
arrayOfVariantOfDouble38(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble38

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble39(11)
arrayOfVariantOfDouble39(0) = 1#
arrayOfVariantOfDouble39(1) = 0#
arrayOfVariantOfDouble39(2) = 0#
arrayOfVariantOfDouble39(3) = 0#
arrayOfVariantOfDouble39(4) = 1#
arrayOfVariantOfDouble39(5) = 0#
arrayOfVariantOfDouble39(6) = 0#
arrayOfVariantOfDouble39(7) = 0#
arrayOfVariantOfDouble39(8) = 1#
arrayOfVariantOfDouble39(9) = 0.025305
arrayOfVariantOfDouble39(10) = 0#
arrayOfVariantOfDouble39(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble39

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble40(11)
arrayOfVariantOfDouble40(0) = 1#
arrayOfVariantOfDouble40(1) = 0#
arrayOfVariantOfDouble40(2) = 0#
arrayOfVariantOfDouble40(3) = 0#
arrayOfVariantOfDouble40(4) = 1#
arrayOfVariantOfDouble40(5) = 0#
arrayOfVariantOfDouble40(6) = 0#
arrayOfVariantOfDouble40(7) = 0#
arrayOfVariantOfDouble40(8) = 1#
arrayOfVariantOfDouble40(9) = 0.039755
arrayOfVariantOfDouble40(10) = 0#
arrayOfVariantOfDouble40(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble40

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble41(11)
arrayOfVariantOfDouble41(0) = 1#
arrayOfVariantOfDouble41(1) = 0#
arrayOfVariantOfDouble41(2) = 0#
arrayOfVariantOfDouble41(3) = 0#
arrayOfVariantOfDouble41(4) = 1#
arrayOfVariantOfDouble41(5) = 0#
arrayOfVariantOfDouble41(6) = 0#
arrayOfVariantOfDouble41(7) = 0#
arrayOfVariantOfDouble41(8) = 1#
arrayOfVariantOfDouble41(9) = 0.039755
arrayOfVariantOfDouble41(10) = 0#
arrayOfVariantOfDouble41(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble41

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble42(11)
arrayOfVariantOfDouble42(0) = 1#
arrayOfVariantOfDouble42(1) = 0#
arrayOfVariantOfDouble42(2) = 0#
arrayOfVariantOfDouble42(3) = 0#
arrayOfVariantOfDouble42(4) = 1#
arrayOfVariantOfDouble42(5) = 0#
arrayOfVariantOfDouble42(6) = 0#
arrayOfVariantOfDouble42(7) = 0#
arrayOfVariantOfDouble42(8) = 1#
arrayOfVariantOfDouble42(9) = 0.06506
arrayOfVariantOfDouble42(10) = 0#
arrayOfVariantOfDouble42(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble42

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble43(11)
arrayOfVariantOfDouble43(0) = 1#
arrayOfVariantOfDouble43(1) = 0#
arrayOfVariantOfDouble43(2) = 0#
arrayOfVariantOfDouble43(3) = 0#
arrayOfVariantOfDouble43(4) = 1#
arrayOfVariantOfDouble43(5) = 0#
arrayOfVariantOfDouble43(6) = 0#
arrayOfVariantOfDouble43(7) = 0#
arrayOfVariantOfDouble43(8) = 1#
arrayOfVariantOfDouble43(9) = 0.025304
arrayOfVariantOfDouble43(10) = 0#
arrayOfVariantOfDouble43(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble43

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble44(11)
arrayOfVariantOfDouble44(0) = 1#
arrayOfVariantOfDouble44(1) = 0#
arrayOfVariantOfDouble44(2) = 0#
arrayOfVariantOfDouble44(3) = 0#
arrayOfVariantOfDouble44(4) = 1#
arrayOfVariantOfDouble44(5) = 0#
arrayOfVariantOfDouble44(6) = 0#
arrayOfVariantOfDouble44(7) = 0#
arrayOfVariantOfDouble44(8) = 1#
arrayOfVariantOfDouble44(9) = 0.014449
arrayOfVariantOfDouble44(10) = 0#
arrayOfVariantOfDouble44(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble44

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble45(11)
arrayOfVariantOfDouble45(0) = 1#
arrayOfVariantOfDouble45(1) = 0#
arrayOfVariantOfDouble45(2) = 0#
arrayOfVariantOfDouble45(3) = 0#
arrayOfVariantOfDouble45(4) = 1#
arrayOfVariantOfDouble45(5) = 0#
arrayOfVariantOfDouble45(6) = 0#
arrayOfVariantOfDouble45(7) = 0#
arrayOfVariantOfDouble45(8) = 1#
arrayOfVariantOfDouble45(9) = 0.065058
arrayOfVariantOfDouble45(10) = 0#
arrayOfVariantOfDouble45(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble45

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble46(11)
arrayOfVariantOfDouble46(0) = 1#
arrayOfVariantOfDouble46(1) = 0#
arrayOfVariantOfDouble46(2) = 0#
arrayOfVariantOfDouble46(3) = 0#
arrayOfVariantOfDouble46(4) = 1#
arrayOfVariantOfDouble46(5) = 0#
arrayOfVariantOfDouble46(6) = 0#
arrayOfVariantOfDouble46(7) = 0#
arrayOfVariantOfDouble46(8) = 1#
arrayOfVariantOfDouble46(9) = 0.039756
arrayOfVariantOfDouble46(10) = 0#
arrayOfVariantOfDouble46(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble46

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble47(11)
arrayOfVariantOfDouble47(0) = 1#
arrayOfVariantOfDouble47(1) = 0#
arrayOfVariantOfDouble47(2) = 0#
arrayOfVariantOfDouble47(3) = 0#
arrayOfVariantOfDouble47(4) = 1#
arrayOfVariantOfDouble47(5) = 0#
arrayOfVariantOfDouble47(6) = 0#
arrayOfVariantOfDouble47(7) = 0#
arrayOfVariantOfDouble47(8) = 1#
arrayOfVariantOfDouble47(9) = 0.065059
arrayOfVariantOfDouble47(10) = 0#
arrayOfVariantOfDouble47(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble47

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble48(11)
arrayOfVariantOfDouble48(0) = 1#
arrayOfVariantOfDouble48(1) = 0#
arrayOfVariantOfDouble48(2) = 0#
arrayOfVariantOfDouble48(3) = 0#
arrayOfVariantOfDouble48(4) = 1#
arrayOfVariantOfDouble48(5) = 0#
arrayOfVariantOfDouble48(6) = 0#
arrayOfVariantOfDouble48(7) = 0#
arrayOfVariantOfDouble48(8) = 1#
arrayOfVariantOfDouble48(9) = 0.01445
arrayOfVariantOfDouble48(10) = 0#
arrayOfVariantOfDouble48(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble48

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble49(11)
arrayOfVariantOfDouble49(0) = 1#
arrayOfVariantOfDouble49(1) = 0#
arrayOfVariantOfDouble49(2) = 0#
arrayOfVariantOfDouble49(3) = 0#
arrayOfVariantOfDouble49(4) = 1#
arrayOfVariantOfDouble49(5) = 0#
arrayOfVariantOfDouble49(6) = 0#
arrayOfVariantOfDouble49(7) = 0#
arrayOfVariantOfDouble49(8) = 1#
arrayOfVariantOfDouble49(9) = 0.025305
arrayOfVariantOfDouble49(10) = 0#
arrayOfVariantOfDouble49(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble49

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble50(11)
arrayOfVariantOfDouble50(0) = 1#
arrayOfVariantOfDouble50(1) = 0#
arrayOfVariantOfDouble50(2) = 0#
arrayOfVariantOfDouble50(3) = 0#
arrayOfVariantOfDouble50(4) = 1#
arrayOfVariantOfDouble50(5) = 0#
arrayOfVariantOfDouble50(6) = 0#
arrayOfVariantOfDouble50(7) = 0#
arrayOfVariantOfDouble50(8) = 1#
arrayOfVariantOfDouble50(9) = 0.039754
arrayOfVariantOfDouble50(10) = 0#
arrayOfVariantOfDouble50(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble50

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble51(11)
arrayOfVariantOfDouble51(0) = 1#
arrayOfVariantOfDouble51(1) = 0#
arrayOfVariantOfDouble51(2) = 0#
arrayOfVariantOfDouble51(3) = 0#
arrayOfVariantOfDouble51(4) = 1#
arrayOfVariantOfDouble51(5) = 0#
arrayOfVariantOfDouble51(6) = 0#
arrayOfVariantOfDouble51(7) = 0#
arrayOfVariantOfDouble51(8) = 1#
arrayOfVariantOfDouble51(9) = 0.014448
arrayOfVariantOfDouble51(10) = 0#
arrayOfVariantOfDouble51(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble51

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble52(11)
arrayOfVariantOfDouble52(0) = 1#
arrayOfVariantOfDouble52(1) = 0#
arrayOfVariantOfDouble52(2) = 0#
arrayOfVariantOfDouble52(3) = 0#
arrayOfVariantOfDouble52(4) = 1#
arrayOfVariantOfDouble52(5) = 0#
arrayOfVariantOfDouble52(6) = 0#
arrayOfVariantOfDouble52(7) = 0#
arrayOfVariantOfDouble52(8) = 1#
arrayOfVariantOfDouble52(9) = 0.06506
arrayOfVariantOfDouble52(10) = 0#
arrayOfVariantOfDouble52(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble52

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble53(11)
arrayOfVariantOfDouble53(0) = 1#
arrayOfVariantOfDouble53(1) = 0#
arrayOfVariantOfDouble53(2) = 0#
arrayOfVariantOfDouble53(3) = 0#
arrayOfVariantOfDouble53(4) = 1#
arrayOfVariantOfDouble53(5) = 0#
arrayOfVariantOfDouble53(6) = 0#
arrayOfVariantOfDouble53(7) = 0#
arrayOfVariantOfDouble53(8) = 1#
arrayOfVariantOfDouble53(9) = 0.039755
arrayOfVariantOfDouble53(10) = 0#
arrayOfVariantOfDouble53(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble53

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble54(11)
arrayOfVariantOfDouble54(0) = 1#
arrayOfVariantOfDouble54(1) = 0#
arrayOfVariantOfDouble54(2) = 0#
arrayOfVariantOfDouble54(3) = 0#
arrayOfVariantOfDouble54(4) = 1#
arrayOfVariantOfDouble54(5) = 0#
arrayOfVariantOfDouble54(6) = 0#
arrayOfVariantOfDouble54(7) = 0#
arrayOfVariantOfDouble54(8) = 1#
arrayOfVariantOfDouble54(9) = 0.039756
arrayOfVariantOfDouble54(10) = 0#
arrayOfVariantOfDouble54(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble54

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble55(11)
arrayOfVariantOfDouble55(0) = 1#
arrayOfVariantOfDouble55(1) = 0#
arrayOfVariantOfDouble55(2) = 0#
arrayOfVariantOfDouble55(3) = 0#
arrayOfVariantOfDouble55(4) = 1#
arrayOfVariantOfDouble55(5) = 0#
arrayOfVariantOfDouble55(6) = 0#
arrayOfVariantOfDouble55(7) = 0#
arrayOfVariantOfDouble55(8) = 1#
arrayOfVariantOfDouble55(9) = 0.025304
arrayOfVariantOfDouble55(10) = 0#
arrayOfVariantOfDouble55(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble55

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Item("y007_1.CATPart")

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

Dim sketches1 As Sketches
Set sketches1 = hybridBody1.HybridSketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Sketch.1")

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("AbsoluteAxis")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("HDirection")

line2D1.ReportName = 1

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("VDirection")

line2D2.ReportName = 2

Dim circle2D1 As Circle2D
Set circle2D1 = factory2D1.CreateClosedCircle(0#, 0#, 10#)

Dim point2D1 As Point2D
Set point2D1 = axis2D1.GetItem("Origin")

circle2D1.CenterPoint = point2D1

circle2D1.ReportName = 3

circle2D1.Construction = True

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddMonoEltCst(catCstTypeRadius, reference1)

constraint1.Mode = catCstModeDrivenDimension

Dim hybridBody2 As HybridBody
Set hybridBody2 = hybridBodies1.Item("External References")

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody2.HybridShapes

Dim hybridShapeCurveExplicit1 As HybridShapeCurveExplicit
Set hybridShapeCurveExplicit1 = hybridShapes1.Item("Curve.1")

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(GSMMonoDim.1;[1,2]);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MonoFond;MFBRepVersion_CXR15)", hybridShapeCurveExplicit1)

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateProjections(reference2)

Dim point2D2 As Point2D
Set point2D2 = factory2D1.CreatePoint(-0.414214, 0#)

point2D2.ReportName = 4

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(-10#, 0#)

point2D3.ReportName = 5

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(-0.414214, 0#, -10#, 0#)

line2D3.ReportName = 6

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D3

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(point2D3)

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeOn, reference3, reference4)

constraint2.Mode = catCstModeDrivingDimension

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D3)

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromObject(line2D1)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddBiEltCst(catCstTypeHorizontality, reference5, reference6)

constraint3.Mode = catCstModeDrivingDimension

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(0.292893, -0.292893)

point2D4.ReportName = 7

Dim point2D5 As Point2D
Set point2D5 = factory2D1.CreatePoint(7.071068, -7.071068)

point2D5.ReportName = 8

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(0.292893, -0.292893, 7.071068, -7.071068)

line2D4.ReportName = 9

line2D4.StartPoint = point2D4

line2D4.EndPoint = point2D5

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(point2D5)

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint4 As Constraint
Set constraint4 = constraints1.AddBiEltCst(catCstTypeOn, reference7, reference8)

constraint4.Mode = catCstModeDrivingDimension

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(line2D4)

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromObject(line2D1)

Dim constraint5 As Constraint
Set constraint5 = constraints1.AddBiEltCst(catCstTypeAngle, reference9, reference10)

constraint5.Mode = catCstModeDrivingDimension

constraint5.AngleSector = catCstAngleSector0

Dim angle1 As Angle
Set angle1 = constraint5.Dimension

angle1.Value = 45#

Dim point2D6 As Point2D
Set point2D6 = factory2D1.CreatePoint(-0.414214, -1#)

point2D6.ReportName = 10

Dim circle2D2 As Circle2D
Set circle2D2 = factory2D1.CreateCircle(-0.414214, -1#, 1#, 0.785398, 1.570796)

circle2D2.CenterPoint = point2D6

circle2D2.ReportName = 11

circle2D2.StartPoint = point2D4

circle2D2.EndPoint = point2D2

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromObject(line2D4)

Dim reference12 As Reference
Set reference12 = part1.CreateReferenceFromObject(point2D1)

Dim constraint6 As Constraint
Set constraint6 = constraints1.AddBiEltCst(catCstTypeOn, reference11, reference12)

constraint6.Mode = catCstModeDrivingDimension

Dim reference13 As Reference
Set reference13 = part1.CreateReferenceFromObject(line2D3)

Dim reference14 As Reference
Set reference14 = part1.CreateReferenceFromObject(point2D1)

Dim constraint7 As Constraint
Set constraint7 = constraints1.AddBiEltCst(catCstTypeOn, reference13, reference14)

constraint7.Mode = catCstModeDrivingDimension

Dim reference15 As Reference
Set reference15 = part1.CreateReferenceFromObject(circle2D2)

Dim reference16 As Reference
Set reference16 = part1.CreateReferenceFromObject(line2D4)

Dim constraint8 As Constraint
Set constraint8 = constraints1.AddBiEltCst(catCstTypeTangency, reference15, reference16)

constraint8.Mode = catCstModeDrivingDimension

Dim reference17 As Reference
Set reference17 = part1.CreateReferenceFromObject(circle2D2)

Dim reference18 As Reference
Set reference18 = part1.CreateReferenceFromObject(line2D3)

Dim constraint9 As Constraint
Set constraint9 = constraints1.AddBiEltCst(catCstTypeTangency, reference17, reference18)

constraint9.Mode = catCstModeDrivingDimension

Dim reference19 As Reference
Set reference19 = part1.CreateReferenceFromObject(circle2D2)

Dim constraint10 As Constraint
Set constraint10 = constraints1.AddMonoEltCst(catCstTypeRadius, reference19)

constraint10.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint10.Dimension

length1.Value = 1#

sketch1.CloseEdition

part1.InWorkObject = hybridBody1

part1.UpdateObject sketch1

part1.InWorkObject = sketch1

Set factory2D1 = sketch1.OpenEdition()

sketch1.CloseEdition

part1.InWorkObject = hybridBody1

part1.UpdateObject sketch1

Dim hybridShapes2 As HybridShapes
Set hybridShapes2 = hybridBody1.HybridShapes

Dim hybridShapeLineExplicit1 As HybridShapeLineExplicit
Set hybridShapeLineExplicit1 = hybridShapes2.Item("Z Axis")

Dim reference20 As Reference
Set reference20 = part1.CreateReferenceFromObject(hybridShapeLineExplicit1)

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim reference21 As Reference
Set reference21 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;5);None:(Limits1:();Limits2:();-1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapeCircleCenterAxis1 As HybridShapeCircleCenterAxis
Set hybridShapeCircleCenterAxis1 = hybridShapeFactory1.AddNewCircleCenterAxis(reference20, reference21, 1#, True)

hybridShapeCircleCenterAxis1.SetLimitation 1

hybridBody1.AppendHybridShape hybridShapeCircleCenterAxis1

part1.InWorkObject = hybridShapeCircleCenterAxis1

part1.Update

Dim reference22 As Reference
Set reference22 = part1.CreateReferenceFromObject(hybridShapeLineExplicit1)

Set hybridShapeCircleCenterAxis1 = hybridShapeFactory1.AddNewCircleCenterAxis(reference22, reference21, 0.5, True)

hybridShapeCircleCenterAxis1.SetLimitation 1

hybridBody1.AppendHybridShape hybridShapeCircleCenterAxis1

part1.InWorkObject = hybridShapeCircleCenterAxis1

part1.Update

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("PartBody")

part1.InWorkObject = body1

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim reference23 As Reference
Set reference23 = part1.CreateReferenceFromObject(hybridShapeCircleCenterAxis1)

Dim reference24 As Reference
Set reference24 = part1.CreateReferenceFromObject(sketch1)

Dim rib1 As Rib
Set rib1 = shapeFactory1.AddNewRibFromRef(reference23, reference24)

rib1.MergeMode = catMergeOn

rib1.MergeMode = catMergeOff

part1.UpdateObject rib1

part1.UpdateObject rib1

Dim product3 As Product
Set product3 = products1.Item("Part1.1")

Dim publications1 As Publications
Set publications1 = product3.Publications

Dim long1 As Long
long1 = publications1.Count

Dim publications2 As Publications
Set publications2 = product1.Publications

Dim long2 As Long
long2 = publications2.Count

Dim product4 As Product
Set product4 = product2.ReferenceProduct

Set publications1 = product3.Publications

Dim long3 As Long
long3 = publications1.Count

Set publications2 = product1.Publications

Dim long4 As Long
long4 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications1 = product3.Publications

Dim long5 As Long
long5 = publications1.Count

Set publications2 = product1.Publications

Dim long6 As Long
long6 = publications2.Count

Set product4 = product2.ReferenceProduct

Dim product5 As Product
Set product5 = products1.Item("part2.1")

Dim publications3 As Publications
Set publications3 = product5.Publications

Dim long7 As Long
long7 = publications3.Count

Set publications2 = product1.Publications

Dim long8 As Long
long8 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long9 As Long
long9 = publications3.Count

Set publications2 = product1.Publications

Dim long10 As Long
long10 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long11 As Long
long11 = publications3.Count

Set publications2 = product1.Publications

Dim long12 As Long
long12 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long13 As Long
long13 = publications3.Count

Set publications2 = product1.Publications

Dim long14 As Long
long14 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long15 As Long
long15 = publications3.Count

Set publications2 = product1.Publications

Dim long16 As Long
long16 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long17 As Long
long17 = publications3.Count

Set publications2 = product1.Publications

Dim long18 As Long
long18 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long19 As Long
long19 = publications3.Count

Set publications2 = product1.Publications

Dim long20 As Long
long20 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long21 As Long
long21 = publications3.Count

Set publications2 = product1.Publications

Dim long22 As Long
long22 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long23 As Long
long23 = publications3.Count

Set publications2 = product1.Publications

Dim long24 As Long
long24 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications3 = product5.Publications

Dim long25 As Long
long25 = publications3.Count

Set publications2 = product1.Publications

Dim long26 As Long
long26 = publications2.Count

Set product4 = product2.ReferenceProduct

Dim parameters1 As Parameters
Set parameters1 = part1.Parameters

Dim length2 As Length
Set length2 = parameters1.CreateDimension("", "LENGTH", 0#)

length2.Rename "supdiam"

length2.Value = 1#

Set publications1 = product3.Publications

Dim long27 As Long
long27 = publications1.Count

Set publications2 = product1.Publications

Dim long28 As Long
long28 = publications2.Count

Set product4 = product2.ReferenceProduct

Dim product6 As Product
Set product6 = products1.Item("Part24.1")

Dim publications4 As Publications
Set publications4 = product6.Publications

Dim long29 As Long
long29 = publications4.Count

Set publications2 = product1.Publications

Dim long30 As Long
long30 = publications2.Count

Set product4 = product2.ReferenceProduct

Dim relations1 As Relations
Set relations1 = part1.Relations

Dim parameters2 As Parameters
Set parameters2 = part1.Parameters

Dim length3 As Length
Set length3 = parameters2.Item("y007_1\Geometrical Set.1\Circle.1\Radius")

Dim formula1 As Formula
Set formula1 = relations1.CreateFormula("Formula.2", "", length3, "supdiam /2")

formula1.Rename "Formula.2"

Dim product7 As Product
Set product7 = products1.Item("y005 .7055475.1")

Dim publications5 As Publications
Set publications5 = product7.Publications

Dim long31 As Long
long31 = publications5.Count

Set publications2 = product1.Publications

Dim long32 As Long
long32 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications5 = product7.Publications

Dim long33 As Long
long33 = publications5.Count

Set publications2 = product1.Publications

Dim long34 As Long
long34 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications5 = product7.Publications

Dim long35 As Long
long35 = publications5.Count

Set publications2 = product1.Publications

Dim long36 As Long
long36 = publications2.Count

Set product4 = product2.ReferenceProduct

Set publications5 = product7.Publications

Dim long37 As Long
long37 = publications5.Count

Set publications2 = product1.Publications

Dim long38 As Long
long38 = publications2.Count

Set product4 = product2.ReferenceProduct

part1.Update

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble56(11)
arrayOfVariantOfDouble56(0) = 1#
arrayOfVariantOfDouble56(1) = 0#
arrayOfVariantOfDouble56(2) = 0#
arrayOfVariantOfDouble56(3) = 0#
arrayOfVariantOfDouble56(4) = 1#
arrayOfVariantOfDouble56(5) = 0#
arrayOfVariantOfDouble56(6) = 0#
arrayOfVariantOfDouble56(7) = 0#
arrayOfVariantOfDouble56(8) = 1#
arrayOfVariantOfDouble56(9) = 0#
arrayOfVariantOfDouble56(10) = 0#
arrayOfVariantOfDouble56(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble56

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble57(11)
arrayOfVariantOfDouble57(0) = 1#
arrayOfVariantOfDouble57(1) = 0#
arrayOfVariantOfDouble57(2) = 0#
arrayOfVariantOfDouble57(3) = 0#
arrayOfVariantOfDouble57(4) = 1#
arrayOfVariantOfDouble57(5) = 0#
arrayOfVariantOfDouble57(6) = 0#
arrayOfVariantOfDouble57(7) = 0#
arrayOfVariantOfDouble57(8) = 1#
arrayOfVariantOfDouble57(9) = -2#
arrayOfVariantOfDouble57(10) = 0#
arrayOfVariantOfDouble57(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble57

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble58(11)
arrayOfVariantOfDouble58(0) = 1#
arrayOfVariantOfDouble58(1) = 0#
arrayOfVariantOfDouble58(2) = 0#
arrayOfVariantOfDouble58(3) = 0#
arrayOfVariantOfDouble58(4) = 1#
arrayOfVariantOfDouble58(5) = 0#
arrayOfVariantOfDouble58(6) = 0#
arrayOfVariantOfDouble58(7) = 0#
arrayOfVariantOfDouble58(8) = 1#
arrayOfVariantOfDouble58(9) = 0#
arrayOfVariantOfDouble58(10) = 0#
arrayOfVariantOfDouble58(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble58

Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble59(11)
arrayOfVariantOfDouble59(0) = 1#
arrayOfVariantOfDouble59(1) = 0#
arrayOfVariantOfDouble59(2) = 0#
arrayOfVariantOfDouble59(3) = 0#
arrayOfVariantOfDouble59(4) = 1#
arrayOfVariantOfDouble59(5) = 0#
arrayOfVariantOfDouble59(6) = 0#
arrayOfVariantOfDouble59(7) = 0#
arrayOfVariantOfDouble59(8) = 1#
arrayOfVariantOfDouble59(9) = -1#
arrayOfVariantOfDouble59(10) = 0#
arrayOfVariantOfDouble59(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble59

product1.Update

Dim hybridShapePlaneExplicit1 As HybridShapePlaneExplicit
Set hybridShapePlaneExplicit1 = hybridShapes1.Item("Plane.2")

Dim reference25 As Reference
Set reference25 = part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)

Dim mirror1 As Mirror
Set mirror1 = shapeFactory1.AddNewMirror(reference25)

part1.UpdateObject mirror1

End Sub

