Private Sub cmdCreerPoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreerPoint.Click
        Dim CATIA As INFITF.Application
        CATIA = GetObject(, "CATIA.Application")

        Dim MyPartDoc As MECMOD.PartDocument = CATIA.ActiveDocument

        Dim MyPart As MECMOD.Part = MyPartDoc.Part
        Dim MyHB As MECMOD.HybridBody = MyPart.HybridBodies.Add()
        Dim MyHSF As HybridShapeFactory = MyPart.HybridShapeFactory
        Dim MyP As HybridShapePointCoord = MyHSF.AddNewPointCoord(TrackBar1.Value, TrackBar2.Value, TrackBar3.Value)
        MyP.Name = "Pt"
        MyHB.AppendHybridShape(MyP)
        Dim reference1 As INFITF.Reference = MyPart.CreateReferenceFromObject(MyP)
        Dim hybridShapeSphere1 As HybridShapeSphere = MyHSF.AddNewSphere(reference1, Nothing, 5.0, -45.0, 45.0, 0.0, 180.0)

        hybridShapeSphere1.Limitation = 1

        MyHB.AppendHybridShape(hybridShapeSphere1)

        MyPart.Update()

        TrackBar1.Enabled = True
        TrackBar2.Enabled = True
        TrackBar3.Enabled = True

        cmdCreerPoint.Enabled = False

    End Sub
