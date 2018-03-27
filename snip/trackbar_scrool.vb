Private Sub TrackBar_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll, TrackBar2.Scroll, TrackBar3.Scroll
        Dim CATIA As INFITF.Application
        CATIA = GetObject(, "CATIA.Application")


        LblX.Text = "X : " & CStr(TrackBar1.Value)
        lblY.Text = "Y : " & CStr(TrackBar2.Value)
        lblZ.Text = "Z : " & CStr(TrackBar3.Value)


        Dim MyPartDoc As MECMOD.PartDocument = CATIA.ActiveDocument
        Dim MyPart As MECMOD.Part = MyPartDoc.Part
        Dim MyHB As MECMOD.HybridBody = MyPart.HybridBodies.GetItem("Set géométrique.1")
        Dim MyHBS As MECMOD.HybridShapes = MyHB.HybridShapes

        Dim MyP As HybridShapeTypeLib.Point = MyHBS.GetItem("Pt")


        MyP.X.Value = CStr(TrackBar1.Value)
        MyP.Y.Value = CStr(TrackBar2.Value)
        MyP.Z.Value = CStr(TrackBar3.Value)

        MyPart.Update()

    End Sub
