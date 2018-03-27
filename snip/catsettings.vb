        Dim CATIA As INFITF.Application
        CATIA = GetObject(, "CATIA.Application") 'Recupère l'appli Catia

        Dim SettingsControllers1 As INFITF.SettingControllers = CATIA.SettingControllers 'Définit l'objet Options

        Dim unitsSheetSettingAtt1 As INFITF.SettingController = SettingsControllers1.Item("CATLieUnitsSheetSettingCtrl") 'Définit l'objet onglet unités

        ' Extraction de l'unité du type

        Dim bSTR1 As String
        bSTR1 = "LENGTH" 'Correspond à la grandeur recherchée
        Dim bSTR2 As String
        bSTR2 = ""  'Correspond à l'unité recherchée

        'Correspond aux réglages d'affichage (nombre de chiffres)
        Dim double1 As Double
        Dim double2 As Double

        unitsSheetSettingAtt1.GetMagnitudeValues(bSTR1, bSTR2, double1, double2) 'Récupère les valeurs de la ligne d'option correspondant aux critères

        'print de l'unité actuelle de Catia dans le texte LabelUnit.

        LabelUnite.Text = bSTR2
