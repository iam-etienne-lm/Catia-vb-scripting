Dim CATIA As INFITF.Application
        CATIA = GetObject(, "CATIA.Application") 'Recupère l'appli Catia

        Dim ciClone As CultureInfo = CType(CultureInfo.InvariantCulture.Clone(), CultureInfo)
        ciClone.NumberFormat.NumberDecimalSeparator = "."

        Dim partDocument1 As MECMOD.PartDocument = CATIA.ActiveDocument  ‘ le document actif

        Dim Part1 As MECMOD.Part = partDocument1.Part   ‘Le doc actif est de type part

        Dim parameters1 As Parameters = Part1.Parameters ‘ le Groupe de paramètres de la part

        Dim DimParam1 As Parameter = parameters1.CreateDimension("", "Length", 0) ‘on crée un objet paramètre de type dimension longueur sans nom et de valeur 0

        DimParam1.Rename(TextBoxNomParam.Text) ‘on renomme le paramètre avec le nom dans la TextBoxNomParam

        DimParam1.Value = Convert.ToDouble(TextBoxValeur.Text, ciClone) ‘On donne une valeur au paramètre par conversion du texte présent dans la TextBoxValeur

        Part1.Update()
