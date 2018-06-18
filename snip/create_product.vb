Sub create_product()

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim productDocument1 As ProductDocument
Set productDocument1 = documents1.Add("Product")
'different than part document
