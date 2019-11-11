 '************************************************************************************************
'
' Description: This macro loops through all nodes a product tree and informs the user as to
'              the node's "type" - part, product or component.  This is done as follows:
'
'               1. The tree is looped through by use of a recursive call to the GetNextNode sub
'
'               2. The node "type" is determine by attempting to set the node to either a *.CATPart
'                  or *.CATProduct filename
'
' Assumptions: 1. Active document must be a catia product
'
'   Developer: Brian Roach
'
'        Date: 5JAN2007
'
' CATIA Level: V5R14 SP8
'
'************************************************************************************************

Sub CATMain()

    GetNextNode CATIA.ActiveDocument.Product
  
End Sub

 

Sub GetNextNode(oCurrentProduct As Product)

    Dim oCurrentTreeNode As Product
    Dim i As Integer
       
    ' Loop through every tree node for the current product
    For i = 1 To oCurrentProduct.Products.Count
        Set oCurrentTreeNode = oCurrentProduct.Products.Item(i)

        ' Determine if the current node is a part, product or component
        If IsPart(oCurrentTreeNode) = True Then
            MsgBox oCurrentTreeNode.PartNumber & " is a part"
           
        ElseIf IsProduct(oCurrentTreeNode) = True Then
            MsgBox oCurrentTreeNode.PartNumber & " is a product"
           
        Else
            MsgBox oCurrentTreeNode.PartNumber & " is a component"
        End If
       
       
        ' if sub-nodes exist below the current tree node, call the sub recursively
        If oCurrentTreeNode.Products.Count > 0 Then
            GetNextNode oCurrentTreeNode
        End If
     
   Next

End Sub

Function IsPart(objCurrentProduct As Product) As Boolean

    Dim oTestPart As PartDocument
   
    Set oTestPart = Nothing
   
    On Error Resume Next
     
        Set oTestPart = CATIA.Documents.Item(objCurrentProduct.PartNumber & ".CATPart"  

        If Not oTestPart Is Nothing Then
            IsPart = True
        Else
            IsPart = False
        End If
        
End Function

Function IsProduct(objCurrentProduct As Product) As Boolean

    Dim oTestProduct As ProductDocument
   
    Set oTestProduct = Nothing
   
    On Error Resume Next
     
        Set oTestProduct = CATIA.Documents.Item(objCurrentProduct.PartNumber & ".CATProduct"  

        If Not oTestProduct Is Nothing Then
            IsProduct = True
        Else
            IsProduct = False
        End If
        
End Function
