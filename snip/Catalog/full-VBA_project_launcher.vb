Sub CATMain()
 
     Dim strPath As String
     Dim strModule As String
     Dim strProcedure As String
     Dim varArgs() As Variant    'Use empty array if procedure has no args
     'Dim varArgs(2) As Variant  'Or size the array if procedure has args
     Dim strMsg As String
 
     'Define where to find the VBA project and what to run inside it
     strPath = "C:\MyMacros\VBAProject1.catvba"
     strModule = "modMain"
     strProcedure = "CATMain"
 
     'If the procedure has arguments, define them in the array
     'As an example, a Sub with 3 input arguments in the VBA project
     'Sub CATMain(iInputfile, iOutputFile, iPart)
     'varArgs(0) = "C:\Data\InputData.txt"
     'varArgs(1) = "C:\Data\OutputData.txt"
     'Set varArgs(2) = CATIA.ActiveDocument.Part
 
     'Launch the VBA project
     On Error Resume Next
     Call CATIA.SystemService.ExecuteScript(strPath, _
                         2, strModule, strProcedure, varArgs)
     If Err.Number <> 0 Then     ‘Any number other than zero is an error
          'Add your own custom error message to the user
          strMsg = "An error occurred while trying to run the macro..."
          MsgBox strMsg, 16, "Error"
     End If
 
End Sub
