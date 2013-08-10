''' VBS to convert all PDF files in a folder to XML for easy parsing.
''' Michael Murphy 2013

if WScript.Arguments.Count < 1 Then
   WScript.Echo "Error! Please specify the source folder"
   Wscript.Quit
End If

Dim oAcroApp      
Dim oAcroAVDoc    
Dim oAcroPDDoc    
Dim oJSO          
Dim filePath
Dim oAcroForm
Dim bo

folderPath	= Wscript.Arguments.Item(0)
Set oFSO	= CreateObject("Scripting.FileSystemObject")
Set oFolder	= oFSO.GetFolder(folderPath)
Set cFiles	= oFolder.Files
Set oAcroApp	= CreateObject("AcroExch.App")
Set oAcroAVDoc	= CreateObject("AcroExch.AVDoc")

For Each oFile in cFiles
   If UCase(oFSO.GetExtensionName(oFile.Name)) = "PDF" Then
      inFilePath = folderPath & oFile.Name
      outFilePath = folderPath & oFile.Name & ".xml"
      bo		= oAcroAVDoc.Open(inFilePath, "XML Export")
      bo                = oAcroApp.Hide
      Set oAcroPDDoc	= oAcroAVDoc.GetPDDoc
      Set oJSO		= oAcroPDDoc.GetJSObject
      oJSO.disclosed	= TRUE
      bo = oJSO.SaveAs(outFilePath,"com.adobe.acrobat.xml-1-00")
      bo = oAcroAVDoc.Close(True)
   End If
Next

oAcroApp.Exit
Set oAcroPDDoc = Nothing
Set oAcroAVDoc = Nothing
Set oAcroApp   = Nothing