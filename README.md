# BoxUpload

This solution is only for BOX API testing purposes

onOpened xls instance by VS, press Alt+F11 add new module and paste code to the module:

'###############################
Sub VSTOcall()
    Dim addIn As COMAddIn
    Dim classObj As Object
    Set addIn = Application.COMAddIns("BoxUpload")
    Set classObj = addIn.Object
    classObj.uploadFile "0", [accessToken], "C:\Users\aciec\OneDrive\Documents\", "test.txt"
End Sub
'###############################

To use uploadFile feature in VBA use exposed method:
void uploadFile(string folderId, string accessToken, string filePath, string fileName);

folderId = 0 [main box.com user folder]
accessToken [login into box.com for developers portal -> create app -> generate DevToken from application dashboard -> paste as accessToken]
filePath [path to folder with file to be uploaded]
fileName [file + extension to be uploaded in given filePath folder] [example.txt]


