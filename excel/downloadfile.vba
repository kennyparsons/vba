Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Function DownloadFileFromWeb(strURL As String, strSavePath As String) As Long
    'strSavePath includes filename
    'returns 0 if download successful
    DownloadFileFromWeb = URLDownloadToFile(0, strURL, strSavePath, 0, 0)
End Function

Sub testdownload()
    Do Until DownloadFileFromWeb("https://url.tofile.com/path/to/file.png", "C:\path\to\folder\and\file.png") = 0    
    Loop
End Sub
