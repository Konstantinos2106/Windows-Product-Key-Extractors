Option Explicit

Dim WshShell, DigitalProductId, ProductKey, ProductName

Set WshShell = CreateObject("WScript.Shell")

On Error Resume Next
' Read Product ID and Product Name
DigitalProductId = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")
ProductName = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
On Error GoTo 0

If IsEmpty(DigitalProductId) Then
    MsgBox "Unable to read DigitalProductId from Registry", vbCritical, "Error"
Else
    ProductKey = ConvertToKey(DigitalProductId)
    
    ' Automatically copy the Product Key to the clipboard
    WshShell.Run "cmd.exe /c echo | set /p=""" & ProductKey & """ | clip", 0, True
    
    ' Display Windows Edition, Product Key, and a confirmation message
    MsgBox "Windows Edition: " & ProductName & vbCrLf & vbCrLf & "Product Key: " & ProductKey & vbCrLf & vbCrLf & "(The Product Key has been successfully copied to your clipboard!)", vbInformation, "System Information"
End If

Function ConvertToKey(Key)
    Const KeyOffset = 52
    Dim i, j, x, Cur, Chars, KeyOutput
    i = 28
    Chars = "BCDFGHJKMPQRTVWXY2346789"
    KeyOutput = ""
    Do
        Cur = 0
        x = 14
        Do
            Cur = Cur * 256
            Cur = Key(x + KeyOffset) + Cur
            Key(x + KeyOffset) = (Cur \ 24) And 255
            Cur = Cur Mod 24
            x = x - 1
        Loop While x >= 0
        i = i - 1
        KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
        
        If (((29 - i) Mod 6) = 0) And (i <> -1) Then
            i = i - 1
            KeyOutput = "-" & KeyOutput
        End If
    Loop While i >= 0
    ConvertToKey = KeyOutput
End Function