Option Explicit

Dim objWMIService, colItems, objItem, ProductKey, WshShell, ProductName

Set WshShell = CreateObject("WScript.Shell")

On Error Resume Next
ProductName = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName")
On Error GoTo 0
If ProductName = "" Then ProductName = "Unknown Windows Edition"

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("SELECT OA3xOriginalProductKey FROM SoftwareLicensingService")

ProductKey = ""

For Each objItem in colItems
    If Not IsNull(objItem.OA3xOriginalProductKey) And Trim(objItem.OA3xOriginalProductKey) <> "" Then
        ProductKey = objItem.OA3xOriginalProductKey
    End If
Next

If ProductKey = "" Then
    MsgBox "Windows Edition: " & ProductName & vbCrLf & vbCrLf & "No OEM Product Key found in BIOS/UEFI." & vbCrLf & vbCrLf & "This usually means your PC uses a Digital License, or it was built from scratch without a pre-injected key.", vbExclamation, "System Information"
Else
    MsgBox "Windows Edition: " & ProductName & vbCrLf & vbCrLf & "BIOS/UEFI OEM Product Key: " & ProductKey, vbInformation, "System Information"
End If