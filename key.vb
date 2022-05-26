Sub PasswordBreaker()
Dim i4 As Integer, i5 As Integer, i6 As Integer
On Error Resume Next
For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
For l = 65 To 66: For m = 65 To 66: For il = 65 To 66
For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
    Chr(l) & Chr(m) & Chr(il) & Chr(i2) & Chr(i3) & _
    Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
If ActiveSheet.ProtectContents = False Then
    MsgBox "One usable password is " & Chr(i) & Chr(j) & _
        Chr(k) & Chr(l) & Chr(m) & Chr(il) & Chr(i2) & _
        Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    Exit Sub
End If
Next: Next: Next: Next: Next: Next
Next: Next: Next: Next: Next: Next
End Sub
