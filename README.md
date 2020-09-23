<div align="center">

## String Replace \*THE FASTEST\*


</div>

### Description

This code is a direct replacement for the VB6 Replace function. Obviously, VB 6 users will have very little (if any) use for the function. It is intended for use with older versions of VB that do not have the in-built Replace function. I noticed a similar piece of code released earlier this week, but it only appeared to handle character replacing and not strings. This version will do both and it's pretty tight and fast.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Leigh Bowers](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/leigh-bowers.md)
**Level**          |Unknown
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/leigh-bowers-string-replace-the-fastest__1-3635/archive/master.zip)





### Source Code

```
Public Function Replace(sExpression As String, sFind As String, sReplace As String) As String
' Title: Replace
' Version: 1.01
' Author: Leigh Bowers
' WWW:  http://www.esheep.freeserve.co.uk/compulsion
Dim lPos As Long
Dim iFindLength As Integer
' Ensure we have all required parameters
 If Len(sExpression) = 0 Or Len(sFind) = 0 Then
  Exit Function
 End If
' Determine the length of the sFind variable
 iFindLength = Len(sFind)
' Find the first instance of sFind
 lPos = InStr(sExpression, sFind)
' Process and find all subsequent instances
 Do Until lPos = 0
  sExpression = Left$(sExpression, lPos - 1) + sReplace + Mid$(sExpression, lPos + iFindLength)
  lPos = InStr(lPos, sExpression, sFind)
 Loop
' Return the result
 Replace = sExpression
End Function
```

