<div align="center">

## Find and Highlight Substring


</div>

### Description

' Given an editable textbox named Text1, this code prompts to find a word and

' searches throught the textbox and highlights the first occurance of the

' found word (if exists).
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bill Chelonis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bill-chelonis.md)
**Level**          |Unknown
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bill-chelonis-find-and-highlight-substring__1-1822/archive/master.zip)





### Source Code

```
Private Sub FindFunction_Click()
Rem Find/highlight first occurance of a word in a textbox named Text1
Dim a As String
Dim y As Integer
a = InputBox("Find text: ", "Find", "")
Call Text1.SetFocus
SendKeys ("^{HOME}")
y = 1
Do Until y = Len(Text1.text)
 Rem check if word was located
 If Mid(UCase$(Text1.text), y, Len(a)) = UCase$(a) Then
   Rem highlight the found word and exit sub
   For x = 1 To Len(a)
    SendKeys ("+{RIGHT}")
   Next x
   Exit Do
 End If
 Rem do nothing if carriage return encountered else highlight found word
 If Mid(Text1.text, y, 1) = Chr$(13) Then
 Else
 Rem move the cursor to the next element of text
 SendKeys ("{RIGHT}")
 End If
 y = y + 1
 If y > Len(Text1.text) Then Exit Do
Loop
End Sub
```

