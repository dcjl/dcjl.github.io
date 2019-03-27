---
Title: "Get the nth word from a text string"
layout: single
classes: wide
categories:
  - vba
tags:
  - text-strings
  - function
date: 2019-03-20 21:27:00
sidebar:
  title: "Popular Links"
  nav: sidebar-sample

---

# Get the nth word from a text string
If the word position is a negative number or exceeds the amount of words in the string then the user is notified.

```vb
'==========================================================================================================
' ## Function: Get the nth word from a text string
'    If the word position is a negative number or exceeds the amount
'    of words in the string then the user is notified
'==========================================================================================================
Function ExtractWord(Source As String, Position As Long)
    Dim arr() As String
    arr = VBA.Split(Source, "/")
    xCount = UBound(arr)
    If xCount < 1 Or (Position - 1) > xCount Or Position < 0 Then
        ExtractWord = "You have either entered a number that is more than the total words" & vbLf & _
            "or" & vbLf & _
            "You have entered a negative number"
    Else
        ExtractWord = arr(Position - 1)
    End If
End Function
```
