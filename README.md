# VSCode VBA Snippets
## Supported languages (file extensions)
* Visual Basic (".vb", ".brs", ".vbs", ".bas")
---
## Usage
After install this snippets add this inside your settings
```js
"editor.snippetSuggestions": "top",
```
---
## Snippets
### Basic code
- Dim
- If Else
- Loops
- Sub
- Function
- SelectCase
- MsgBox
---
### Dim [Dim declaration]
```vb
Dim arr()
Dim bol As Boolean
Dim lng As Long
Dim dbl As Double
Dim str As String
Dim obj As Object
Private
```
### If [If code block]
```vb
If condition Then

End If
```
### IfElse [If Else code block]
```vb
If condition Then

Else

End If
```
### IfElseIf [If ElseIf Else code block]
```vb
If condition1 Then

ElseIf condition2 Then

Else

End If
```
### With [With code block]
```vb
With

End With
```
### for [For Next Loop]
```vb
For index = lower To upper

Next index
```
### ForEach [For Each]
```vb
For Each variable In collection

Next variable
```
### DoLoopWhile [Do Loop While code block]
```vb
Do

Loop While condition
```
### DoWhile [Do While Loop code block]
```vb
Do While condition

Loop
```
### SubWithComments [Sub code block with comments]
```vb
Sub Main()
'
' @purpose:
'
'
'
' @algorithm:
'
'
'
' @test:
'
'
'



End Sub
```
### Sub [Sub code block]
```vb
Sub subName()
'
'
'
' @param
' @returns



End Sub
```
### Function [Function code block]
```vb
Function functionName()
'
'
'
' @param
' @returns



End Function
```
### SelectCase [Select Case code block]
```vb
Select Case test

  Case lists

    statements

  Case Else

    elseStatement

End Select
```
### MsgBox [Message box code block]
```vb
MsgBox("message", buttonType, "title")
```
