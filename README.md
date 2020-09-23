<div align="center">

## Start or Restart an exe from within your Appl\.


</div>

### Description

starts an exe from within your application. But if the exe is already loaded, it becomes the focus! Normaly it starts with the poor shell-Command again and again...  //  IN GERMAN: Startet eine EXE aus Deiner VB-Applikation. Wenn Die EXE jedoch schon einmal gestartet wurde, wird sie lediglich fokusiert(!). Normalerweise würde sie wieder und wieder gestartet werden, wenn sie zuvor vom User nicht geschlossen wurde!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Joachim Berz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joachim-berz.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/joachim-berz-start-or-restart-an-exe-from-within-your-appl__1-6010/archive/master.zip)

### API Declarations

```
Option Explicit
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
```


### Source Code

```
Private Sub Command1_Click()
Dim Handle As Long
' the FindWindow-API needs the Caption-Name of the exe-File (e.g. Calculator for the Calc.exe!)
' Handle = FindWindow(vbNullString, "<CaptionNameOfExe>")
Handle = FindWindow(vbNullString, "Calculator") ' Is the exe already loaded?
' *! im deutschen Windows muss bei diesem Beispiel statt "Calculator" das Wort "Rechner" stehen!!!
If Handle = 0 Then ' _if the Handle becomes 0 then START the EXE-File
 Handle = Shell("CALC.EXE", 1)
 Else ' _if fires a Handle, the exe is there! Let´s focus it...
 ShowWindow Handle, 0 ' Hide the EXE (huh! Where is the exe???)
 ShowWindow Handle, 1 ' Show the EXE (now it becomes the Focus!!!)
End If
End Sub
```

