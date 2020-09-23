<div align="center">

## Full example of Drag and Drop within a application


</div>

### Description

Suppose you have a listbox with some elements and want to drag&drop a selected one into a textbox. http://137.56.41.168:2080/VisualBasicSource/vbdraganddrop.txt
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Found on the World Wide Web](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/found-on-the-world-wide-web.md)
**Level**          |Unknown
**User Rating**    |4.2 (176 globes from 42 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/found-on-the-world-wide-web-full-example-of-drag-and-drop-within-a-application__1-650/archive/master.zip)





### Source Code

```
Make a form with a textbox (text1) and a listbox (list1). Fill the listbox with some items...
Make a label (label1). Set it invisible = False
Put the next code at the appropiate places:
Sub List1_MouseDown (Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim DY
  DY = TextHeight("A")
  Label1.Move list1.Left, list1.Top + Y - DY / 2, list1.Width, DY
  Label1.Drag
End Sub
Sub List1_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
  If State = 0 Then Source.MousePointer = 12
  If State = 1 Then Source.MousePointer = 0
End Sub
Sub Form_DragOver (Source As Control, X As Single, Y As Single, State As Integer)
  If State = 0 Then Source.MousePointer = 12
  If State = 1 Then Source.MousePointer = 0
End Sub
Sub Text1_DragDrop (Index As Integer, Source As Control, X As Single, Y As Single)
  text1.text = list1
End Sub
```

