<div align="center">

## A simple way to add contents from one RichTextBox to another without losing formatting


</div>

### Description

This shows how the richtext of one richtextbox control can be inserted into another without losing any of its formatting.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dound](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dound.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dound-a-simple-way-to-add-contents-from-one-richtextbox-to-another-without-losing-formatti__1-48717/archive/master.zip)





### Source Code

```
Private Sub cmdAddRTBs_Click() 'Adds rtb2 to the end of rtb1
  'Set insert point (can be at ANY point in rtb1)
  rtb1.SelStart = Len(rtb1.Text)
  'Select rich text to add
  rtb2.SelStart = 0
  rtb2.SelLength = Len(rtb2.Text)
  'Add the selected rich text
  rtb1.SelRTF = rtb2.SelRTF
End Sub
```

