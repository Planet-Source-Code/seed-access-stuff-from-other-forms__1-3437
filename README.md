<div align="center">

## Access Stuff From Other Forms


</div>

### Description

This just shows how to access CommandButton's, Label's etc.. from another Command or label. So when you click Form1's Command, you can actually be clicking Form2's command. Very useful for shortcuts.
 
### More Info
 
1. make 2 forms

2. make a command button on each form

returns the action of the other command button


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SeeD](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/seed.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/seed-access-stuff-from-other-forms__1-3437/archive/master.zip)





### Source Code

```
'So you have your 2 forms? Good.
'Use the code below in the specified
'areas...
  Private Sub Command1_Click()
    Form2.Command1.Value = True
  End Sub
'That goes in form1's command button. Just add an action to the command button on form 2 to work it!
```

