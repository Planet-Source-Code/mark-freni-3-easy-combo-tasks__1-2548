<div align="center">

## 3 Easy Combo Tasks


</div>

### Description

The code demonstrates 3 common combobox tasks:

1.) Filling a cbo with a recordset

2.) Setting the cbo Text to a recordset field using a numeric rst field

3.) Setting the cbo Text to a recordset field using a non-numeric rst field
 
### More Info
 
The name of a combobox control, and the recordset field names

The user needs to know how to open a recordset

Nothing, but they could be modified to boolean functions very easily


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Freni](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-freni.md)
**Level**          |Unknown
**User Rating**    |4.3 (60 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-freni-3-easy-combo-tasks__1-2548/archive/master.zip)





### Source Code

```
' Three simplified combobox Tasks:
'	1. Filling a cboBox with a Recordset
' 	2. Setting the cboText to a recordset field
'	  using an numeric recorset field.
'	3. Setting the cboText to a recordset field
'	  using a non-numeric recordset field.
'
'
Public Sub GetCBOList(cbo As ComboBox)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Filling a cboBox
' To make this more dynamic, pass the
' Sub the Desc as a string, and the ID
' As a long or integer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  On Error GoTo FUNCT_ERR
  Dim obj As New cClass
  Dim rst As New ADODB.Recordset
  ' I am using a class Method to get
  ' My Recordset. Getlist is a Class
  ' Function that returns a disconnected Recordset
  Set rst = obj.GetList
  ' Test the Recordset State to see
  ' it is open.
  If rst.State = 1 Then
	' Make sure I don't have an empty rst
    Do Until rst.EOF
      ' Always test for nulls
      If Not IsNull(rst!Desc) Then cbo.AddItem rst!Desc
      If Not IsNull(rst!UomID) Then cbo.ItemData(cbo.NewIndex) = rst!UomID
      ' Forget the movenext and you get an endless loop and
      ' an overflow error.
      rst.MoveNext
    Loop
    rst.Close
  End If
FUNCT_EXIT:
  Set obj = Nothing
  Set rst = Nothing
  Exit Sub
FUNCT_ERR:
  Err.Raise Err.Number, Err.Source, Err.Description
  Resume FUNCT_EXIT
End Sub
Public Sub SetCboText(cbo As ComboBox, val As Variant)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  PASS THE PROCEDURE A CBO NAME AND A RECORDSET FIELD
'  IF THE FIELD IS IN THE DROP-DOWN LIST IT WILL SET THE TEXT
'  VALUE FOR THAT CBO TO the listItem.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Dim i As Long
  ' LOOP THROUGH CBO Items
  For i = 0 To cbo.ListCount - 1
    If cbo.ItemData(i) = val Then
      cbo.ListIndex = i
      GoTo FUNCT_EXIT
    End If
  Next i
FUNCT_EXIT:
End Sub
Public Sub SetCboText_NonNumeric(cbo As ComboBox, val As Variant)
'  SUB USES cboBOXES THAT DO NOT HAVE A NUMERIC ITEMDATA VALUE
'  PASS THE PROCEDURE A CBO NAME AND A RECORDSET FIELD
'  IF THE FIELD IS IN THE DROP-DOWN LIST IT WILL SET THE TEXT
'  VALUE FOR THAT FIELD.
'  A good example of Non-Numeric ID is a StateCode: ie.
'  TX, MA, NY...
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Dim i As Long
  ' Loop through the CBO items, remember the cbo & lstBox
  ' are zero based lists
  For i = 0 To cbo.ListCount - 1
    If cbo.List(i) = val Then
      cbo.Text = cbo.List(i)
      ' DoEvents isn't really necessary
      DoEvents
      GoTo FUNCT_EXIT
    End If
  Next i
FUNCT_EXIT:
End Sub
```

