Option Base 1
Option Explicit

Type line_prop

    name As String
    x_start As Double
    y_start As Double
    x_end As Double
    y_end As Double
    colourr As Integer
    colourg As Integer
    colourb As Integer
    weight As Integer
    
'    .name =
'    .x_start =
'    .y_start =
'    .x_end =
'    .y_end =
'    .colourr = 255
'    .colourg = 255
'    .colourb = 255
'    .weight = 6
    
End Type

Dim shape_names(160) As Variant
Dim shape_counter As Integer
Dim colour_counter As Integer
Dim colour_array As Variant



Sub draw_line(liner As line_prop, Optional ArrowType As String = "arrow")
'On Error GoTo error1

With ActiveSheet.Shapes.AddConnector(msoConnectorStraight, liner.x_start, liner.y_start, liner.x_end, liner.y_end)
    If ArrowType = "arrow" Then
        .line.EndArrowheadStyle = msoArrowheadOpen
    Else
        .line.EndArrowheadStyle = msoArrowheadNone
        '.line.StartArrowheadStyle = msoArrowheadNone
    End If
    .ShapeStyle = msoLineStylePreset8
    .name = liner.name
    .Shadow.visible = msoFalse
    .line.weight = liner.weight
    
    If ActiveSheet.CheckBoxes("rng1_checkbox_colour").Value = xlOn Then
    .line.ForeColor.RGB = RGB(colour_array(colour_counter, 1), colour_array(colour_counter, 2), colour_array(colour_counter, 3))                                                'made by: Bennet Heidenreich 2014
    
    Else
    .line.ForeColor.RGB = RGB(liner.colourr, liner.colourg, liner.colourb)
    End If
End With
'error1:
End Sub

Sub test_draw()
Dim i As line_prop
With i

'    .name =
'    .x_start =
'    .y_start =
'    .x_end =
'    .y_end =
'    .colourr = 255
'    .colourg = 255
'    .colourb = 255
'    .weight = 6

    .name = "Test Arrow"
    .x_start = 50
    .y_start = 75
    .x_end = 400
    .y_end = 1800
    .colourr = 255
    .colourg = 0
    .colourb = 0
    .weight = 5
End With

Call draw_line(i)

End Sub

Sub dimensions(ByVal newname, ByRef j, ByRef k)
 'based on what dimensions are chosen, function assigns j and k

If Range("horz" & CStr(newname)) = "Z" Then
    j = 3
ElseIf Range("horz" & CStr(newname)) = "Y" Then
    j = 2
Else: j = 1
End If

If Range("vert" & CStr(newname)) = "Z" Then
    k = 3
ElseIf Range("vert" & CStr(newname)) = "Y" Then
    k = 2
Else: k = 1
End If

End Sub

Sub transformer(ByRef linest As line_prop, _
ByVal rotate As Integer, _
ByVal x_mirror As String, _
ByVal y_mirror As String)

Dim temp As Long

With linest
'------ Rotate
If rotate = 1 Then
ElseIf rotate = 2 Then
    temp = .x_start
    .x_start = -.y_start
    .y_start = temp
    temp = .x_end
    .x_end = -.y_end
    .y_end = temp
    
ElseIf rotate = 3 Then
    .x_start = -.x_start
    .y_start = -.y_start
    .x_end = -.x_end
    .y_end = -.y_end
    
ElseIf rotate = 4 Then
    temp = .x_start
    .x_start = .y_start
    .y_start = -temp
    temp = .x_end
    .x_end = .y_end
    .y_end = -temp

End If
'----end rotate
'----Flip
If x_mirror = "Yes" Then
.x_start = -.x_start
.x_end = -.x_end
End If
If y_mirror = "Yes" Then
.y_start = -.y_start
.y_end = -.y_end
End If
'----end flip
End With
End Sub

Sub check_center(ByRef newline As line_prop, _
ByRef i, _
ByRef x_offset, _
ByRef y_offset, _
ByVal zerox As Double, _
ByVal zeroy As Double)

If i = 1 And ActiveSheet.CheckBoxes("rng1_checkbox_center").Value = xlOn Then ' recenter offsets
x_offset = zerox - newline.x_start
Range("xoff1").Value = x_offset
y_offset = zeroy - newline.y_start
Range("yoff1").Value = y_offset
End If

End Sub

Sub offset(ByRef linest As line_prop, _
ByVal x_offset As Long, _
ByVal y_offset As Long)

With linest
.x_start = .x_start + x_offset
.y_start = .y_start + y_offset
.x_end = .x_end + x_offset
.y_end = .y_end + y_offset
End With

End Sub

Sub draw_forces(p As Integer, n As Integer)
'n is hardpoint set number
'p is loadcase set number

'On Error GoTo error2
'Worksheets("Draw").Activate

Dim i, j, k, m, newname, rotate As Integer 'j and k set dimensions. m is a counter for applied forces
Dim newline As line_prop
Dim unitVector_All, forces, applied, names, hp_p2, hp_applied As Variant
Dim scales, x_offset, y_offset, temp, force_scales, zerox, zeroy As Double
Dim x_mirror, y_mirror As String
Dim resultant1 As Range

newname = 1
x_mirror = Range("xmirror" & CStr(newname)).Value
y_mirror = Range("ymirror" & CStr(newname)).Value
force_scales = 500 / Range("force_scales" & CStr(newname)).Value
rotate = Range("rotate" & CStr(newname)).Value
x_offset = Range("xoff" & CStr(newname)).Value
y_offset = Range("yoff" & CStr(newname)).Value
scales = 50 / Range("scale" & CStr(newname)).Value
zerox = 575
zeroy = 125

Set applied = Range("applied" & CStr(n))
Set resultant1 = Range("resultant" & CStr(n))
Set hp_p2 = Range("HP_P2" & CStr(n))
Set hp_applied = Range("HP_applied" & CStr(n))
Set names = Range("names" & CStr(n))
Set forces = Range("forces" & CStr(n))
Set unitVector_All = Range("unitVector_All" & CStr(n))
With newline
'    .name =
'    .x_start =
'    .y_start =
'    .x_end =
'    .y_end =
'     .colourr = 0
'     .colourg = 0
'     .colourb = 0
      .weight = 4
End With

Call dimensions(newname, j, k) 'based on what dimensions are chosen, function assigns j and k

'---------main members
For i = 1 To 6 'i cycles through each force

    With newline

     .name = names(1, i).Value & "-" & CStr(n) & ": " & CInt(forces(p, i)) & "N"
If forces(p, i) >= 0 Then 'if force is POSITIVE, member is in COMPRESSION
    .x_start = -(unitVector_All(i, j) * forces(p, i) / force_scales) + hp_p2(i, j) / scales   'unitVector_All(i, j) *
    .y_start = -(unitVector_All(i, k) * forces(p, i) / force_scales) + hp_p2(i, k) / scales  'unitVector_All(i, k) *
    .x_end = hp_p2(i, j) / scales
    .y_end = hp_p2(i, k) / scales
    .colourr = 255
    .colourg = 0
    .colourb = 0

Else 'if force is NEGATIVE, member is in TENSION
    .x_start = hp_p2(i, j) / scales
    .y_start = hp_p2(i, k) / scales
    .x_end = (unitVector_All(i, j) * forces(p, i) / force_scales) + hp_p2(i, j) / scales   'unitVector_All(i, j) *
    .y_end = (unitVector_All(i, k) * forces(p, i) / force_scales) + hp_p2(i, k) / scales   'unitVector_All(i, k) *
    .colourr = 50
    .colourg = 200
    .colourb = 50
End If
End With

shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion
Call transformer(newline, rotate, x_mirror, y_mirror)
Call check_center(newline, i, x_offset, y_offset, zerox, zeroy)
Call offset(newline, x_offset, y_offset) 'do offsets

Call draw_line(newline) ' DRAW FINALLY
Next

'-------Resultant
Debug.Print ActiveSheet.CheckBoxes("rng1_checkbox_resultant").Value
If ActiveSheet.CheckBoxes("rng1_checkbox_resultant").Value = xlOn Then

With newline
    .name = "Result-" & CStr(n)
    .x_start = hp_p2(1, j) / scales
    .y_start = hp_p2(1, k) / scales
    .x_end = (resultant1(p, j)) / force_scales + hp_p2(1, j) / scales
    .y_end = (resultant1(p, k)) / force_scales + hp_p2(1, k) / scales
    .colourr = 0
    .colourg = 50
    .colourb = 250
End With

shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion BH2014
Call transformer(newline, rotate, x_mirror, y_mirror)
Call offset(newline, x_offset, y_offset)
Call draw_line(newline)

'add resultant2
With newline
    .name = "rng1_Result-" & CStr(n)
    .x_start = hp_p2(4, j) / scales 'picks outboard hardpoint of member 4
    .y_start = hp_p2(4, k) / scales
    .x_end = (resultant1(p, j + 3)) / force_scales + hp_p2(4, j) / scales '"+3" shifts from 1st resultant to 2nd
    .y_end = (resultant1(p, k + 3)) / force_scales + hp_p2(4, k) / scales
    .colourr = 0
    .colourg = 50
    .colourb = 250
End With


shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion
Call transformer(newline, rotate, x_mirror, y_mirror)
Call offset(newline, x_offset, y_offset)
Call draw_line(newline)
End If

'---------draw applied forces
If ActiveSheet.CheckBoxes("rng1_checkbox_applied").Value = xlOn Then
With newline
    For m = 1 To 3
        If (m = j) Then '----- draw horizontal
        .name = "Applied_" & Range("horz" & CStr(newname)).Value & "-" & CStr(n)
        .x_start = -applied(p, j) / force_scales + hp_applied(m, j) / scales
        .y_start = hp_applied(m, k) / scales
        .x_end = hp_applied(m, j) / scales
        .y_end = hp_applied(m, k) / scales
        .colourr = 128
        .colourg = 0
        .colourb = 128
        
        
        shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion
        Call transformer(newline, rotate, x_mirror, y_mirror)
        Call offset(newline, x_offset, y_offset)
        Call draw_line(newline)
        '-----
        
        ElseIf (m = k) Then ' -----draw vertical
        .name = "Applied_" & Range("vert" & CStr(newname)).Value & "-" & CStr(n)
        .x_start = hp_applied(m, j) / scales
        .y_start = -applied(p, k) / force_scales + hp_applied(m, k) / scales
        .x_end = hp_applied(m, j) / scales
        .y_end = hp_applied(m, k) / scales
        .colourr = 128
        .colourg = 0
        .colourb = 128
        
        
        shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion
        Call transformer(newline, rotate, x_mirror, y_mirror)
        Call offset(newline, x_offset, y_offset)
        Call draw_line(newline)
        
        End If
        
    Next
End With
End If
'error2:
End Sub

Sub draw_members(loadcase, hardpoint)
Dim newline As line_prop
Dim HP2, HP1, HP_bell As Variant
Dim scales, x_offset, y_offset, temp, zerox, zeroy As Double
Dim rotate As Integer
Dim x_mirror, y_mirror As String
Dim i, j, k, n, newname As Integer
Dim index As Integer

newname = 1

x_mirror = Range("xmirror" & CStr(newname)).Value
y_mirror = Range("ymirror" & CStr(newname)).Value
rotate = Range("rotate" & CStr(newname)).Value
x_offset = Range("xoff" & CStr(newname)).Value
y_offset = Range("yoff" & CStr(newname)).Value
scales = 50 / Range("scale" & CStr(newname)).Value
zerox = 575
zeroy = 125

Set HP2 = Range("HP_P2" & CStr(hardpoint))
Set HP1 = Range("HP_P1" & CStr(hardpoint))
Set HP_bell = Range("HP_bell" & CStr(hardpoint))

'Call check_center(i, x_offset, y_offset, zerox, zeroy) 'Move!

With newline
    .name = "Member"
    .colourb = 0
    .colourg = 0
    .colourr = 0
    .weight = 8
End With

Call dimensions(newname, j, k)

'Draw main 6 members-------------------------
For i = 1 To 6
With newline
    .x_end = HP2(i, j) / scales
    .y_end = HP2(i, k) / scales
    .x_start = HP1(i, j) / scales
    .y_start = HP1(i, k) / scales
    
End With
Call transformer(newline, rotate, x_mirror, y_mirror)
Call offset(newline, x_offset, y_offset)
Call draw_line(newline, "no arrowheads")
shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion
Next

'draw bellcrank---------------------------

For i = 1 To 6

Select Case i
    Case 1: n = 2
    Case 2: n = 4
    Case 3: GoTo nextloop
    Case 4: n = 1
    Case 5: GoTo nextloop
    Case 6: n = 2 'damper
End Select

With newline
    .x_end = HP_bell(i, j) / scales
    .y_end = HP_bell(i, k) / scales
    .x_start = HP_bell(n, j) / scales
    .y_start = HP_bell(n, k) / scales
End With
Call transformer(newline, rotate, x_mirror, y_mirror)
Call offset(newline, x_offset, y_offset)
Call draw_line(newline, "no arrowheads")
shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion

nextloop:
Next

'Draw damper----------------------
'With newline
'    .x_end = HP_bell(2, j) / scales
'    .y_end = HP_bell(2, k) / scales
'    .x_start = HP_bell(6, j) / scales
'    .y_start = HP_bell(6, k) / scales
'End With
'Call transformer(newline, rotate, x_mirror, y_mirror)
'Call offset(newline, x_offset, y_offset)
'Call draw_line(newline, "no arrowheads")
'shape_counter = shape_counter + 1: shape_names(shape_counter) = newline.name 'Tracks names of shapes for easy deletion

End Sub

Sub clear_shapes()
If shape_counter < 1 Then shape_counter = 1
Dim i As Integer

For i = 1 To 6
On Error Resume Next
    ActiveSheet.Shapes.Range(shape_names(shape_counter)).Delete
    If shape_counter > 1 Then shape_counter = shape_counter - 1
Next

End Sub

Sub clear_all()

'Try with current shape_names
For shape_counter = 1 To UBound(shape_names)
On Error Resume Next
    ActiveSheet.Shapes.Range(shape_names(shape_counter)).Delete
Next

shape_counter = 1
End Sub




Sub draw_sets()
Dim i As Integer 'counter for draw_who rows
Dim j As Integer 'counter for draw_who columns
Dim draw_who As Variant
Dim temp_loadcase As Integer

Set colour_array = Range("colour_array")
Set draw_who = Range("draw_who")
temp_loadcase = Range("Loadcase").Value
If (colour_counter < 1) Or (colour_counter > 15) Then colour_counter = 1

'Worksheets("6x6 Matrix").Calculate

For i = 1 To draw_who.Rows.Count
    For j = 1 To draw_who.Columns.Count
        If draw_who(i, j) = 1 Then
        Range("Loadcase").Value = i
        Call draw_forces(i, j)
        If ActiveSheet.CheckBoxes("rng1_checkbox_colour").Value = xlOn Then colour_counter = colour_counter + 1
        End If
    Next
    j = 1
Next
Range("Loadcase").Value = temp_loadcase
End Sub

Sub draw_members_all()

Dim i As Integer 'counter for draw_who rows
Dim j As Integer 'counter for draw_who columns
Dim draw_who As Variant
Dim temp_loadcase As Integer

Set colour_array = Range("colour_array")
Set draw_who = Range("draw_who")

temp_loadcase = Range("Loadcase").Value
'Worksheets("6x6 Matrix").Calculate

If (colour_counter < 1) Or (colour_counter > 15) Then colour_counter = 1
For i = 1 To draw_who.Rows.Count
    For j = 1 To draw_who.Columns.Count
        If draw_who(i, j) = 1 Then
        Range("Loadcase").Value = i
        Call draw_members(i, j)
        If ActiveSheet.CheckBoxes("rng1_checkbox_colour").Value = xlOn Then colour_counter = colour_counter + 1
        End If
    Next
    j = 1 'restart counter
Next

Range("Loadcase").Value = temp_loadcase

End Sub

Sub clearall()
ActiveSheet.DrawingObjects.Delete

End Sub

Sub mySelectAll()
    MsgBox ("Deselect all of the buttons and checkboxes before hitting delete!")
    ActiveSheet.Shapes.SelectAll
End Sub

