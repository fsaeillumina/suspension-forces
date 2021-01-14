Option Explicit
Option Base 1

Sub hide_help()

On Error Resume Next
Sheets("Formulas help").visible = False
Sheets("Calcs Help").visible = False
Sheets("Dynamic HP Help").visible = False
Sheets("Anti Ref").visible = False

End Sub

Sub show_help()

On Error Resume Next
Sheets("Formulas help").visible = True
Sheets("Calcs Help").visible = True
Sheets("Dynamic HP Help").visible = True
Sheets("Anti Ref").visible = True


End Sub



Function magnitude(ByRef i() As Double) As Variant
'returns magnitude of a [x,y,z] vector
'for use in VBA code
magnitude = Math.Sqr(i(1) * i(1) + i(2) * i(2) + i(3) * i(3))
End Function

Function magn(ByVal i As Variant) ' As Double
    'returns magnitude of a [x,y,z] vector
    'Suggestion: =magn(range1-range2) for distance between two xyz points
    'for use in spreadsheet. MUST use array equation (ctrl+shift+enter)
    magn = Math.Sqr(i(1) * i(1) + i(2) * i(2) + i(3) * i(3))
End Function

Function Unit_Vector(ByVal i As Variant)
    'for use in spreadsheet. MUST use array equation (ctrl+shift+enter)
    Dim returner As Variant
    ReDim returner(1 To 3)
    
    returner(1) = i(1) / magn(i)
    returner(2) = i(2) / magn(i)
    returner(3) = i(3) / magn(i)
    
    Unit_Vector = returner

End Function


Function AddVariant(ByRef a As Variant, ByRef b As Variant) As Variant
'adds two variables of variant type. For use in VBA code
AddVariant = Array(a(1) + b(1), a(2) + b(2), a(3) + b(3))
End Function

Function SubVariant(ByRef a As Variant, ByRef b As Variant) As Variant
'adds two variables of variant type. For use in VBA code
SubVariant = Array(a(1) - b(1), a(2) - b(2), a(3) - b(3))
End Function
Function cross_product(v1 As Variant, v2 As Variant) As Variant
    'use in spreadsheet or VBA code
    'Matrix equation, remember to use ctrl+shift+enter
    cross_product = Array(v1(2) * v2(3) - v1(3) * v2(2), _
    v1(3) * v2(1) - v1(1) * v2(3), _
    v1(1) * v2(2) - v1(2) * v2(1))
End Function


Function dot_product(v1 As Variant, v2 As Variant) As Variant
'use in spreadsheet or VBA code
'Matrix equation, remember to use ctrl+shift+enter
dot_product = v1(1) * v2(1) + v1(2) * v2(2) + v1(3) * v2(3)
    
End Function

Function plane_3p(p1 As Variant, p2 As Variant, p3 As Variant)
'returns [A,B,C,D] of the equation of a plane Ax+By+Cz+D=0
'input 3 points [x,y,z]
'Matrix equation, remember to use ctrl+shift+enter
Dim n_hat As Variant

n_hat = cross_product(SubVariant(p2, p1), SubVariant(p3, p1))
plane_3p = Array(n_hat(1), n_hat(2), n_hat(3), -dot_product(n_hat, p1))

End Function

Function plane_ln_intsct(p0 As Variant, p1 As Variant, plane As Variant) As Variant
'Returns [x,y,z] of the intersection between a
'line given by
    'p0=[x,y,z] and
    'p1=[x,y,z] and
'a plane in form
    'plane=[A,B,C,D]
    'where Ax+By+Cz+D=0
'Matrix equation, remember to use ctrl+shift+enter
Dim t, x0, y0, z0, x1, y1, z1, a, b, c, d As Double

x0 = p0(1): y0 = p0(2): z0 = p0(3)
x1 = p1(1): y1 = p1(2): z1 = p1(3)
a = plane(1): b = plane(2): c = plane(3): d = plane(4)

t = -(a * x0 + b * y0 + c * z0 + d) / (a * (x1 - x0) + b * (y1 - y0) + c * (z1 - z0))
plane_ln_intsct = Array(x0 + t * (x1 - x0), y0 + t * (y1 - y0), z0 + t * (z1 - z0))

End Function

Sub clearall()
ActiveSheet.DrawingObjects.Delete

End Sub

Sub CopyPaste()

Dim rPaste, rCopy, rTemp As Range


Err.Clear

Set rCopy = Selection

rCopy.Replace What:="=", Replacement:="#", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'Selection.Copy
On Error GoTo ExitSub
Set rTemp = Application.InputBox("Destination?", "Copy/Paste", , , , , , 8)

Set rPaste = Range(rTemp.Cells(1, 1), rTemp.Cells(1, 1).offset(rCopy.Rows.Count - 1, rCopy.Columns.Count - 1))

'rPaste.Activate
'rPaste.Activate
rCopy.Copy
rPaste.PasteSpecial

'ActiveSheet.Paste
'rTemp.PasteSpecial (xlPasteAll)
Application.CutCopyMode = False

rPaste.Replace What:="#", Replacement:="=", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

ExitSub:
On Error GoTo 0
rCopy.Replace What:="#", Replacement:="=", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
Err.Clear


End Sub


Sub Switch_Points()

Dim Current_Points As String
Dim New_Points, Old_Points As Range

Application.ScreenUpdating = False

'On Error GoTo ErrorLabel
Current_Points = Range("Points_Flag").Value
Set New_Points = Worksheets("Points").Range(Range("Switch_SwitchTo").Value)
Set Old_Points = Worksheets("Points").Range(Range("Switch_Current").Value)

'==================================
'Save Old Values
With Old_Points
    Range(.offset(1, 2), .offset(24, 4)).Value = Range("Calc_Points").Value
    Range(.offset(1, 5), .offset(24, 5)).Value = Range("Calc_Tri_Soln").Value
    Range(.offset(1, 8), .offset(22, 12)).Value = Range("Calc_Loadcases").Value
End With

'Insert New Values in "Points"
With New_Points
    Range("Calc_Points").Value = Range(.offset(1, 2), .offset(24, 4)).Value
    Range("Calc_Tri_Soln").Value = Range(.offset(1, 5), .offset(24, 5)).Value '.offset(1, 5, 24, 1).Value
    Range("Calc_Loadcases").Value = Range(.offset(1, 8), .offset(22, 12)).Value 'New_Points.offset(1, 8, 22, 7).Value ' NOTE: CHANGE 7 to 8
    Range("Points_Flag").Value = .Value
End With

'Updates Data Tables if automatic calculation is turned off. This is necessary to get new values for MR, anti, etc
Application.Calculate

'========================
'Update Values in "Points Index"
With Sheets("Points Index").Range("A1", "A999").Find(Range("Calcs_SwitchTo").Value)

'If any of the values in the cells are errors (like #VALUE), it causes a runtime error in the code
'On Error Resume Next
    .offset(0, 3).Value = Range("Calcs_MR").Value
    .offset(0, 4).Value = Range("Calcs_MRLinearity").Value

If .offset(0, 2).Value = "Front" Then
    .offset(0, 5).Value = Range("Calcs_AntiDive").Value
    
ElseIf .offset(0, 2).Value = "Rear" Then
    .offset(0, 6).Value = Range("Calcs_AntiLift").Value
    .offset(0, 7).Value = Range("Calcs_AntiSquat").Value
    .offset(0, 8).Value = Range("Calcs_RCHeight").Value
Else
    MsgBox "Front of Rear not specified"

End If
'====================================

On Error GoTo 0
ActiveWorkbook.Worksheets("Points Index").Range("A1").Value = "Test from xl"

End With


Application.ScreenUpdating = True

Exit Sub
ErrorLabel:
Application.ScreenUpdating = True
MsgBox "Error switching points"

End Sub

Sub test()

Dim xlapp As Object
Dim xlWb As Object
Set xlapp = GetObject(, "Excel.Application")
Set xlWb = xlapp.ActiveWorkbook

Dim i As Integer
Dim rTemp As Range
Dim CopyPoints As String
CopyPoints = "Rear"

Debug.Print Worksheets("Points").Range("A1:A999").Find(CopyPoints, LookIn:=xlValues)


With xlWb.Worksheets("Points")
'    Debug.Print .Range("A1", "A999").Find(CopyPoints).Value
    With .Range(.Range("A1", "A999").Find(CopyPoints, LookIn:=xlValues).offset(1, 6), _
                .Range("A1", "A999").Find(CopyPoints, LookIn:=xlValues).offset(24, 6))

        For i = 1 To 24
        If (i <> 7) And (i <> 18) Then
            Debug.Print i & " " & .Cells(i, 1).Value
            'Me.Controls("Textbox" & i).Text = .range(i, 0).Value
        End If
        Next i
    End With
End With
End Sub
