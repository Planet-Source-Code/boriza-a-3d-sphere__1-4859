<div align="center">

## A  3D Sphere  \!\!\!


</div>

### Description

Draw a real 3D shpere using ONLY lines. This code can be easily modified to show other 3D objects. All you need to do is an array of coordinates. It does everything else for you (display, rotation, zoom etc) HAVE YOUR PERSONAL 3D ENGINE!!!
 
### More Info
 
'Try to change:

' number of angles and size of the polygon

'Also: Size of the sphere

Just create form (form1), module and PASTE the code!

This code rotates a polygon to create a sphere.

It does not always draw a full sphere, may be you can fix it...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Boriza](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/boriza.md)
**Level**          |Unknown
**User Rating**    |4.2 (67 globes from 16 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/boriza-a-3d-sphere__1-4859/archive/master.zip)

### API Declarations

```
Option Explicit
Public Const Pi = 3.1415926
'Number of angles of polygon
Public Const N_Angles = 5
'Diameter of the sphere
Public Const Sphere_Diam = 6000
Type Dot
  X As Double
  Y As Double
  Z As Double
End Type
'polygon in array
Public Object(1 To N_Angles + 1) As Dot
Public H_Globe, V_Globe
Public X, Y, Z
Public Me_to_Obj
Public Obj_to_Me
Public Polygon_R
Public Turn_Angle As Double
Function CRad(Deg)
'convert deg to rad
  CRad = Deg * Pi / 180
End Function
Function CDeg(Rad)
  CDeg = Rad * 180 / Pi
End Function
Public Sub GenPolygon()
'generate polygon
  Dim Angle
  Dim n As Double
    Angle = 360 / N_Angles
    For n = 1 To UBound(Object())
      Object(n).X = Sin(CRad(202.5 + (n - 1) * Angle)) * Polygon_R
      Object(n).Y = Cos(CRad(202.5 + (n - 1) * Angle)) * Polygon_R
      Object(n).Z = Sphere_Diam / 2
    Next n
    n = 1 - ((Polygon_R * 2) ^ 2) / (2 * ((Sphere_Diam / 2) ^ 2))
    n = n ^ 2
    n = Sqr(1 / n - 1)
    Turn_Angle = Atn(n)
End Sub
Public Sub Rotate(Obj() As Dot, HAngle, VAngle)
'this function rotates dots in array around the axes
  Dim X, Y, Z, c As Double
  Dim Ha, Va As Double
    Ha = HAngle + CRad(H_Globe)
    Va = VAngle + CRad(V_Globe)
    For c = 1 To UBound(Obj())
     If Ha <> 0 Then
      X = Obj(c).X
      Y = Obj(c).Y
      Z = Obj(c).Z
      Obj(c).Z = Z * Cos(Ha) - X * Sin(Ha)
      Obj(c).X = X * Cos(Ha) + Z * Sin(Ha)
     End If
     If Va <> 0 Then
      X = Obj(c).X
      Y = Obj(c).Y
      Z = Obj(c).Z
      Obj(c).Y = Y * Cos(Va) - Z * Sin(Va)
      Obj(c).Z = Z * Cos(Va) + Y * Sin(Va)
     End If
    Next c
End Sub
Public Sub DrawArray(Obj() As Dot)
'display array of dots on the screen
'Note: all dots are connected by lines
On Error Resume Next
  Dim n, d, dz
  Dim R, X1, Y1, X2, Y2
  d = Me_to_Obj
  dz = d + Obj_to_Me
  X2 = (Obj(1).X) * d / (Obj(1).Z + dz) + X
  Y2 = (Obj(1).Y) * d / (Obj(1).Z + dz) + Y
  For n = 0 To UBound(Obj()) - 1
    X1 = X2
    Y1 = Y2
    X2 = (Obj(n + 1).X) * d / (Obj(n + 1).Z + dz) + X
    Y2 = (Obj(n + 1).Y) * d / (Obj(n + 1).Z + dz) + Y
    'Swap next 2 lines to get full sphere:
     'Form1.Line (X1, Y1)-(X2, Y2)
     If Obj(n + 1).Z < 0 Then Form1.Line (X1, Y1)-(X2, Y2)
  Next n
End Sub
Public Sub Sphere()
'Displays polygons under different angles to construct a sphere
Form1.Cls
  Dim H, V, A, n
  A = Turn_Angle
  n = Val(2 * Pi / Turn_Angle)
  For H = 1 To n / 2
    For V = 1 To n
      DrawArray Object
      Rotate Object, A, 0
    Next V
    Rotate Object, 0, A
  Next H
End Sub
```


### Source Code

```
Option Explicit
Private Sub Form_Load()
Me.WindowState = 2
Me.BackColor = vbBlack
Me.ForeColor = vbWhite
Me.Caption = "3D Sphere - Your own 3D engine!               Programed by BORIZA"
Me.Show
'Position of sphere on the screen
Y = 4000
X = 6000
'Size of a polygon:
Polygon_R = 100
'Distance of the object from you
Me_to_Obj = 10000
Obj_to_Me = 1000
GenPolygon
DrawArray Object
Rotate Object, 0, -Pi / 2
Sphere
End Sub
```

