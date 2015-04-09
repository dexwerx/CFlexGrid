CFlexGrid
=========

VB6 MSHFlexgrid Routines and Fixes (Scrollbar Thumbsize, and 64K+ Rows without Scrolltrack)

Includes generic subclassing framework.

CFlexGrid.cls - all Flexgrid Routines are here<br>
MSubclass.bas - Subclassing Framework, redirects Hook to an ISubclass Implemented class<br>
ISubclass.cls - Subclassing Interface Stub<br>
CMinMax.cls - Example Class Implementing Subclassing Framework
  
To use CFlexGrid or CMinMax from within a Form with a MSHFlexgrid on it:
```vbnet
Option Explicit
Private mm As New CMinMax
Private WithEvents fg As CFlexGrid
Private Sub Form_Load()
    Set fg = New CFlexGrid
    fg.Attach MSHFlexGrid1, Me
    mm.Attach Me, 320, 200
End Sub
```

