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

Example of Subclassing interface from within a class:
```vbnet
Option Explicit
Implements ISubclass
Private m_Parent As Form
Public Function Attach(Parent As Form) As Long
    Set m_Parent = Parent
    SetSubclass m_Grid.hWnd, Me
End Function
Private Sub Class_Terminate()
    If Not m_Parent Is Nothing Then RemoveSubclass m_Parent.hWnd, Me
    Set m_Parent = Nothing
End Sub
Private Function ISubclass_SubclassProc(ByVal hWnd As Long, _
                                        ByVal uMsg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long, _
                                        ByVal dwRefData As Long) As Long
    Const WM_MOUSEWHEEL As Long = &H20A
    Select Case uMsg
    Case WM_MOUSEWHEEL
        'FlexWheelScroll m_Grid, CSng(GET_WHEEL_DELTA_WPARAM(wParam)) / WHEEL_DELTA
        'Exit Function
    End Select
    ISubclass_SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function
```

