VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMinMax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2014 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CMinMax.cls
'
' Subclasses WM_GETMINMAXINFO, restricts minimum/maximum Window size
'   - Max is optional
'
Option Explicit

Implements ISubclass

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type POINT
   X    As Long
   Y    As Long
End Type

Private Type MINMAXINFO
   ptReserved       As POINT
   ptMaxSize        As POINT
   ptMaxPosition    As POINT
   ptMinTrackSize   As POINT
   ptMaxTrackSize   As POINT
End Type

Private m_Min   As POINT
Private m_Max   As POINT
Private m_hWnd  As Long
'

Private Function ISubclass_SubclassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    Const WM_GETMINMAXINFO As Long = &H24
    Select Case uMsg
    Case WM_GETMINMAXINFO
        Dim mmi As MINMAXINFO
        RtlMoveMemory mmi, ByVal lParam, LenB(mmi)
        If m_Max.X > 0 Then mmi.ptMaxTrackSize = m_Max
        If m_Min.X > 0 Then mmi.ptMinTrackSize = m_Min
        RtlMoveMemory ByVal lParam, mmi, LenB(mmi)
        ISubclass_SubclassProc = 1
        Exit Function
    End Select
    ISubclass_SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

Private Sub Class_Terminate()
    If m_hWnd Then RemoveSubclass m_hWnd, Me
End Sub

Public Function Attach(Parent As Form, MinX As Long, MinY As Long, Optional MaxX As Long, Optional MaxY As Long) As Long
    m_Min.X = MinX
    m_Min.Y = MinY
    m_Max.X = MaxX
    m_Max.Y = MaxY
    m_hWnd = Parent.hWnd
    Attach = SetSubclass(m_hWnd, Me)
End Function
