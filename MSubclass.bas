Attribute VB_Name = "MSubclass"
' Copyright © 2014 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' MSubclass.bas
'
' Subclassing Routines
'   - Dependancy: ISubclass.cls
'   - Windows handles unchaining the Message Subclassing
'   - Alias functions to comctl32.dll #410 #412 #413 for Windows 98/ME/2K
'
Option Explicit

Private Declare Function SetWindowSubclass Lib "comctl32" Alias "#410" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" Alias "#412" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
'Private Declare Function DefSubclassProc_ Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Function SubclassProc(ByVal hWnd As Long, _
                              ByVal uMsg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long, _
                              ByVal uIdSubclass As ISubclass, _
                              ByVal dwRefData As Long) As Long
    SubclassProc = uIdSubclass.SubclassProc(hWnd, uMsg, wParam, lParam, dwRefData)
End Function

Public Function SetSubclass(ByVal hWnd As Long, ByRef This As ISubclass, Optional dwRefData As Long) As Long
    SetSubclass = SetWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(This), dwRefData)
End Function

Public Function RemoveSubclass(ByVal hWnd As Long, ByRef This As ISubclass) As Long
    RemoveSubclass = RemoveWindowSubclass(hWnd, AddressOf SubclassProc, ObjPtr(This))
End Function

'Public Function DefSubclassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    DefSubclassProc = DefSubclassProc_(hWnd, uMsg, wParam, lParam)
'End Function


