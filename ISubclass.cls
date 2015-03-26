VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2014 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' ISubclass.cls
'
' Subclassing Interface
'
Option Explicit

Public Function SubclassProc(ByVal hWnd As Long, _
                             ByVal uMsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long, _
                             ByVal dwRefData As Long) As Long
End Function

'Implements ISubclass
'Private Function ISubclass_SubclassProc(ByVal hWnd As Long, _
'										ByVal uMsg As Long, _
'										ByVal wParam As Long, _
'										ByVal lParam As Long, _
'										ByVal dwRefData As Long) As Long
'	ISubclass_SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
'End Function



