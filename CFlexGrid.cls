VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Copyright © 2014 Dexter Freivald. All Rights Reserved. DEXWERX.COM
'
' CFlexGrid.cls
'
' Compilation of MSHFlexgrid code
'   - Fixes Scrollbar thumbsize on a timer to lessen flicker
'   - Latency property to set delay of Scrollbar Page/Thumbsize fix
'   - PageCols and PageRows Read-Only properties
'   - Mousewheel support
'   - Row Numbering
'   - Alternate Row Highlighting via HighlightColor property
'   - Auto Column resizing to width of contents
'   - User ColumnResized Event (some false positives)
'   - Handles > 64K Rows (Not recommended...)
'   - Shuts off scrolltracking when Scrolling past 64K Rows
'   - Dependancy: ISubclass.cls, MSubclass.bas
'
Option Explicit

Public Event ColumnResized(Col As Long)

Private Const INITIALSCROLLFIXDELAY     As Long = 100       'milliseconds
Private Const FLEX_GRIDLINEWIDTH        As Long = 1         'pixels
Private Const FLEX_CELLPADDING          As Long = 2         'pixels
Private Const ERR_SUBSCRIPTOUTOFRANGE   As Long = 9
Private Const UINT16_MAXVALUE           As Long = &HFFFF&   '64K Standard Scrollbar / MSHFlexGrid limitation

Private Type SCROLLINFO
    cbSize      As Long
    fMask       As Long
    nMin        As Long
    nMax        As Long
    nPage       As Long
    nPos        As Long
    nTrackPos   As Long
End Type

Private Type THighlightColor
    Color       As Long
    Modulus     As Long
End Type

Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Long, ByRef lpsi As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal fnBar As Long, ByRef lpsi As SCROLLINFO) As Long
Private Declare Function SystemParametersInfoW Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private m_Grid                  As MSHFlexGrid
Private m_Parent                As Form
Private m_Redraw                As Boolean
Private WithEvents m_TimerFix   As Timer
Attribute m_TimerFix.VB_VarHelpID = -1
Private m_SetScrollOnTimer      As Boolean
Private m_HighlightColor()      As THighlightColor

Implements ISubclass

Public Function Attach(Grid As MSHFlexGrid, Parent As Form) As Long
    Set m_Grid = Grid
    Set m_Parent = Parent
    Set m_TimerFix = m_Parent.Controls.Add("VB.Timer", "tmrSetScroll" & m_Parent.Controls.Count)
    m_TimerFix.Interval = INITIALSCROLLFIXDELAY
    m_SetScrollOnTimer = True
    SetSubclass m_Parent.hWnd, Me
    SetSubclass m_Grid.hWnd, Me
    FixScrollbarPageSize
End Function

Private Sub Class_Initialize()
    ReDim m_HighlightColor(0) As THighlightColor
    m_Redraw = True
End Sub

Private Sub Class_Terminate()
    Set m_TimerFix = Nothing
    If Not m_Grid Is Nothing Then RemoveSubclass m_Grid.hWnd, Me
    Set m_Grid = Nothing
    If Not m_Parent Is Nothing Then RemoveSubclass m_Parent.hWnd, Me
    Set m_Parent = Nothing
End Sub

Private Function ISubclass_SubclassProc(ByVal hWnd As Long, _
                                        ByVal uMsg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long, _
                                        ByVal dwRefData As Long) As Long
    Const WM_MOUSEWHEEL     As Long = &H20A
    Const WM_LBUTTONDOWN    As Long = &H201
    Const WM_LBUTTONUP      As Long = &H202
    Const WM_ERASEBKGND     As Long = &H14
    Const WM_SIZE           As Long = &H5
    Const WM_HSCROLL        As Long = &H114
    Const WM_VSCROLL        As Long = &H115
    Const SB_LINEDOWN       As Long = 1
    Const SB_PAGEDOWN       As Long = 3
    Const SB_LINERIGHT      As Long = 1
    Const SB_PAGERIGHT      As Long = 3
    Const SB_ENDSCROLL      As Long = 8
    Const SB_VERT           As Long = 1
    Const SB_THUMBPOSITION  As Long = 4
    Const SB_THUMBTRACK     As Long = 5
    Const SIF_TRACKPOS      As Long = &H10
    Const WHEEL_DELTA       As Long = 120
    
    Select Case hWnd
    Case m_Parent.hWnd
        Select Case uMsg
        Case WM_MOUSEWHEEL
            'FlexWheelScroll m_Grid, CSng(GET_WHEEL_DELTA_WPARAM(wParam)) / WHEEL_DELTA
            'Exit Function
        End Select
    Case m_Grid.hWnd
        Static ClickFixedColumn As Boolean
        Static ClickComplete    As Boolean
        Static ResizedColumn    As Long
        Static SavedScrollTrack As Long
        Select Case uMsg
        Case WM_LBUTTONDOWN
            If Not ClickFixedColumn And m_Grid.MouseRow < m_Grid.FixedRows Then
                ResizedColumn = m_Grid.MouseCol
                ClickFixedColumn = True
            End If
        Case WM_LBUTTONUP
            If ClickFixedColumn Then
                ClickFixedColumn = False
                ClickComplete = True
            End If
        Case WM_ERASEBKGND
            If ClickComplete And Not ClickFixedColumn Then
                ClickComplete = False
                RaiseEvent ColumnResized(ResizedColumn)
            End If
        Case WM_MOUSEWHEEL
            FlexWheelScroll m_Grid, CSng(GET_WHEEL_DELTA_WPARAM(wParam)) / WHEEL_DELTA
            Exit Function
        Case WM_VSCROLL
            Static Tracking As Long
            Dim si As SCROLLINFO
            Dim MaxRow As Long
            MaxRow = m_Grid.Rows - FlexPageSize(m_Grid, True)
            Select Case LOWORD(wParam)
            Case SB_LINEDOWN
                Tracking = False
                If m_Grid.TopRow >= MaxRow Then Exit Function
            Case SB_PAGEDOWN
                Tracking = False
                If m_Grid.TopRow + FlexPageSize(m_Grid, True) >= MaxRow Then
                    m_Grid.TopRow = MaxRow
                    Exit Function
                End If
            Case SB_THUMBPOSITION, SB_THUMBTRACK
                Tracking = True
                si.cbSize = LenB(si)
                si.fMask = SIF_TRACKPOS
                GetScrollInfo hWnd, SB_VERT, si
                If si.nTrackPos > UINT16_MAXVALUE Then
                    If m_Grid.ScrollTrack Then SavedScrollTrack = True
                    m_Grid.ScrollTrack = False
                ElseIf SavedScrollTrack Then
                    m_Grid.ScrollTrack = True
                End If
            Case SB_ENDSCROLL
                If Tracking Then
                    Tracking = False
                    si.cbSize = LenB(si)
                    si.fMask = SIF_TRACKPOS
                    GetScrollInfo hWnd, SB_VERT, si
                    m_Grid.TopRow = si.nTrackPos
                    If si.nTrackPos > UINT16_MAXVALUE Then
                        HighlightRows VisibleOnly:=True
                        m_Grid.TopRow = si.nTrackPos
                        FixScrollbarPageSize
                    End If
                    Exit Function
                End If
                Tracking = False
            End Select
        Case WM_HSCROLL
            Dim MaxCol As Long
            MaxCol = m_Grid.Cols - FlexPageSizeH(m_Grid, True)
            Select Case LOWORD(wParam)
            Case SB_LINERIGHT
                If m_Grid.LeftCol >= MaxCol Then Exit Function
            Case SB_PAGERIGHT
                If m_Grid.LeftCol + FlexPageSizeH(m_Grid, True) >= MaxCol Then
                    m_Grid.LeftCol = MaxCol
                    Exit Function
                End If
            End Select
        Case WM_SIZE
            FixScrollbarPageSize
        End Select
    End Select
    ISubclass_SubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

Public Sub FixScrollbarPageSize()
    If m_SetScrollOnTimer Then
        m_TimerFix.Enabled = False
        m_SetScrollOnTimer = False
        m_TimerFix.Enabled = True
    ElseIf m_TimerFix.Enabled Then
        m_TimerFix.Enabled = False
        m_TimerFix.Enabled = True
    End If
End Sub

Private Sub m_TimerFix_Timer()
    m_TimerFix.Enabled = False
    FlexSetScrollPageSize m_Grid
    m_SetScrollOnTimer = True
End Sub

Public Property Get Latency() As Long
    If Not m_TimerFix Is Nothing Then Latency = m_TimerFix.Interval
End Property

Public Property Let Latency(Value As Long)
    If Not m_TimerFix Is Nothing Then m_TimerFix.Interval = Value
End Property

Private Function FlexSetScrollPageSize(Grid As MSHFlexGrid) As Long
    Const SB_VERT As Long = 1
    Const SB_HORZ As Long = 0
    Const SIF_RANGE As Long = &H1
    Const SIF_PAGE  As Long = &H2
    
    Dim MaxRow As Long, PreserveTopRow As Boolean
    MaxRow = m_Grid.Rows - FlexPageSize(m_Grid)
    PreserveTopRow = (Grid.TopRow = MaxRow)
    
    Dim si As SCROLLINFO
    si.cbSize = LenB(si)
    si.fMask = SIF_PAGE Or SIF_RANGE
    
    si.nMin = Grid.FixedCols
    si.nMax = Grid.Cols - 1
    si.nPage = FlexPageSizeH(Grid)
    FlexSetScrollPageSize = SetScrollInfo(Grid.hWnd, SB_HORZ, si, True)
    
    si.nMin = Grid.FixedRows
    si.nMax = Grid.Rows - 1
    si.nPage = FlexPageSize(Grid, True)
    FlexSetScrollPageSize = SetScrollInfo(Grid.hWnd, SB_VERT, si, True)
    
    If PreserveTopRow Then Grid.TopRow = MaxRow
End Function

Public Property Get FlexGrid() As MSHFlexGrid
    Set FlexGrid = m_Grid
End Property

Public Property Get PageCols(Optional UseCached As Boolean = True) As Long
    PageCols = FlexPageSizeH(m_Grid, UseCached)
End Property

Public Property Get PageRows(Optional UseCached As Boolean = True) As Long
    PageRows = FlexPageSize(m_Grid, UseCached)
End Property

Private Function FlexPageSize(Grid As MSHFlexGrid, Optional UseCached As Boolean = False) As Long
    Const SM_CYHSCROLL As Long = 3
    Static StaticFlexPageSize As Long
    If UseCached Then
        FlexPageSize = StaticFlexPageSize
        Exit Function
    End If
    Dim VisibleHeight As Long
    VisibleHeight = m_Parent.ScaleY(Grid.Height, m_Parent.ScaleMode, vbTwips)
    If ScrollbarIsVisible(Grid.hWnd, vbAlignBottom) Then
        VisibleHeight = VisibleHeight - GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
    End If
    If Grid.FixedRows Then VisibleHeight = VisibleHeight - Grid.RowHeight(0) * Grid.FixedRows
    FlexPageSize = VisibleHeight \ Grid.RowHeight(Grid.Rows - 1)
    If FlexPageSize < 1 Then FlexPageSize = 1
    StaticFlexPageSize = FlexPageSize
End Function

Private Function FlexPageSizeH(Grid As MSHFlexGrid, Optional UseCached As Boolean = False) As Long
    Const SM_CXVSCROLL As Long = 2
    Static StaticFlexPageSizeH As Long
    If UseCached Then
        FlexPageSizeH = StaticFlexPageSizeH
        Exit Function
    End If
    With Grid
        .Redraw = FlexViewState(Grid)
        Dim VisibleWidth As Long
        VisibleWidth = m_Parent.ScaleX(.Width, m_Parent.ScaleMode, vbTwips)
        If ScrollbarIsVisible(Grid.hWnd, vbAlignRight) Then
            VisibleWidth = VisibleWidth - m_Parent.ScaleX(GetSystemMetrics(SM_CXVSCROLL), vbPixels, vbTwips)
        End If
        Dim CurCol As Long
        For CurCol = 0 To .FixedCols - 1
            .Col = CurCol
            VisibleWidth = VisibleWidth - (.CellWidth + FLEX_GRIDLINEWIDTH * Screen.TwipsPerPixelX)
        Next
        Dim TotalColWidth As Long
        Dim ColCount As Long
        ColCount = 0
        For CurCol = .Cols - 1 To .FixedCols Step -1
            .Col = CurCol
            TotalColWidth = TotalColWidth + (.CellWidth + FLEX_GRIDLINEWIDTH * Screen.TwipsPerPixelX)
            If TotalColWidth > VisibleWidth Then Exit For
            ColCount = ColCount + 1
        Next
        .Redraw = FlexViewState(Grid, True)
    End With
    FlexPageSizeH = ColCount
    If FlexPageSizeH < 1 Then FlexPageSizeH = 1
    StaticFlexPageSizeH = FlexPageSizeH
End Function

Private Function ScrollbarIsVisible(hWnd As Long, Optional Orientation As AlignConstants = vbAlignBottom) As Boolean
    Const GWL_STYLE As Long = -16
    Const WS_HSCROLL As Long = &H100000
    Const WS_VSCROLL As Long = &H200000
    Dim WindowStyle As Long
    WindowStyle = GetWindowLongW(hWnd, GWL_STYLE)
    If Orientation = vbAlignBottom Then
        If WindowStyle And WS_HSCROLL Then ScrollbarIsVisible = True
    Else
        If WindowStyle And WS_VSCROLL Then ScrollbarIsVisible = True
    End If
End Function

Private Function HIWORD(ByVal dwValue As Long) As Long
    HIWORD = (dwValue And &HFFFF0000) \ &H10000
End Function

Private Function LOWORD(ByVal dwValue As Long) As Long
    If dwValue And &H8000& Then
        LOWORD = dwValue Or &HFFFF0000
    Else
        LOWORD = dwValue And &HFFFF&
    End If
End Function

Private Function GET_WHEEL_DELTA_WPARAM(ByVal wParam As Long) As Long
    GET_WHEEL_DELTA_WPARAM = HIWORD(wParam)
End Function

Private Function FlexWheelScroll(Grid As MSHFlexGrid, ByVal Rotation As Single) As Long
    Const SPI_GETWHEELSCROLLLINES As Long = &H68
    Static ScrollLines As Long
    If ScrollLines = 0 Then SystemParametersInfoW SPI_GETWHEELSCROLLLINES, 0, ScrollLines, 0
    Dim StepSize As Long
    If ScrollLines < 0 Then
        StepSize = FlexPageSize(Grid, True)
    Else
        StepSize = ScrollLines
    End If
    If Grid.Rows - Grid.FixedRows = 0 Then Exit Function
    Dim NewRow As Long
    NewRow = Grid.TopRow - Rotation * StepSize
    Dim MaxRow As Long
    MaxRow = Grid.Rows - FlexPageSize(Grid, True)
    If NewRow > MaxRow Then
        Grid.TopRow = MaxRow
    ElseIf NewRow < Grid.FixedRows Then
        Grid.TopRow = Grid.FixedRows
    Else
        Grid.TopRow = NewRow
    End If
End Function

Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property

Public Property Let Redraw(Value As Boolean)
    m_Redraw = Value
    m_Grid.Redraw = m_Redraw
End Property

Public Function SaveState() As Long
    SaveState = FlexViewState(m_Grid)
End Function

Public Function RestoreState() As Long
    RestoreState = FlexViewState(m_Grid, True)
End Function

Private Function FlexViewState(Grid As MSHFlexGrid, Optional Restore As Boolean = False) As Long
    Static LastCol As Long
    Static LastRow As Long
    Static LastRowSel As Long
    Static LastColSel As Long
    Static LastTopRow As Long
    Static LastLeftCol As Long
    Static LastFillStyle As Long
    On Error Resume Next
    
    If Restore Then
        Grid.Col = LastCol
        Grid.Row = LastRow
        Grid.RowSel = LastRowSel
        Grid.ColSel = LastColSel
        Grid.LeftCol = LastLeftCol
        Grid.FillStyle = LastFillStyle
        Grid.TopRow = LastTopRow
        FlexViewState = m_Redraw
    Else
        LastCol = Grid.Col
        LastRow = Grid.Row
        LastRowSel = Grid.RowSel
        LastColSel = Grid.ColSel
        LastTopRow = Grid.TopRow
        LastLeftCol = Grid.LeftCol
        LastFillStyle = Grid.FillStyle
        FlexViewState = Restore
    End If
End Function

Public Function ClearCol(Col As Long, Optional ColWidth As Long = -1) As Long
    With m_Grid
        .ColWidth(Col) = ColWidth
        If .Rows - .FixedRows = 0 Then Exit Function
        .Redraw = FlexViewState(m_Grid)
        .FillStyle = flexFillRepeat
        .Col = Col
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .Text = vbNullString
        .Redraw = FlexViewState(m_Grid, True)
    End With
End Function

Public Function NumberRows(Optional Col As Long, Optional StartAt As Long = 1, Optional CountBy As Long = 1) As Long
    With m_Grid
        If .Rows - .FixedRows = 0 Then Exit Function
        .Redraw = FlexViewState(m_Grid)
        If Col < .FixedCols Then
            TextWidthFont = .FontFixed
        Else
            TextWidthFont = .Font
        End If
        Dim ColWidth As Long
        ColWidth = m_Parent.TextWidth(CStr(StartAt + (.Rows - .FixedRows - 1) * CountBy))
        ColWidth = m_Parent.ScaleX(ColWidth, m_Parent.ScaleMode, vbTwips)
        ColWidth = ColWidth + 2 * (FLEX_CELLPADDING + FLEX_GRIDLINEWIDTH) * Screen.TwipsPerPixelX
        If ColWidth > .RowHeight(0) Then
            .ColWidth(Col) = ColWidth
        Else
            .ColWidth(Col) = .RowHeight(0)
        End If
        Dim Row As Long
        Dim RowLabel As Long
        RowLabel = StartAt
        For Row = .FixedRows To .Rows - 1
            .TextMatrix(Row, Col) = RowLabel
            RowLabel = RowLabel + CountBy
        Next
        .Redraw = FlexViewState(m_Grid, True)
    End With
End Function

Public Property Let HighlightColor(Optional Modulus As Long = 2, Value As Long)
    Dim CurHighlight As Long
    If UBound(m_HighlightColor) = 0 Then
        ReDim m_HighlightColor(0 To 1) As THighlightColor
        CurHighlight = 1
    Else
        For CurHighlight = UBound(m_HighlightColor) To 1 Step -1
            If m_HighlightColor(CurHighlight).Modulus = Modulus Then Exit For
        Next
        If CurHighlight = 0 Then
            ReDim Preserve m_HighlightColor(0 To UBound(m_HighlightColor) + 1) As THighlightColor
            CurHighlight = UBound(m_HighlightColor)
        End If
    End If
    m_HighlightColor(CurHighlight).Modulus = Modulus
    m_HighlightColor(CurHighlight).Color = Value
    HighlightRows Not m_Grid.ScrollTrack
End Property

Public Property Get HighlightColor(Optional Modulus As Long = 2) As Long
    Dim CurHighlight As Long
    For CurHighlight = UBound(m_HighlightColor) To 1 Step -1
        If m_HighlightColor(CurHighlight).Modulus = Modulus Then Exit For
    Next
    If CurHighlight = 0 Then Err.Raise ERR_SUBSCRIPTOUTOFRANGE
    HighlightColor = m_HighlightColor(CurHighlight).Color
End Property

Public Function ClearHighlightColors()
    ReDim m_HighlightColor(0 To 0) As THighlightColor
    With m_Grid
        If .Rows - .FixedRows = 0 Then Exit Function
        .Redraw = FlexViewState(m_Grid)
        .FillStyle = flexFillRepeat
        .Col = .FixedCols
        .Row = .FixedRows
        .ColSel = .Cols - 1
        .RowSel = .Rows - 1
        .CellBackColor = .BackColor
        .Redraw = FlexViewState(m_Grid, True)
    End With
End Function

Public Function HighlightRows(Optional VisibleOnly As Boolean = False) As Long
    If UBound(m_HighlightColor) = 0 Then Exit Function
    With m_Grid
        If .Rows - .FixedRows = 0 Then Exit Function
        .Redraw = FlexViewState(m_Grid)
        .FillStyle = flexFillRepeat
        .Col = .FixedCols
        Dim StartRow As Long, EndRow As Long, LastRow As Long
        LastRow = .Rows - .FixedRows
        If VisibleOnly Then
            StartRow = 1 + .TopRow - .FixedRows
            EndRow = StartRow + FlexPageSize(m_Grid, True) + 1
            If EndRow > LastRow Then EndRow = LastRow
        Else
            StartRow = 1
            EndRow = .Rows - .FixedRows
            If EndRow > UINT16_MAXVALUE Then EndRow = UINT16_MAXVALUE
        End If
        Dim Row As Long
        Dim AltRow As Long
        If UBound(m_HighlightColor) = 1 Then If m_HighlightColor(1).Modulus = 2 Then AltRow = 1
        For Row = StartRow To EndRow
            .Row = .FixedRows + Row - 1
            .ColSel = .Cols - 1
            If AltRow Then
                If (Row And 1) = 0 Then .CellBackColor = m_HighlightColor(1).Color
            Else
                Dim CurHighlight As Long
                For CurHighlight = 1 To UBound(m_HighlightColor)
                    If Row Mod m_HighlightColor(CurHighlight).Modulus = 0 Then
                        .CellBackColor = m_HighlightColor(CurHighlight).Color
                    End If
                Next
            End If
        Next
        .Redraw = FlexViewState(m_Grid, True)
    End With
End Function

Public Function HighlightRow(Row As Long, Color As Long, Optional StartCol As Long = -1, Optional EndCol As Long = -1) As Long
    With m_Grid
        If Row > .Rows - 1 Then Exit Function
        If StartCol < 0 Then StartCol = .FixedCols
        If EndCol < 0 Then EndCol = .Cols - 1
        If EndCol - StartCol <= 0 Then Exit Function
        .Redraw = FlexViewState(m_Grid)
        .FillStyle = flexFillRepeat
        .Row = Row
        .Col = StartCol
        .ColSel = EndCol
        .CellBackColor = Color
        .Redraw = FlexViewState(m_Grid, True)
    End With
End Function

Public Function HighlightCol(Col As Long, Color As Long, Optional StartRow As Long = -1, Optional EndRow As Long = -1) As Long
    With m_Grid
        If Col > .Cols - 1 Then Exit Function
        If StartRow < 0 Then StartRow = .FixedRows
        If EndRow < 0 Then EndRow = .Rows - 1
        If EndRow - StartRow <= 0 Then Exit Function
        .Redraw = FlexViewState(m_Grid)
        .FillStyle = flexFillRepeat
        .Col = Col
        .Row = StartRow
        .RowSel = EndRow
        .CellBackColor = Color
        .Redraw = FlexViewState(m_Grid, True)
    End With
End Function


Private Property Let TextWidthFont(Value As StdFont)
    With m_Parent.Font
        .Name = Value.Name
        .Size = Value.Size
        .Bold = Value.Bold
    End With
End Property

Public Function AutoSizeColumns(Optional VisibleOnly As Boolean = True, Optional StartCol As Long = -1, Optional EndCol As Long = -1) As Long
    With m_Grid
        If .Cols <= .FixedCols Then Exit Function
        Dim TopRow As Long
        TopRow = .TopRow
        .Redraw = FlexViewState(m_Grid)
        If StartCol < 0 Then StartCol = .FixedCols
        If EndCol < 0 Then EndCol = .Cols - 1
        Dim StartRow As Long, EndRow As Long, LastRow As Long
        LastRow = .Rows - 1 - .FixedRows
        If VisibleOnly Then
            StartRow = .TopRow
            EndRow = .TopRow + FlexPageSize(m_Grid, True) + 1
            If EndRow > LastRow Then EndRow = LastRow
        Else
            StartRow = .FixedRows
            EndRow = .Rows - 1
        End If
        Dim Width As Single
        Dim MaxWidth As Single
        Dim CurCol As Long
        For CurCol = StartCol To EndCol
            MaxWidth = 0
            TextWidthFont = .FontFixed
            Dim CurRow As Long
            For CurRow = 0 To .FixedRows - 1
                Width = m_Parent.TextWidth(.TextMatrix(CurRow, CurCol))
                If Width > MaxWidth Then MaxWidth = Width
            Next
            If .Rows > .FixedRows Then
                TextWidthFont = .Font
                For CurRow = StartRow To EndRow
                    Width = m_Parent.TextWidth(.TextMatrix(CurRow, CurCol))
                    If Width > MaxWidth Then MaxWidth = Width
                Next
            End If
            MaxWidth = m_Parent.ScaleX(MaxWidth, m_Parent.ScaleMode, vbTwips)
            If MaxWidth = 0 Then
                If CurCol < .FixedCols Then
                    .ColWidth(CurCol) = .RowHeight(0)
                Else
                    .ColWidth(CurCol) = -1
                End If
            Else
                .ColWidth(CurCol) = MaxWidth + 2 * (FLEX_CELLPADDING + FLEX_GRIDLINEWIDTH) * Screen.TwipsPerPixelX
            End If
        Next
        .Redraw = FlexViewState(m_Grid, True)
        If TopRow > 0 Then .TopRow = TopRow
    End With
End Function
