VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clPrgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type API_POINT
        x       As Long
        y       As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (pDest As Any, _
         pSource As Any, _
         ByVal ByteLen As Long)

Private Declare Function MoveToEx Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal x As Long, _
         ByVal y As Long, _
         lpPoint As API_POINT) As Long

Private Declare Function LineTo Lib "gdi32" _
        (ByVal hDC As Long, _
         ByVal x As Long, _
         ByVal y As Long) As Long

Private Declare Function GetSysColor Lib "user32" _
        (ByVal nIndex As Long) As Long
        
'Public drawObj As Object

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public DrawHdc As Long
Public DrawWidth As Long
Public DrawHeight As Long
Public DrawLeft As Long
Public DrawTop As Long

Const def_BackColor = 14671839
Const def_MiddleColor = 16777215
Const def_AlternateBackColor = 12566463 '13619151
Const def_ProgressColor = 16744448
Const def_TotalProgressColor = 4368140

'local variable(s) to hold property value(s)
Private mvarBackColor As Long 'local copy
Private mvarMiddleColor As Long 'local copy
Private mvarAlternateBackColor As Long 'local copy
Private mvarProgressColor As Long 'local copy
Private mvarTotalProgressColor As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarProgressSegments As Segments 'local copy

Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Private Const GRADIENT_FILL_TRIANGLE = &H2&
Private Declare Function CreateSolidBrush Lib "gdi32" ( _
   ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" ( _
   ByVal hDC As Long, lpRect As RECT, _
   ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
   ByVal hObject As Long) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" ( _
   ByVal OLE_COLOR As Long, _
   ByVal HPALETTE As Long, _
   pccolorref As Long) As Long

Const CLR_INVALID = -1

Private Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum

Private Declare Function DrawTextA Lib "user32" ( _
    ByVal hDC As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) As Long


'local variable(s) to hold property value(s)
Private mvarDrawTotalValueBar As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarTextColor As Long 'local copy
Private mvarDrawText As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarCustomText As String 'local copy
Public Property Let CustomText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CustomText = 5
    mvarCustomText = vData
End Property


Public Property Get CustomText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CustomText
    CustomText = mvarCustomText
End Property



Public Property Let DrawText(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawText = 5
    mvarDrawText = vData
End Property


Public Property Get DrawText() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawText
    DrawText = mvarDrawText
End Property



Public Property Let TextColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TextColor = 5
    mvarTextColor = vData
End Property


Public Property Get TextColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TextColor
    TextColor = mvarTextColor
End Property




Public Property Let DrawTotalValueBar(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DrawTotalValueBar = 5
    mvarDrawTotalValueBar = vData
End Property


Public Property Get DrawTotalValueBar() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DrawTotalValueBar
    DrawTotalValueBar = mvarDrawTotalValueBar
End Property




Private Sub GradientFillRect( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal oStartColor As OLE_COLOR, _
      ByVal oEndColor As OLE_COLOR, _
      ByVal eDir As GradientFillRectType _
   )
Dim hBrush As Long
Dim lStartColor As Long
Dim lEndColor As Long
Dim lR As Long
   
   ' Use GradientFill:
   lStartColor = TranslateColor(oStartColor)
   lEndColor = TranslateColor(oEndColor)

   Dim tTV(0 To 1) As TRIVERTEX
   Dim tGR As GRADIENT_RECT
   
   setTriVertexColor tTV(0), lStartColor
   tTV(0).x = tR.left
   tTV(0).y = tR.top
   setTriVertexColor tTV(1), lEndColor
   tTV(1).x = tR.right
   tTV(1).y = tR.bottom
   
   tGR.UpperLeft = 0
   tGR.LowerRight = 1
   
   GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
      
   If (Err.Number <> 0) Then
      ' Fill with solid brush:
      hBrush = CreateSolidBrush(TranslateColor(oEndColor))
      FillRect lHDC, tR, hBrush
      DeleteObject hBrush
   End If
   
End Sub

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
Dim lRed As Long
Dim lGreen As Long
Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   setTriVertexColorComponent tTV.Red, lRed
   setTriVertexColorComponent tTV.Green, lGreen
   setTriVertexColorComponent tTV.Blue, lBlue
End Sub
Private Sub setTriVertexColorComponent( _
   ByRef iColor As Integer, _
   ByVal lComponent As Long _
   )
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub

Private Function TranslateColor( _
    ByVal oClr As OLE_COLOR, _
    Optional hPal As Long = 0 _
    ) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


Public Property Set ProgressSegments(ByVal vData As Segments)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ProgressSegments = Form1
    Set mvarProgressSegments = vData
End Property


Public Property Get ProgressSegments() As Segments
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProgressSegments
    Set ProgressSegments = mvarProgressSegments
End Property



Public Property Let TotalProgressColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TotalProgressColor = 5
    mvarTotalProgressColor = vData
End Property


Public Property Get TotalProgressColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalProgressColor
    TotalProgressColor = mvarTotalProgressColor
End Property



Public Property Let ProgressColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ProgressColor = 5
    mvarProgressColor = vData
End Property


Public Property Get ProgressColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProgressColor
    ProgressColor = mvarProgressColor
End Property



Public Property Let AlternateBackColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AlternateBackColor = 5
    mvarAlternateBackColor = vData
End Property


Public Property Get AlternateBackColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AlternateBackColor
    AlternateBackColor = mvarAlternateBackColor
End Property



Public Property Let MiddleColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MiddleColor = 5
    mvarMiddleColor = vData
End Property


Public Property Get MiddleColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MiddleColor
    MiddleColor = mvarMiddleColor
End Property


Public Property Let BackColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BackColor = 5
    mvarBackColor = vData
    Redraw
End Property


Public Property Get BackColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BackColor
    BackColor = mvarBackColor
End Property

Public Sub Redraw()
    Dim i As Integer
    Dim MaxValue As Long
    Dim mSegment As Segment
    Dim Alternate As Boolean
    Dim StartPoint As Double
    Dim Percent As Double
    Dim SegLen As Double
    Dim ValLen As Double
    Dim TotalValue As Double
    
    On Error GoTo errHandler
        
    MaxValue = 0
    TotalValue = 0
    
   
    For i = 1 To mvarProgressSegments.Count
        Set mSegment = mvarProgressSegments(i)
        MaxValue = MaxValue + mSegment.MaxValue - mSegment.MinValue
        TotalValue = TotalValue + mSegment.Value - mSegment.MinValue
    Next i
    
    Alternate = False
    
    StartPoint = 0
        
    If MaxValue = 0 Then
        Percent = 0
    Else
        Percent = DrawWidth / MaxValue
    End If
        
    For i = 1 To mvarProgressSegments.Count
        Set mSegment = mvarProgressSegments(i)
        
        
        
        SegLen = (mSegment.MaxValue - mSegment.MinValue) * Percent
        ValLen = (mSegment.Value - mSegment.MinValue) * Percent
        
        If SegLen = 0 Then
            SegLen = 1
        End If
        
        If Alternate Then
            DrawBar Int(StartPoint), Int(SegLen) + 1, mvarAlternateBackColor, mvarMiddleColor
        Else
            DrawBar Int(StartPoint), Int(SegLen) + 1, mvarBackColor, mvarMiddleColor
        End If
        
        DrawBar Int(StartPoint), Int(ValLen) + 1, mvarProgressColor, mvarMiddleColor
        
        Alternate = Not Alternate
        
        StartPoint = StartPoint + SegLen
    Next i
    
    DrawValueBar TotalValue * Percent, mvarTotalProgressColor, mvarTotalProgressColor
    
errHandler:
End Sub

Private Sub DrawBar(fromX As Long, toX As Long, fromColor As Long, toColor As Long)
    Dim r As RECT

    r.left = fromX + DrawLeft
    r.right = fromX + DrawLeft + toX
    r.top = DrawTop
    r.bottom = DrawTop + DrawHeight / 2
    
    GradientFillRect DrawHdc, r, fromColor, toColor, GRADIENT_FILL_RECT_V

    r.left = fromX + DrawLeft
    r.right = fromX + DrawLeft + toX
    r.top = DrawTop + DrawHeight / 2
    r.bottom = DrawTop + DrawHeight
    
    GradientFillRect DrawHdc, r, toColor, fromColor, GRADIENT_FILL_RECT_V
    
End Sub

Private Sub DrawValueBar(toX As Long, fromColor As Long, toColor As Long)
    Dim fromX As Long
    Dim r As RECT
    
    On Error Resume Next
    
    fromX = 0
    
    If mvarDrawTotalValueBar Then
        r.left = fromX + DrawLeft
        r.right = fromX + DrawLeft + toX
        r.top = DrawTop
        r.bottom = DrawTop + 2

        GradientFillRect DrawHdc, r, fromColor, toColor, GRADIENT_FILL_RECT_V
    End If
    
    If mvarDrawText Then
        SetTextColor DrawHdc, mvarTextColor
        
        r.top = DrawTop + 2
        r.bottom = DrawTop + DrawHeight
        r.left = fromX + DrawLeft
        r.right = fromX + DrawLeft + DrawWidth
        
        If mvarCustomText = "" Then
            DrawTextA DrawHdc, Int((toX - fromX) * 100 / DrawWidth) & "%", -1, r, 1
        Else
            DrawTextA DrawHdc, mvarCustomText, -1, r, 1
        End If
    End If


End Sub

Private Sub Class_Initialize()
    mvarAlternateBackColor = def_AlternateBackColor
    mvarBackColor = def_BackColor
    mvarMiddleColor = def_MiddleColor
    mvarProgressColor = def_ProgressColor
    mvarTotalProgressColor = def_TotalProgressColor
    Set mvarProgressSegments = New Segments
End Sub
