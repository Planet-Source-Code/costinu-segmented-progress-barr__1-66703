VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmMain 
   Caption         =   "SGrid 2.0 Example"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2310
      Top             =   0
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdMain 
      Height          =   3615
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DisableIcons    =   -1  'True
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Start"
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1635
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IGridCellOwnerDraw

Dim clBar As clPrgBar
Const NbSegments = 30

Private Sub Form_Load()
    Set grdMain.OwnerDrawImpl = Me
End Sub

Private Sub IGridCellOwnerDraw_Draw( _
      cell As cGridCell, _
      ByVal lHDC As Long, _
      ByVal eDrawStage As ECGDrawStage, _
      ByVal lLeft As Long, ByVal lTop As Long, _
      ByVal lRight As Long, ByVal lBottom As Long, _
      bSkipDefault As Boolean _
   )
   
   Dim mPrBar As clPrgBar
   
   If (eDrawStage = ecgBeforeIconAndText) Then
      If (cell.Column = 3) Then
        
        
        
        If cell.Row = 1 Then
            clBar.DrawHdc = lHDC
            clBar.DrawLeft = lLeft
            clBar.DrawTop = lTop
            clBar.DrawHeight = lBottom - lTop
            clBar.DrawWidth = lRight - lLeft
            
            clBar.Redraw
        Else
            Set mPrBar = New clPrgBar
            
            mPrBar.DrawHdc = lHDC
            mPrBar.DrawLeft = lLeft
            mPrBar.DrawTop = lTop
            mPrBar.DrawHeight = lBottom - lTop
            mPrBar.DrawWidth = lRight - lLeft
            mPrBar.DrawText = True
            mPrBar.ProgressSegments.Add 1, clBar.ProgressSegments(cell.Row - 1).MaxValue, clBar.ProgressSegments(cell.Row - 1).Value
            
            If clBar.ProgressSegments(cell.Row - 1).MaxValue = clBar.ProgressSegments(cell.Row - 1).Value Then
                mPrBar.ProgressColor = 1361931
            End If
            
            mPrBar.Redraw
            
            Set mPrBar = Nothing
        End If
         bSkipDefault = True
      End If
   End If
End Sub


Private Sub InitBar()
    Dim i As Long
    Dim SegLen As Long
    
    Set clBar = New clPrgBar
    
    With clBar
        
'        .DrawTop = 0
'        .DrawLeft = 0
'        .DrawWidth = picBar.ScaleWidth
'        .DrawHeight = picBar.ScaleHeight
'
'        .DrawText = True
'        .DrawTotalValueBar = True
        
        For i = 1 To NbSegments
            SegLen = Int(Rnd * 80) + 20
            clBar.ProgressSegments.Add 1, SegLen, 1
        Next i
        
        clBar.DrawText = False
        clBar.ProgressColor = 16737894
        
    End With
End Sub

Private Sub cmdAction_Click()
    If tmrProgress.Enabled Then
        tmrProgress.Enabled = False
        cmdAction.Caption = "Start"
    Else
        InitBar
        InitGrid
        cmdAction.Caption = "Stop"
        tmrProgress.Enabled = True
    End If
End Sub

Private Sub IncreaseValues()
    Dim i As Integer
    Dim NbCompleted As Integer
    
    NbCompleted = 0
    
    For i = 1 To clBar.ProgressSegments.Count
        If clBar.ProgressSegments(i).Value = clBar.ProgressSegments(i).MaxValue Then
            NbCompleted = NbCompleted + 1
        Else
            clBar.ProgressSegments(i).Value = clBar.ProgressSegments(i).Value + Int(Rnd * 2)
        End If
    Next i
    
    grdMain.Redraw = True
    
    If NbCompleted = clBar.ProgressSegments.Count Then
        clBar.ProgressColor = 5485618
        clBar.MiddleColor = 9565566
        clBar.DrawText = True
        clBar.CustomText = "Action Completed"
        
        grdMain.Redraw = True
        
        cmdAction_Click
    End If
End Sub


Private Sub tmrProgress_Timer()
    IncreaseValues
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    cmdAction.Move 100, 100
    
    grdMain.Move 100, cmdAction.Height + 200, Me.ScaleWidth - 200, Me.ScaleHeight - cmdAction.Height - 300
End Sub


Private Sub InitGrid()
    Dim i As Integer
    
    On Error Resume Next
       
    With grdMain

        .Redraw = False
        .RowMode = True
        .MultiSelect = True
        .DefaultRowHeight = 18
        .HeaderFlat = True
        .StretchLastColumnToFit = True
        
        .GroupRowBackColor = &HC0C0C0
        .AddColumn "row", "Row", , , 100
        .AddColumn "desc", "Description", , , 200
      
        .AddColumn "progres", "Progres", , , 290, , , , , , , , True
        .AllowGrouping = False
        .HideGroupingBox = True
          
    
        .Redraw = True
        
        .CellDetails 1, 1, "row #0"
        .CellDetails 1, 2, "Grad Total Row "
        
        For i = 1 To NbSegments
            .CellDetails i + 1, 1, "row #" & i
            .CellDetails i + 1, 2, "Segment Number #" & i
        Next i
        
   End With

End Sub

