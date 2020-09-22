VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress Bar Example"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   150
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Start"
      Height          =   465
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   1635
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   690
      Width           =   6075
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clBar As clPrgBar
Const NbSegments = 30

Private Sub InitBar()
    Dim i As Long
    Dim SegLen As Long
    
    Set clBar = New clPrgBar
    
    With clBar
        .DrawHdc = picBar.hDC
        
        .DrawTop = 0
        .DrawLeft = 0
        .DrawWidth = picBar.ScaleWidth
        .DrawHeight = picBar.ScaleHeight
        
        .DrawText = True
        .CustomText = ""
        .ProgressColor = 16744448
        .DrawTotalValueBar = True
        
        For i = 1 To NbSegments
            SegLen = Int(Rnd * 80) + 20
            clBar.ProgressSegments.Add 1, SegLen, 1
        Next i
        
        .Redraw
        picBar.Refresh
        
    End With
End Sub

Private Sub cmdAction_Click()
    If tmrProgress.Enabled Then
        tmrProgress.Enabled = False
        cmdAction.Caption = "Start"
    Else
        InitBar
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
    
    clBar.Redraw
    picBar.Refresh
    
    If NbCompleted = clBar.ProgressSegments.Count Then
        clBar.ProgressColor = 5485618
        clBar.MiddleColor = 9565566
        clBar.DrawText = True
        clBar.CustomText = "Action Completed"
        clBar.Redraw
        picBar.Refresh
        
        cmdAction_Click
    End If
End Sub


Private Sub tmrProgress_Timer()
    IncreaseValues
End Sub
