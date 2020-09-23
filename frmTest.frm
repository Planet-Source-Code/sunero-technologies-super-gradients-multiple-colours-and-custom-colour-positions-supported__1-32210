VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sunero Super Gradient Test"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSimple4 
      Caption         =   "White-Blue "
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   2700
      Width           =   1275
   End
   Begin VB.CommandButton cmdSimple5 
      Caption         =   "White-Red "
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   3060
      Width           =   1275
   End
   Begin VB.CommandButton cmdSimple6 
      Caption         =   "White-Green"
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   3420
      Width           =   1275
   End
   Begin VB.CommandButton cmdSimple3 
      Caption         =   "Black-Green"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdSimple2 
      Caption         =   "Black-Red "
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   1980
      Width           =   1275
   End
   Begin VB.CommandButton cmdSimple 
      Caption         =   "Black-Blue "
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   1620
      Width           =   1275
   End
   Begin VB.CommandButton cmdBoxed 
      Caption         =   "Random Rects"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   1275
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Positioning"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   1275
   End
   Begin VB.VScrollBar VBar 
      Enabled         =   0   'False
      Height          =   4755
      LargeChange     =   25
      Left            =   5940
      SmallChange     =   25
      TabIndex        =   3
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   1380
      ScaleHeight     =   4470
      ScaleWidth      =   4470
      TabIndex        =   2
      Top             =   180
      Width           =   4500
   End
   Begin VB.CommandButton cmdFillChrome 
      Caption         =   "Chrome Effect"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   1275
   End
   Begin VB.CommandButton cmdGold 
      Caption         =   "Gold Effect"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private xCls As New clsSuperGradient

Dim lColours() As Long
Dim lPos() As Long
Dim iLoop As Long

Private Sub cmdBoxed_Click()
    VBar.Enabled = False
    
    ReDim lColours(0 To 15)
    ReDim lPos(0)
    
    Dim iLoop As Long
    
    For iLoop = 0 To 15
        lColours(iLoop) = RGB(Rnd * 255, 0, Rnd * 128)
    Next iLoop
    
    
    xCls.UseColourPositions = False
    xCls.SetColours lColours
    xCls.SetColourPositions lPos
    xCls.DrawRectangularGradients 150, 150, 150
    
    picBox.Refresh
    
End Sub

Private Sub cmdFillChrome_Click()
    
    ReDim lColours(0 To 7)
    ReDim lPos(0 To 7)
    
    VBar.Enabled = False
    
    lColours(0) = vbWhite
    lColours(1) = RGB(128, 128, 128)
    lColours(2) = RGB(64, 64, 64)
    lColours(3) = vbWhite
    lColours(4) = RGB(64, 64, 64)
    lColours(5) = RGB(64, 64, 64)
    lColours(6) = vbBlack
    lColours(7) = vbWhite
    
        Dim xPix As Long
    
    xPix = (picBox.Height / 100)
    
    lPos(0) = 0
    lPos(1) = 10 * xPix
    lPos(2) = 10 * xPix
    lPos(3) = 60 * xPix
    lPos(4) = 70 * xPix
    lPos(5) = 80 * xPix
    lPos(6) = 80 * xPix
    lPos(7) = picBox.Height
    
    RequestRender
    
End Sub

Private Sub cmdGold_Click()
    ' Redimension the arrays from 0 to one less than the
    'actual number of colours
    
    ReDim lColours(0 To 7)
    ReDim lPos(0 To 7)
    
    VBar.Enabled = False
    
    ' Fill the colour values
    lColours(0) = RGB(226, 200, 139)
    lColours(1) = RGB(192, 148, 48)
    lColours(2) = RGB(192, 148, 48)
    lColours(3) = RGB(238, 222, 185)
    lColours(4) = RGB(192, 148, 48)
    lColours(5) = RGB(192, 148, 48)
    lColours(6) = RGB(136, 96, 24)
    lColours(7) = RGB(226, 200, 139)
    
    Dim xPix As Long
    
    xPix = (picBox.Height / 100)
    
    ' Fill the required positions
    lPos(0) = 0
    lPos(1) = 10 * xPix
    lPos(2) = 10 * xPix
    lPos(3) = 60 * xPix
    lPos(4) = 70 * xPix
    lPos(5) = 80 * xPix
    lPos(6) = 80 * xPix
    lPos(7) = picBox.Height
    
    ' Render
    RequestRender

End Sub

Private Sub cmdSimple_Click()
     DrawTwoColours vbBlack, vbBlue
End Sub

Private Sub cmdSimple2_Click()
    DrawTwoColours vbBlack, vbRed
End Sub

Private Sub cmdSimple3_Click()
    DrawTwoColours vbBlack, vbGreen
End Sub

Private Sub cmdSimple4_Click()
    DrawTwoColours vbWhite, vbBlue
End Sub

Private Sub cmdSimple5_Click()
    DrawTwoColours vbWhite, vbRed
End Sub

Private Sub cmdSimple6_Click()
    DrawTwoColours vbWhite, vbGreen
End Sub

Private Sub cmdTest_Click()
    ReDim lColours(0 To 2)
    

    lColours(0) = RGB(255, 128, 0)
    lColours(1) = vbWhite
    lColours(2) = RGB(0, 128, 0)
    
    ReDim lPos(0 To 2)
    
    lPos(0) = 0
    lPos(1) = picBox.Height / 2
    lPos(2) = picBox.Height
    
    VBar.Enabled = True
    
    VBar.Min = 10
    VBar.Max = picBox.Height - 10
    VBar.Value = (picBox.Height / 2)
    
    RequestRender
End Sub


Private Sub Form_Load()
    xCls.Attach picBox.hdc
End Sub

Public Function DrawTwoColours(Colour1 As Long, Colour2 As Long)
    'Draw two colour gradients
    
    ReDim lColours(0 To 1)
    ReDim lPos(0)
    
    VBar.Enabled = False
    
    lColours(0) = Colour1
    lColours(1) = Colour2
    
    xCls.UseColourPositions = False
    xCls.SetColours lColours
    xCls.SetColourPositions lPos
    
    xCls.DrawLinearGradient 0, 0, 300, 300, True
    picBox.Refresh
End Function

Public Function RequestRender()
       
    ' If you use colour positions
    ' Set the property below to
    ' true
    xCls.UseColourPositions = True
    
    'Set the colours
    xCls.SetColours lColours
    
    'Set the colour positions
    xCls.SetColourPositions lPos
    
    'Render linear gradients
    xCls.DrawLinearGradient 0, 0, picBox.ScaleWidth, picBox.ScaleHeight, True

    picBox.Refresh
End Function

Private Sub VBar_Change()
    lPos(1) = VBar.Value
    RequestRender
End Sub

Private Sub VBar_Scroll()
    VBar_Change
End Sub
