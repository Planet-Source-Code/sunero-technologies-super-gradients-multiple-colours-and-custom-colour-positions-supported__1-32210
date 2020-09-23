VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuperGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       clsSuperGradient
' AUTHOR:       Sunero Technologies
' CREATED:      01/03/2002
' COPYRIGHT:    Copyright 2002 Sunero Technologies. All Rights Reserved.
'
' DESCRIPTION:
' This is a class that
' you can use to render
' multicolour gradients
' with colours at custom
' positions.
'*******************************************************************************
Option Explicit


Private hdc                 As Long ' Private copy of the hDC this class can draw on

Private Colours()           As Long ' Private colour array
Private pixelPositions()    As Long ' Private postions array

Private Const PS_SOLID      As Long = 0 ' Solid pens style for API

Private Type POINTAPI               ' API co-ordinate structure
    X                       As Long
    y                       As Long
End Type

' The APIs the class uses

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public UseColourPositions   As Boolean

'*******************************************************************************
' Attach (FUNCTION)
'
' DESCRIPTION:
' This function should be called first and foremost
' before calling any other function of this class.
' Failure to do so may give unexpected results
'*******************************************************************************
Public Function Attach(lhDC As Long)
Attribute Attach.VB_UserMemId = 1610809344

    hdc = lhDC
End Function

'*******************************************************************************
' ColourStep (FUNCTION)
'
' DESCRIPTION:
' This function blends all the colours
' in the required positions (customizable)
' and fills an array with the values
' ************************

' Note: This is the main function for blending
' the colours. You can use this to implement
' your own gradient drawing styles also. Please
' see the linear gradient function to understand
' how this function is used.
'*******************************************************************************
Private Function ColourStep(Results() As Long, lLength As Long, Optional bUsePixelPositions As Boolean = False)
    
    Dim Red1            As Long
    Dim Green1          As Long
    Dim Blue1           As Long
    
    Dim Red2            As Long
    Dim Green2          As Long
    Dim Blue2           As Long
    
    Dim Red3            As Long
    Dim Green3          As Long
    Dim Blue3           As Long
    
    Dim xLoop           As Long
    Dim pixelLoop       As Long
    
    Dim lStepR          As Single
    Dim lStepG          As Single
    Dim lStepB          As Single
    
    Dim lCurrent As Long
    Dim lLoopBlock As Long
    Dim lBlock As Long
    
    ReDim Results(0 To lLength)
    
    If bUsePixelPositions = False Then
        ReDim pixelPositions(0 To UBound(Colours))
        pixelPositions(0) = 0
        pixelPositions(UBound(pixelPositions)) = lLength
        For xLoop = 1 To UBound(pixelPositions) - 1
            pixelPositions(xLoop) = xLoop * Div(CSng(lLength), UBound(pixelPositions))
        Next xLoop
    Else
        If UBound(pixelPositions) <> UBound(Colours) Then MsgBox ("Position array and Colour array do not contain equal number of items")
    End If
    
    For xLoop = 0 To UBound(Colours) - 1
        
        lLoopBlock = pixelPositions(xLoop + 1) - pixelPositions(xLoop)
        
        Red1 = RGBRed(Colours(xLoop))
        Green1 = RGBGreen(Colours(xLoop))
        Blue1 = RGBBlue(Colours(xLoop))
        
        Red2 = RGBRed(Colours(xLoop + 1))
        Green2 = RGBGreen(Colours(xLoop + 1))
        Blue2 = RGBBlue(Colours(xLoop + 1))
        
        lStepR = Div((Red2 - Red1), CSng(lLoopBlock))
        lStepG = Div((Green2 - Green1), CSng(lLoopBlock))
        lStepB = Div((Blue2 - Blue1), CSng(lLoopBlock))
        
        For pixelLoop = 0 To lLoopBlock
            Red3 = Red1 + (pixelLoop * lStepR)
            Green3 = Green1 + (pixelLoop * lStepG)
            Blue3 = Blue1 + (pixelLoop * lStepB)
            
            Results(lBlock + pixelLoop) = RGB(Red3, Green3, Blue3)
        Next pixelLoop
        lBlock = lBlock + lLoopBlock
    Next xLoop
    
End Function

'*******************************************************************************
' Div (FUNCTION)
'
' DESCRIPTION:
' Private use function to avoid the division
' by zero error
'*******************************************************************************
Private Function Div(Numerator As Single, Denominator As Single) As Single
    If Numerator = 0 Or Denominator = 0 Then
        Div = 0
    Else
        Div = Numerator / Denominator
    End If
End Function

'*******************************************************************************
' DrawLinearGradient (FUNCTION)
'
' DESCRIPTION:
' This function is used to draw linear
' gradients. If TopToBottom is true then
' the gradient is draw horizontally or
' else it is drawn vertically.
'*******************************************************************************
Public Function DrawLinearGradient(Left As Long, Top As Long, Width As Long, Height As Long, Optional TopToBottom As Boolean = False)
Attribute DrawLinearGradient.VB_UserMemId = 1610809347
    
    Dim Results() As Long
    Dim iMax As Long
    Dim iLoop As Long
    Dim lpAPI As POINTAPI
    
    Dim hPen As Long
    Dim hOldPen As Long
    Dim lRet As Long
    
    Dim xFrom As Long
    Dim yFrom As Long
    
    Dim xTo As Long
    Dim yTo As Long
    
    If TopToBottom = True Then
        iMax = Height
    Else
        iMax = Width
    End If
    
    ColourStep Results(), iMax, UseColourPositions
    
    For iLoop = 0 To iMax
        If TopToBottom = True Then
            xFrom = Left
            xTo = Left + Width
            yFrom = Top + iLoop
            yTo = Top + iLoop
        Else
            xFrom = Left + iLoop
            xTo = Left + iLoop
            yFrom = Top
            yTo = Top + Height
        End If
        
        lRet = MoveToEx(hdc, xFrom, yFrom, lpAPI)
        'Create the colored pen and select it into the DC
        hPen = CreatePen(PS_SOLID, 1, Results(iLoop))
        hOldPen = SelectObject(hdc, hPen)
        'Draw the line
        lRet = LineTo(hdc, xTo, yTo)
        'Get the pen back out of the DC and destroy it
        lRet = SelectObject(hdc, hOldPen)
        lRet = DeleteObject(hPen)
        
    Next iLoop
    
End Function


'*******************************************************************************
' DrawRectangularGradients (FUNCTION)
'
' DESCRIPTION:
' You can draw box style gradients using this
' function. Also, this function can be used
' to draw circular gradients
'*******************************************************************************
Public Function DrawRectangularGradients(CenterX As Long, CenterY As Long, Radius As Long)
Attribute DrawRectangularGradients.VB_UserMemId = 1610809348
    Dim Results() As Long
    Dim iLoop As Long
    
    ColourStep Results(), Radius, UseColourPositions
    
    For iLoop = 0 To Radius
        DrawSquare CenterX, CenterY, iLoop, Results(iLoop)
    Next iLoop
    
End Function

'*******************************************************************************
' DrawSquare (FUNCTION)
'
' DESCRIPTION:
' Draws a square from the specified center
' and the given radius of it's circumcircle
' *****************************************
' IMPORTANT NOTE: You can replace the word
' 'Rectangle' with 'Ellipse' to draw circular
' gradients. However the quality of circular
' gradients drawn using pixels as the unit is
' not very good :(
'*******************************************************************************
Private Function DrawSquare(CenterX As Long, CenterY As Long, Radius As Long, Colour As Long)
    Dim hPen As Long
    Dim hOldPen As Long
    
    Dim X1 As Long
    Dim Y1 As Long
    Dim X2 As Long
    Dim Y2 As Long
    
    X1 = CenterX - Radius
    X2 = CenterX + Radius
    Y1 = CenterY - Radius
    Y2 = CenterY + Radius
    
    hPen = CreatePen(PS_SOLID, 1, Colour)
    hOldPen = SelectObject(hdc, hPen)
    
    Rectangle hdc, X1, Y1, X2, Y2
    
    SelectObject hdc, hOldPen
    
    DeleteObject hPen
    
End Function

'' Functions to retrieve RGB values from long

Private Function RGBBlue(RGBCol As Long) As Integer
    'Return the Blue component from an RGB C
    '     olor
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function


Private Function RGBGreen(RGBCol As Long) As Integer
    'Return the Green component from an RGB
    '     Color
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Private Function RGBRed(RGBCol As Long) As Integer
    'Return the Red component from an RGB Co
    '     lor
    RGBRed = RGBCol And &HFF
End Function

'*******************************************************************************
' SetColourPositions (FUNCTION)
'
' DESCRIPTION:
' Feed this function an array containing
' the positions of the colours in pixels.
' For best results the first item should
' be zero and the last item should be the
' length of the gradient.
'*******************************************************************************
Public Function SetColourPositions(ColourPositions() As Long)
Attribute SetColourPositions.VB_UserMemId = 1610809353
    Dim iLoop As Long
    ReDim pixelPositions(0 To UBound(ColourPositions))
    
    For iLoop = 0 To UBound(ColourPositions)
        pixelPositions(iLoop) = ColourPositions(iLoop)
    Next iLoop
    
End Function

'*******************************************************************************
' SetColours (FUNCTION)
'
' DESCRIPTION:
'
'*******************************************************************************
Public Function SetColours(lColours() As Long)
Attribute SetColours.VB_UserMemId = 1610809354
    Dim iLoop As Long
    
    ReDim Colours(0 To UBound(lColours))
    
    For iLoop = 0 To UBound(lColours)
        Colours(iLoop) = lColours(iLoop)
    Next iLoop
    
End Function

