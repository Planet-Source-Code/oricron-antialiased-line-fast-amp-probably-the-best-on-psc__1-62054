Attribute VB_Name = "modAALine"
Option Explicit

Private Type eCoordinate 'we need this, to calculate each point exact position
    x As Double
    y As Double
End Type

Private Type tPix       'point settings
    COR As eCoordinate
    Alpha As Integer
End Type

Public Type ColorAndAlpha
    r                   As Byte
    G                   As Byte
    b                   As Byte
    A                   As Byte
End Type

'Needed API
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal Destination As Any, ByVal Source As Any, ByVal Length As Long)

'We call only this sub, when we need to draw a line
Public Sub AALine(hdc As Long, X1 As Double, Y1 As Double, x2 As Double, Y2 As Double, Color As OLE_COLOR, Border As Double, Alpha As Double)

'hdc, object.hdc where we draw the line
'x1,x2,y1,y2 - start and end of the line
'Color - line color
'Border -line thicknes (border size)
'Alpha - opacity from 0 - 1 (double ex. 0.73 = 73 %)


If X1 = x2 And Y1 = Y2 Then Exit Sub 'we cannot drawjust a single point

'Needed variables
Dim iCnt As Double
Dim iCnt2 As Long
Dim Pp() As tPix
ReDim Pp(0)

Dim iX1 As Double
Dim iX2 As Double
Dim iY1 As Double
Dim iY2 As Double
Dim ky As Double
Dim kx As Double

Dim A As Double
Dim a2 As Integer
Dim a3 As Double

Dim kk  As Double

Dim PPX() As tPix
ReDim PPX(0)

Dim ICNT3 As Long
Dim k As Long
Dim iCnt0 As Long
Dim h As Double

'And here we go...
If Abs(X1 - x2) >= Abs(Y1 - Y2) Then 'we separate this in two categorys - acordingly to the angle... - the second part is the same as the first, just x=y and y = x
    
    'We set the new coordinates for easier drawing
    If X1 > x2 Then
        iX1 = x2
        iX2 = X1
        iY1 = Y2
        iY2 = Y1
    Else
        iX1 = X1
        iX2 = x2
        iY1 = Y1
        iY2 = Y2
    End If
    
    'nedded - so it draws to exact point as meant when calling the function

    kk = Abs(iY1 - iY2) / Abs(iX1 - iX2)
    
    iX2 = iX2 + 1
    iY2 = iY2 + kk

    
    For iCnt = iX1 To iX2 Step 0.25 'we use 1/4 of a point, so that we go 4 times troug each point -> AntiAliasing = 4x
        ky = (iCnt - iX1) * kk 'point y position
        
        ReDim Preserve Pp(UBound(Pp) + 1)
        
        Pp(UBound(Pp)).COR.x = iCnt
        
        'folowing depends on the angle...
        If iY1 > iY2 Then
            Pp(UBound(Pp)).COR.y = iY1 - ky
        Else
            Pp(UBound(Pp)).COR.y = iY1 + ky
        End If
    Next
    
    h = Border 'border size
    
    If iY1 < iY2 Then 'again depends on the angle (past elese is just a slight differnet you'll figure it out...
        For iCnt0 = 1 To UBound(Pp) - 1 Step 4 'we take it by pixels (since AntiAliasing = 4x)
            For ICNT3 = Int(h - 0.001) + 1 To -1 Step -1 'we go troug to draw different thicknes)
                a2 = 0
                k = UBound(PPX) + 1
                ReDim Preserve PPX(k)
                A = 0 'this wiil be the alpha value of the pixel (not including the overall alpha of the line!)
                For iCnt2 = 0 To 3
                    If ICNT3 = Int(h - 0.001 + 1) Then 'botm par of the line
                        a3 = Pp(iCnt0 + iCnt2).COR.y - Int(Pp(iCnt0 + 3).COR.y)
    
                        If a3 > h Then a3 = h
        
                        If Int(Pp(iCnt0 + iCnt2).COR.y) < Int(Pp(iCnt0 + 3).COR.y) Then
                            a3 = 0
                        End If
                        
                    ElseIf ICNT3 = 0 Then 'top part of the line
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.y - Int(Pp(iCnt0 + iCnt2).COR.y))
                        
                        If a3 > h Then a3 = h
                        If Int(Pp(iCnt0 + iCnt2).COR.y) < Int(Pp(iCnt0 + 3).COR.y) Then
                            If h < 1 Then
                                a3 = h
                            Else
                                a3 = 1
                            End If
                        End If
                    ElseIf ICNT3 = -1 Then 'top part of the line - just aliasing corrections
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.y - Int(Pp(iCnt0 + iCnt2).COR.y))
                        
                        If a3 > h Then a3 = h
                        If Int(Pp(iCnt0 + iCnt2).COR.y) >= Int(Pp(iCnt0 + 3).COR.y) Then
                            a3 = 0
                        End If
                    Else 'middle part of the line (if thicknes > 1)
                        a3 = 1
                    End If
    
                    A = A + a3
                    
                    If iCnt2 = 3 Then  'if the pas is = pixel x position then we set the alpha)
                        PPX(k).COR.x = Int(Pp(iCnt0).COR.x)
                        PPX(k).COR.y = Int(Pp(iCnt0 + 3).COR.y) + ICNT3 - Int(h / 2)
                        PPX(k).Alpha = 255 * A / 4
                    End If
                Next iCnt2
            Next ICNT3
        Next iCnt0
    Else 'same as above just depends on the anlge...
        For iCnt0 = 1 To UBound(Pp) - 1 Step 4
            For ICNT3 = Int(h - 0.001) + 1 To -1 Step -1
                a2 = 0
                k = UBound(PPX) + 1
                ReDim Preserve PPX(k)
                A = 0
                For iCnt2 = 0 To 3
                    If ICNT3 = Int(h - 0.001 + 1) Then
                        a3 = Pp(iCnt0 + iCnt2).COR.y - Int(Pp(iCnt0).COR.y)
    
                        If a3 > h Then a3 = h
        
                        If Int(Pp(iCnt0 + iCnt2).COR.y) < Int(Pp(iCnt0).COR.y) Then
                            a3 = 0
                        End If
                        
                    ElseIf ICNT3 = 0 Then
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.y - Int(Pp(iCnt0 + iCnt2).COR.y))
                        
                        If a3 > h Then a3 = h
                        
                        If Int(Pp(iCnt0 + iCnt2).COR.y) < Int(Pp(iCnt0).COR.y) Then
                            If h < 1 Then
                                a3 = h
                            Else
                                a3 = 1
                            End If
                        End If
                    ElseIf ICNT3 = -1 Then
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.y - Int(Pp(iCnt0 + iCnt2).COR.y))
                        
                        If a3 > h Then a3 = h
                        If Int(Pp(iCnt0 + iCnt2).COR.y) >= Int(Pp(iCnt0).COR.y) Then
                            a3 = 0
                        End If
                    Else
                        a3 = 1
                    End If
    
                    A = A + a3
                    
                    If iCnt2 = 3 Then
                        PPX(k).COR.x = Int(Pp(iCnt0).COR.x)
                        PPX(k).COR.y = Int(Pp(iCnt0).COR.y) + ICNT3 - Int(h / 2)
                        PPX(k).Alpha = 255 * A / 4
                    End If
                Next iCnt2
            Next ICNT3
        Next iCnt0
    End If
Else

'just the same as the first part - only turned around -> x=y and y=x
    If Y1 > Y2 Then
        iY1 = Y2
        iY2 = Y1
        iX1 = x2
        iX2 = X1
    Else
        iY1 = Y1
        iY2 = Y2
        iX1 = X1
        iX2 = x2
    End If
    
    kk = Abs(iX1 - iX2) / Abs(iY1 - iY2)
    
    iY2 = iY2 + 1
    iX2 = iX2 + kk
    

    For iCnt = iY1 To iY2 Step 0.25
            kx = (iCnt - iY1) * kk
            
            ReDim Preserve Pp(UBound(Pp) + 1)
            
            Pp(UBound(Pp)).COR.y = iCnt
            
            If iX1 > iX2 Then
                Pp(UBound(Pp)).COR.x = iX1 - kx
            Else
                Pp(UBound(Pp)).COR.x = iX1 + kx
            End If
    Next
    
    h = Border

    If iX1 < iX2 Then
        For iCnt0 = 1 To UBound(Pp) - 1 Step 4
            For ICNT3 = Int(h - 0.001) + 1 To -1 Step -1
                a2 = 0
                k = UBound(PPX) + 1
                ReDim Preserve PPX(k)
                A = 0
                For iCnt2 = 0 To 3
                    If ICNT3 = Int(h - 0.001 + 1) Then
                        a3 = Pp(iCnt0 + iCnt2).COR.x - Int(Pp(iCnt0 + 3).COR.x)
    
                        If a3 > h Then a3 = h
        
                        If Int(Pp(iCnt0 + iCnt2).COR.x) < Int(Pp(iCnt0 + 3).COR.x) Then
                            a3 = 0
                        End If
                        
                    ElseIf ICNT3 = 0 Then
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.x - Int(Pp(iCnt0 + iCnt2).COR.x))
                        
                        If a3 > h Then a3 = h
                        If Int(Pp(iCnt0 + iCnt2).COR.x) < Int(Pp(iCnt0 + 3).COR.x) Then
                            If h < 1 Then
                                a3 = h
                            Else
                                a3 = 1
                            End If
                        End If
                    ElseIf ICNT3 = -1 Then
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.x - Int(Pp(iCnt0 + iCnt2).COR.x))
                        
                        If a3 > h Then a3 = h
                        If Int(Pp(iCnt0 + iCnt2).COR.x) >= Int(Pp(iCnt0 + 3).COR.x) Then
                            a3 = 0
                        End If
                    Else
                        a3 = 1
                    End If
    
                    A = A + a3
                    
                    If iCnt2 = 3 Then
                        PPX(k).COR.y = Int(Pp(iCnt0).COR.y)
                        PPX(k).COR.x = Int(Pp(iCnt0 + 3).COR.x) + ICNT3 - Int(h / 2)
                        PPX(k).Alpha = 255 * A / 4
                    End If
                Next iCnt2
            Next ICNT3
        Next iCnt0
    Else
        For iCnt0 = 1 To UBound(Pp) - 1 Step 4
            For ICNT3 = Int(h - 0.001) + 1 To -1 Step -1
                a2 = 0
                k = UBound(PPX) + 1
                ReDim Preserve PPX(k)
                A = 0
                For iCnt2 = 0 To 3
                    If ICNT3 = Int(h - 0.001 + 1) Then
                        a3 = Pp(iCnt0 + iCnt2).COR.x - Int(Pp(iCnt0).COR.x)
    
                        If a3 > h Then a3 = h
        
                        If Int(Pp(iCnt0 + iCnt2).COR.x) < Int(Pp(iCnt0).COR.x) Then
                            a3 = 0
                        End If
                        
                    ElseIf ICNT3 = 0 Then
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.x - Int(Pp(iCnt0 + iCnt2).COR.x))
                        
                        If a3 > h Then a3 = h
                        
                        If Int(Pp(iCnt0 + iCnt2).COR.x) < Int(Pp(iCnt0).COR.x) Then
                            If h < 1 Then
                                a3 = h
                            Else
                                a3 = 1
                            End If
                        End If
                    ElseIf ICNT3 = -1 Then
                        a3 = 1 - (Pp(iCnt0 + iCnt2).COR.x - Int(Pp(iCnt0 + iCnt2).COR.x))
                        
                        If a3 > h Then a3 = h
                        If Int(Pp(iCnt0 + iCnt2).COR.x) >= Int(Pp(iCnt0).COR.x) Then
                            a3 = 0
                        End If
                    Else
                        a3 = 1
                    End If
    
                    A = A + a3
                    
                    If iCnt2 = 3 Then
                        PPX(k).COR.y = Int(Pp(iCnt0).COR.y)
                        PPX(k).COR.x = Int(Pp(iCnt0).COR.x) + ICNT3 - Int(h / 2)
                        PPX(k).Alpha = 255 * A / 4
                    End If
                Next iCnt2
            Next ICNT3
        Next iCnt0
    End If
End If

'finaly we go trough every calculated pixel and draw it
For iCnt0 = 1 To UBound(PPX)
    SetPixelV hdc, PPX(iCnt0).COR.x, PPX(iCnt0).COR.y, AlphaBlend(Color, GetPixel(hdc, PPX(iCnt0).COR.x, PPX(iCnt0).COR.y), PPX(iCnt0).Alpha * Alpha)
Next iCnt0

'and that's it;)
End Sub

Public Function AlphaBlend(ByVal FirstColor As Long, ByVal SecondColor As Long, ByVal AlphaValue As Long) As Long
'This code is from:
'   Outlook Bar Project
'   Copyright (c) 2002 Vlad Vissoultchev (wqw@myrealbox.com)
'Found on PSC (www.planet-source-code.com\vb)

'All credits for the function are his...Only some variables were renamed...

Dim c1         As ColorAndAlpha
Dim c2         As ColorAndAlpha

OleTranslateColor FirstColor, 0, VarPtr(c1)
OleTranslateColor SecondColor, 0, VarPtr(c2)
If AlphaValue > 255 Then AlphaValue = 255
On Error Resume Next
With c1
    .r = (.r * AlphaValue + c2.r * (255 - AlphaValue)) / 255
    .G = (.G * AlphaValue + c2.G * (255 - AlphaValue)) / 255
    .b = (.b * AlphaValue + c2.b * (255 - AlphaValue)) / 255
End With

CopyMemory VarPtr(AlphaBlend), VarPtr(c1), 4
    
End Function
