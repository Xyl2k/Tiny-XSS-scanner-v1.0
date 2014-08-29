Attribute VB_Name = "basTransforms"

'Module by Mr.ambrosino

Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam As Integer, ByVal lParam&) As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Sub GenerateTransForm(ByVal Frm As Form, ByVal Pct As PictureBox, ByVal ColorValue As Long)

Dim lngX As Long, lngY As Long
Dim lngStartX As Long, lngStartY As Long
Dim lngEndX As Long, lngEndY As Long
Dim lngHRectregion As Long, lngTempHRectRegion As Long
Dim lngVoidReturn As Long

Dim blnStatus As Boolean
    

    DoEvents
    blnStatus = False
    For lngX = 0 To Pct.ScaleWidth
        blnStatus = False
        For lngY = 0 To Pct.ScaleHeight
            If blnStatus Then
                If Pct.Point(lngX, lngY) = ColorValue Then
                    lngEndX = lngX
                    lngEndY = lngY
                    If lngHRectregion = 0 Then
                        lngHRectregion = CreateRectRgn(lngStartX, lngStartY, lngEndX + 1, lngEndY)
                    Else
                        lngTempHRectRegion = CreateRectRgn(lngStartX, lngStartY, lngEndX + 1, lngEndY)
                        lngVoidReturn = CombineRgn(lngHRectregion, lngHRectregion, lngTempHRectRegion, RGN_OR)
                        DeleteObject lngTempHRectRegion
                    End If
                    blnStatus = False
                End If
             Else
                If Pct.Point(lngX, lngY) <> ColorValue Then
                    lngStartX = lngX
                    lngStartY = lngY
                    lngEndX = lngX
                    lngEndY = lngY
                    blnStatus = True
                End If
            End If
        Next

        If blnStatus Then
          lngEndX = lngX
          lngEndY = lngY
          If lngHRectregion = 0 Then
            lngHRectregion = CreateRectRgn(lngStartX, lngStartY, lngEndX + 1, lngEndY)
          Else
            lngTempHRectRegion = CreateRectRgn(lngStartX, lngStartY, lngEndX + 1, lngEndY)
            lngVoidReturn = CombineRgn(lngHRectregion, lngHRectregion, lngTempHRectRegion, RGN_OR)
            DeleteObject lngTempHRectRegion
          End If
        End If
    Next
    lngVoidReturn = SetWindowRgn(Frm.hWnd, lngHRectregion, True)
    lngVoidReturn = DeleteObject(lngHRectregion)
End Sub



