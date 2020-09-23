Attribute VB_Name = "Module1"
'****************************
'By     Jim Jose
'email  jimjosev33@yahoo.com
'****************************

'PLEASE READ THIS
'It is made to get useful for anyone without 'much' changes.
'If you feel Satisfactory
'   Please 'Rate' this code
'Else
'   Give feedback to improve this code
'End If
'Good luck
'****************************
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Sub ShadePicture(PicSource As PictureBox, PicTarget As PictureBox, WithColor As Long, Thickness As Integer)
On Error Resume Next
Dim sRate, Col As Long
Dim X, Y As Single
Dim XMax, YMax As Single
Dim cBlue, cGreen, cRed As Double   'Determines the pixel color
Dim sBlue, sGreen, sRed As Double   'Determines the SHADING color
    'Getting the RGB values of selected color
    sBlue = Fix((WithColor / 256) / 256)
    sGreen = Fix((WithColor - ((sBlue * 256) * 256)) / 256)
    sRed = Fix(WithColor - ((sBlue * 256) * 256) - (sGreen * 256))
    'Calculate screen height & width of the image
    XMax = PicSource.Width / Screen.TwipsPerPixelX - 1
    YMax = PicSource.Height / Screen.TwipsPerPixelY - 1
    'Initialising Shading
    PicTarget.Cls
    sRate = Thickness / 10
    'Process all pixels and alter them accordingly
    For X = 0 To XMax
      For Y = 0 To YMax
        Col = GetPixel(PicSource.hdc, X, Y)
        If Not Col = 0 Then     'Because black colors are usually the borders of an image and never change border color.It will affect the clarity.
            'Getting the RGB values of current pixel
            cBlue = Fix((Col / 256) / 256)
            cGreen = Fix((Col - ((cBlue * 256) * 256)) / 256)
            cRed = Fix(Col - ((cBlue * 256) * 256) - (cGreen * 256))
            'Resetting the RGB values of current pixel with  the  sRate of  shading
            cRed = cRed + (sRed - cRed) * sRate
            cGreen = cGreen + (sGreen - cGreen) * sRate
            cBlue = cBlue + (sBlue - cBlue) * sRate
            If Not Col = 12632256 Then SetPixel PicTarget.hdc, X, Y, RGB(cRed, cGreen, cBlue)   'Skipping transparent col and setting the pixel
        Else
            SetPixel PicTarget.hdc, X, Y, Col
        End If
      Next Y
    PicTarget.Refresh
Next X
End Sub

