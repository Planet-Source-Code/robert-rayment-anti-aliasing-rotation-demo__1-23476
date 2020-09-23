VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   563
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Anti-aliasing Rotate"
      Height          =   315
      Left            =   870
      TabIndex        =   4
      Top             =   4320
      Width           =   1545
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Rotate && Smooth"
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rotate"
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   3270
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   1
      Top             =   420
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   495
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   163
      TabIndex        =   0
      Top             =   420
      Width           =   2445
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Picture2"
      Height          =   270
      Left            =   4590
      TabIndex        =   6
      Top             =   90
      Width           =   2325
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Picture1 long color"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   45
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Anti-aliasing Demo by Robert Rayment

Option Base 1

DefLng A-W
DefSng X-Z

'Windows API - much faster then VB's PSet and Point
Private Declare Function SetPixelV Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Dim LongCul()
Dim redc As Byte, greenc As Byte, bluec As Byte

Const pi# = 3.1415926

Private Sub Form_Load()
Caption = "Rotation & Anti-aliasing  by  Robert Rayment"

Picture1.Picture = LoadPicture("Test.bmp")
ReDim LongCul(4)

End Sub


Private Sub Command1_Click()

'Rotate

Picture2.Cls

zang = pi# / 6

zCos = Cos(zang)
zSin = Sin(zang)

picW = Picture2.Width
picH = Picture2.Height

ixc1 = Picture1.Width / 2: iyc1 = Picture1.Height / 2
ixc2 = Picture2.Width / 2: iyc2 = Picture2.Height / 2

For iy = 1 To picH
For ix = 1 To picW


'   'For each Picture2 point find rotated Picture1 source point
   ixs = ixc1 + (ix - ixc2) * zCos + (iy - iyc2) * zSin
   iys = iyc1 + (iy - iyc2) * zCos - (ix - ixc2) * zSin

'   'Move valid rotated Picture1 source points to Picture2
   LongCulAV = GetPixel(Picture1.hdc, ixs, iys)
   If LongCulAV <> -1 Then
      res& = SetPixelV(Picture2.hdc, ix, iy, LongCulAV)
   End If

Next ix
Next iy


End Sub

Private Sub Command2_Click()

'ReDim LongCul(4)

' Rotate & Smooth

Picture2.Cls

zang = pi# / 6

zCos = Cos(zang)
zSin = Sin(zang)

picW = Picture2.Width
picH = Picture2.Height

ixc1 = Picture1.Width / 2: iyc1 = Picture1.Height / 2
ixc2 = Picture2.Width / 2: iyc2 = Picture2.Height / 2

For iy = 1 To picH
For ix = 1 To picW


   ixs = ixc1 + (ix - ixc2 - 1) * zCos + (iy - iyc2) * zSin
   iys = iyc1 + (iy - iyc2) * zCos - (ix - ixc2 - 1) * zSin
   LongCul(1) = GetPixel(Picture1.hdc, ixs, iys)
   If LongCul(1) = -1 Then GoTo Nextix
   
   ixs = ixc1 + (ix - ixc2 + 1) * zCos + (iy - iyc2) * zSin
   iys = iyc1 + (iy - iyc2) * zCos - (ix - ixc2 + 1) * zSin
   LongCul(2) = GetPixel(Picture1.hdc, ixs, iys)
   If LongCul(2) = -1 Then GoTo Nextix
   
   ixs = ixc1 + (ix - ixc2) * zCos + (iy - iyc2 + 1) * zSin
   iys = iyc1 + (iy - iyc2 + 1) * zCos - (ix - ixc2) * zSin
   LongCul(3) = GetPixel(Picture1.hdc, ixs, iys)
   If LongCul(3) = -1 Then GoTo Nextix

   ixs = ixc1 + (ix - ixc2) * zCos + (iy - iyc2 - 1) * zSin
   iys = iyc1 + (iy - iyc2 - 1) * zCos - (ix - ixc2) * zSin
   LongCul(4) = GetPixel(Picture1.hdc, ixs, iys)
   If LongCul(4) = -1 Then GoTo Nextix
   
   SMOOTH LongCulAV
   
   res& = SetPixelV(Picture2.hdc, ix, iy, LongCulAV)


Nextix:
Next ix
Next iy


End Sub

Private Sub Command3_Click()

' AA Rotate  Anti-aliasing

' 4 point weighting

' Let xs, ys be the point in Picture1 that would go to Picture2 on
' rotation.  xs & ys will not often be integers and in general will
' lie inside a rectangle on Picture1 from Int(xs), Int(ys) to
' Int(xs)+1, Int(ys)+1.  After checking that the point lies inside
' Picture1 a proportionate weighting of the colors at Int(xs) & Int(xs)+1
' at Int(ys) and Int(xs) & Int(xs)+1 at Int(ys)+1 gives 2 colors
' which are then proportionately weighted in the y direction.


Picture2.Cls

zang = pi# / 6 ' Rotation angle

zCos = Cos(zang)
zSin = Sin(zang)

PicW1 = Picture1.Width
PicH1 = Picture1.Height
picW2 = Picture2.Width
picH2 = Picture2.Height

ixc1 = Picture1.Width / 2: iyc1 = Picture1.Height / 2
ixc2 = Picture2.Width / 2: iyc2 = Picture2.Height / 2

For iy = 0 To picH2 - 1
For ix = 0 To picW2 - 1

   xs = ixc1 + (ix - ixc2) * zCos + (iy - iyc2) * zSin
   ys = iyc1 + (iy - iyc2) * zCos - (ix - ixc2) * zSin

   ' Bottom left coords of bounding
   ' rectangle on Picture1
   ixs0 = Int(xs)
   iys0 = Int(ys)
   
   If ixs0 > 0 And ixs0 < PicW1 - 1 And iys0 > 0 And iys0 < PicH1 - 1 Then
      
      xfs1 = xs - Int(xs)
      yfs1 = ys - Int(ys)
      
      'ixs0->ixs0+1, iyso
      LongCul0 = GetPixel(Picture1.hdc, ixs0, iys0)
      CulToRGB LongCul0, redc, greenc, bluec
      culr = (1 - xfs1) * redc
      culg = (1 - xfs1) * greenc
      culb = (1 - xfs1) * bluec
   
      LongCul1 = GetPixel(Picture1.hdc, ixs0 + 1, iys0)
      CulToRGB LongCul1, redc, greenc, bluec
      culr0 = culr + xfs1 * redc
      culg0 = culg + xfs1 * greenc
      culb0 = culb + xfs1 * bluec
      
      'ixs0->ixs0+1, iys0+1
      LongCul0 = GetPixel(Picture1.hdc, ixs0, iys0 + 1)
      CulToRGB LongCul0, redc, greenc, bluec
      culr = (1 - xfs1) * redc
      culg = (1 - xfs1) * greenc
      culb = (1 - xfs1) * bluec
      
      LongCul1 = GetPixel(Picture1.hdc, ixs0 + 1, iys0 + 1)
      CulToRGB LongCul1, redc, greenc, bluec
      culr1 = culr + xfs1 * redc
      culg1 = culg + xfs1 * greenc
      culb1 = culb + xfs1 * bluec
   
      ' Weight along y axis
      culr = (1 - yfs1) * culr0 + yfs1 * culr1
      culg = (1 - yfs1) * culg0 + yfs1 * culg1
      culb = (1 - yfs1) * culb0 + yfs1 * culb1
   
      res& = SetPixelV(Picture2.hdc, ix, iy, RGB(culr, culg, culb))

   End If
   
Next ix
Next iy

End Sub


Public Sub SMOOTH(LongCulAV)

'ReDim LongCul(4)

   D = 4
   CulToRGB LongCul(1), redc, greenc, bluec
   culr = redc / D
   culg = greenc / D
   culb = bluec / D
   
   CulToRGB LongCul(2), redc, greenc, bluec
   culr = culr + redc / D
   culg = culg + greenc / D
   culb = culb + bluec / D
   
   CulToRGB LongCul(3), redc, greenc, bluec
   culr = culr + redc / D
   culg = culg + greenc / D
   culb = culb + bluec / D
   
   CulToRGB LongCul(4), redc, greenc, bluec
   culr = culr + redc / D
   culg = culg + greenc / D
   culb = culb + bluec / D
   
   LongCulAV = RGB(culr, culg, culb)

End Sub
Public Sub CulToRGB(LongCul, redc As Byte, greenc As Byte, bluec As Byte)
'Global redc As Byte, greenc As Byte, bluec As Byte
'Input longcul:  Output: R G B
redc = LongCul And &HFF&
greenc = (LongCul& And &HFF00&) / &H100&
bluec = (LongCul& And &HFF0000) / &H10000
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = Str$(GetPixel(Picture1.hdc, X, Y))
End Sub
