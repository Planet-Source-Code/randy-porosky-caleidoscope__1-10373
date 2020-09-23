VERSION 5.00
Begin VB.Form Caleido 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   1095
   ClientTop       =   1485
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Caleido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Randy Porosky
Option Explicit
DefInt A-Y
' GDI functions:
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const PIXEL = 3
Const low = 6
Const Xlow = 0
Const Ylow = 0
Const Xhigh = 640
Const Yhigh = 480
Const Xlow1 = 320
Const Xhigh1 = 320

Dim Ico$
Dim IconFileName$
Dim path As String

Dim a1 As Integer
Dim a2 As Integer
Dim a3 As Integer
Dim a4 As Integer
Dim b1 As Integer
Dim b2 As Integer
Dim b3 As Integer
Dim b4 As Integer
Dim c1 As Integer
Dim c2 As Integer
Dim c3 As Integer
Dim c4 As Integer
Dim CurX As Integer
Dim CurY As Integer
Dim d1 As Integer
Dim d2 As Integer
Dim d3 As Integer
Dim d4 As Integer
Dim hi As Integer
Dim k As Integer
Dim lo As Integer
Dim m  As Integer
Dim max As Integer
Dim repeats As Integer
Dim speed As Integer
Dim v1 As Integer
Dim v1b As Integer
Dim v2 As Integer
Dim v2b As Integer
Dim v3 As Integer
Dim v3b As Integer
Dim v4 As Integer
Dim v4b As Integer
Dim var1 As Integer
Dim var2 As Integer
Dim LocX1 As Integer
Dim LocX2 As Integer
Dim LocY1 As Integer
Dim LocY2 As Integer
Dim z As Integer
Dim X%
Dim Y%
Dim hDestDC%
Dim nWidth%
Dim nHeight%
Dim hSrcDC%
Dim XSrc%
Dim YSrc%
Dim dwRop&
Dim Suc&
Dim I As Integer
Dim J As Integer
Dim s1!, s2!

Sub PresentIcon()
' Assign information of the destination bitmap.
' Note that Bitblt requires coordination in pixels.
X% = 0
Y% = 0
nWidth% = 32
nHeight% = 32
' Assign the SRCCOPY constant to the Raster operation.
dwRop& = &HCC0020

Caleido.Picture = LoadPicture(IconFileName$)
Caleido.ScaleMode = PIXEL
hDestDC% = Caleido.hDC
' Assign information of the source bitmap.
hSrcDC% = Caleido.hDC
XSrc% = 0
YSrc% = 0

For I% = 0 To 479 Step 32
For J% = 0 To 639 Step 32
Rem Call ShowIcon(IconFileName$, J%, I%)
X% = J%
Y% = I%
Suc& = BitBlt&(hDestDC%, X%, Y%, nWidth%, nHeight%, hSrcDC%, XSrc%, YSrc%, dwRop&)
Refresh
Next J%
Next I%

End Sub



Private Sub Form_Load()
Caleido.Show
IconMake
Caleido.Icon = LoadPicture(IconFileName$)
Main
End Sub


Private Sub Main()
Randomize Timer
speed = 10
CurX = 0
CurY = 0
Do
'-load random colors
lo = 1
hi = 14
'-------------
k = 4
max = 74
Line (Xlow, Ylow)-(Xhigh, Yhigh), QBColor(15), BF  'white background
'-----------------------------
'starting coordinates
a1 = Rnd * Xhigh
b1 = Rnd * Yhigh
a2 = Rnd * Xhigh
b2 = Rnd * Yhigh
'-------------
v1 = 2
v2 = 2
v3 = 2
v4 = 1
'-------------
c1 = a1
d1 = b1
c2 = a2
d2 = b2
'----------------
v1b = 2
v2b = 2
v3b = 2
v4b = 1

For repeats = 1 To 4000
DoEvents
'-----------------------
'mirror coordinates
a3 = Xhigh - a1
a4 = Xhigh - a2
b3 = Yhigh - b1
b4 = Yhigh - b2
'------------
c3 = Xhigh - c1
c4 = Xhigh - c2
d3 = Yhigh - d1
d4 = Yhigh - d2
'------------
'animated lines for each part of screen
Line (a1, b1)-(a2, b2), QBColor(k)
Line (a3, b1)-(a4, b2), QBColor(k)
Line (a1, b3)-(a2, b4), QBColor(k)
Line (a3, b3)-(a4, b4), QBColor(k)
'------------------------
a1 = a1 - v1
a2 = a2 + v2
b1 = b1 + v3
b2 = b2 - v4

'------bounce effect------
If a1 > Xhigh Then v1 = 2: k = Rnd * 15
If a1 < low Then v1 = -2: k = Rnd * 15
If a2 > Xhigh Then v2 = -2: k = Rnd * 15
If a2 < low Then v2 = 2: k = Rnd * 15
If b1 > Yhigh Then v3 = -2: k = Rnd * 15
If b1 < low Then v3 = 2: k = Rnd * 15
If b2 > Yhigh Then v4 = 2: k = Rnd * 15
If b2 < low Then v4 = -1: k = Rnd * 15

'---white lines----
If repeats > max Then
Line (c1, d1)-(c2, d2), QBColor(15)
Line (c3, d1)-(c4, d2), QBColor(15)
Line (c1, d3)-(c2, d4), QBColor(15)
Line (c3, d3)-(c4, d4), QBColor(15)
'--------------------------
c1 = c1 - v1b
c2 = c2 + v2b
d1 = d1 + v3b
d2 = d2 - v4b
'------------
If c1 > Xhigh Then v1b = 2
If c1 < low Then v1b = -2
If c2 > Xhigh Then v2b = -2
If c2 < low Then v2b = 2
If d1 > Yhigh Then v3b = -2
If d1 < low Then v3b = 2
If d2 > Yhigh Then v4b = 2
If d2 < low Then v4b = -1
End If
'----------------------
For z = 1 To 100 * speed
Next z
Next repeats
Sleep 1

'--Clear Screen in Black----
'-----------------
For m = 1 To 325
DoEvents
Line (Xlow, Ylow)-(Xhigh, Yhigh), QBColor(0), B
LocX1 = LocX1 + var1
LocX2 = LocX2 - var2
For z = 1 To 50 * speed
Next z
Next m

'---Clear Screen in White----
LocX1 = Xlow1
LocX2 = Xhigh1
var1 = 1
var2 = 1
'-----------------
For m = 1 To 325
DoEvents
Line (LocX1, Ylow)-(LocX2, Yhigh), QBColor(15), B
LocX1 = LocX1 - var1
LocX2 = LocX2 + var2
For z = 1 To 50 * speed
Next z
Next m
Loop


End Sub

Private Sub Sleep(Seconds As Double)
   Dim TempTime As Double
   TempTime = Timer
   While Timer - TempTime < Seconds
      DoEvents
      If Timer < TempTime Then
         TempTime = TempTime - 24# * 3600#
      End If
   Wend
End Sub

Sub IconMake()
Rem
Rem ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
Rem Remove this snippet when the ico is made
Rem An Icon structure:
Rem Head = 126 bytes, the last 64 bytes are color table
Rem
Rem Body = 512 bytes , a 16 (cols) * 32 (rows) Down up Color Chart
Rem                    each byte in cols are two color mixed, use div 16/and 15 for extract
Rem Tail =128 bytes all null
Rem ¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬¬
Rem
path = App.path
If Right$(path, 1) <> "\" Then path = path + "\"
IconFileName$ = path + "caleido.ico"
Ico$ = String$(766, 0)
Mid$(Ico$, 3, 1) = Chr$(1)
Mid$(Ico$, 5, 1) = Chr$(1)
Mid$(Ico$, 7, 1) = Chr$(32)
Mid$(Ico$, 8, 1) = Chr$(32)
Mid$(Ico$, 9, 1) = Chr$(16)
Mid$(Ico$, 15, 1) = Chr$(232)
Mid$(Ico$, 16, 1) = Chr$(2)
Mid$(Ico$, 19, 1) = Chr$(22)
Mid$(Ico$, 23, 1) = Chr$(40)
Mid$(Ico$, 27, 1) = Chr$(32)
Mid$(Ico$, 31, 1) = Chr$(64)
Mid$(Ico$, 35, 1) = Chr$(1)
Mid$(Ico$, 37, 1) = Chr$(4)
Mid$(Ico$, 69, 1) = Chr$(128)
Mid$(Ico$, 72, 1) = Chr$(128)
Mid$(Ico$, 76, 1) = Chr$(128)
Mid$(Ico$, 77, 1) = Chr$(128)
Mid$(Ico$, 79, 1) = Chr$(128)
Mid$(Ico$, 83, 1) = Chr$(128)
Mid$(Ico$, 85, 1) = Chr$(128)
Mid$(Ico$, 87, 1) = Chr$(128)
Mid$(Ico$, 88, 1) = Chr$(128)
Mid$(Ico$, 91, 1) = Chr$(192)
Mid$(Ico$, 92, 1) = Chr$(192)
Mid$(Ico$, 93, 1) = Chr$(192)
Mid$(Ico$, 95, 1) = Chr$(128)
Mid$(Ico$, 96, 1) = Chr$(128)
Mid$(Ico$, 97, 1) = Chr$(128)
Mid$(Ico$, 101, 1) = Chr$(255)
Mid$(Ico$, 104, 1) = Chr$(255)
Mid$(Ico$, 108, 1) = Chr$(255)
Mid$(Ico$, 109, 1) = Chr$(255)
Mid$(Ico$, 111, 1) = Chr$(255)
Mid$(Ico$, 115, 1) = Chr$(255)
Mid$(Ico$, 117, 1) = Chr$(255)
Mid$(Ico$, 119, 1) = Chr$(255)
Mid$(Ico$, 120, 1) = Chr$(255)
Mid$(Ico$, 123, 1) = Chr$(255)
Mid$(Ico$, 124, 1) = Chr$(255)
Mid$(Ico$, 125, 1) = Chr$(255)
Mid$(Ico$, 127, 1) = Chr$(79)
Mid$(Ico$, 128, 1) = Chr$(255)
Mid$(Ico$, 129, 1) = Chr$(255)
Mid$(Ico$, 130, 1) = Chr$(255)
Mid$(Ico$, 131, 1) = Chr$(79)
Mid$(Ico$, 132, 1) = Chr$(255)
Mid$(Ico$, 133, 1) = Chr$(255)
Mid$(Ico$, 134, 1) = Chr$(255)
Mid$(Ico$, 135, 1) = Chr$(79)
Mid$(Ico$, 136, 1) = Chr$(255)
Mid$(Ico$, 137, 1) = Chr$(255)
Mid$(Ico$, 138, 1) = Chr$(255)
Mid$(Ico$, 139, 1) = Chr$(79)
Mid$(Ico$, 140, 1) = Chr$(255)
Mid$(Ico$, 141, 1) = Chr$(255)
Mid$(Ico$, 142, 1) = Chr$(255)
Mid$(Ico$, 143, 1) = Chr$(244)
Mid$(Ico$, 144, 1) = Chr$(255)
Mid$(Ico$, 145, 1) = Chr$(255)
Mid$(Ico$, 146, 1) = Chr$(255)
Mid$(Ico$, 147, 1) = Chr$(244)
Mid$(Ico$, 148, 1) = Chr$(255)
Mid$(Ico$, 149, 1) = Chr$(255)
Mid$(Ico$, 150, 1) = Chr$(240)
Mid$(Ico$, 154, 1) = Chr$(15)
Mid$(Ico$, 155, 1) = Chr$(244)
Mid$(Ico$, 156, 1) = Chr$(255)
Mid$(Ico$, 157, 1) = Chr$(255)
Mid$(Ico$, 158, 1) = Chr$(255)
Mid$(Ico$, 159, 1) = Chr$(255)
Mid$(Ico$, 160, 1) = Chr$(79)
Mid$(Ico$, 161, 1) = Chr$(255)
Mid$(Ico$, 162, 1) = Chr$(255)
Mid$(Ico$, 163, 1) = Chr$(255)
Mid$(Ico$, 164, 1) = Chr$(79)
Mid$(Ico$, 165, 1) = Chr$(240)
Mid$(Ico$, 166, 1) = Chr$(1)
Mid$(Ico$, 167, 1) = Chr$(16)
Mid$(Ico$, 168, 1) = Chr$(1)
Mid$(Ico$, 169, 1) = Chr$(16)
Mid$(Ico$, 171, 1) = Chr$(15)
Mid$(Ico$, 172, 1) = Chr$(79)
Mid$(Ico$, 173, 1) = Chr$(255)
Mid$(Ico$, 174, 1) = Chr$(255)
Mid$(Ico$, 175, 1) = Chr$(255)
Mid$(Ico$, 176, 1) = Chr$(244)
Mid$(Ico$, 177, 1) = Chr$(255)
Mid$(Ico$, 180, 1) = Chr$(4)
Mid$(Ico$, 181, 1) = Chr$(1)
Mid$(Ico$, 182, 1) = Chr$(16)
Mid$(Ico$, 183, 1) = Chr$(1)
Mid$(Ico$, 184, 1) = Chr$(16)
Mid$(Ico$, 185, 1) = Chr$(17)
Mid$(Ico$, 186, 1) = Chr$(16)
Mid$(Ico$, 188, 1) = Chr$(244)
Mid$(Ico$, 189, 1) = Chr$(255)
Mid$(Ico$, 190, 1) = Chr$(255)
Mid$(Ico$, 191, 1) = Chr$(255)
Mid$(Ico$, 192, 1) = Chr$(244)
Mid$(Ico$, 194, 1) = Chr$(17)
Mid$(Ico$, 195, 1) = Chr$(16)
Mid$(Ico$, 196, 1) = Chr$(208)
Mid$(Ico$, 197, 1) = Chr$(16)
Mid$(Ico$, 198, 1) = Chr$(1)
Mid$(Ico$, 199, 1) = Chr$(17)
Mid$(Ico$, 200, 1) = Chr$(17)
Mid$(Ico$, 201, 1) = Chr$(17)
Mid$(Ico$, 202, 1) = Chr$(4)
Mid$(Ico$, 203, 1) = Chr$(16)
Mid$(Ico$, 204, 1) = Chr$(244)
Mid$(Ico$, 205, 1) = Chr$(79)
Mid$(Ico$, 206, 1) = Chr$(255)
Mid$(Ico$, 207, 1) = Chr$(255)
Mid$(Ico$, 208, 1) = Chr$(64)
Mid$(Ico$, 209, 1) = Chr$(17)
Mid$(Ico$, 210, 1) = Chr$(16)
Mid$(Ico$, 211, 1) = Chr$(17)
Mid$(Ico$, 212, 1) = Chr$(16)
Mid$(Ico$, 213, 1) = Chr$(17)
Mid$(Ico$, 214, 1) = Chr$(17)
Mid$(Ico$, 215, 1) = Chr$(1)
Mid$(Ico$, 216, 1) = Chr$(17)
Mid$(Ico$, 217, 1) = Chr$(17)
Mid$(Ico$, 218, 1) = Chr$(16)
Mid$(Ico$, 219, 1) = Chr$(17)
Mid$(Ico$, 220, 1) = Chr$(15)
Mid$(Ico$, 221, 1) = Chr$(244)
Mid$(Ico$, 222, 1) = Chr$(255)
Mid$(Ico$, 223, 1) = Chr$(244)
Mid$(Ico$, 224, 1) = Chr$(1)
Mid$(Ico$, 225, 1) = Chr$(16)
Mid$(Ico$, 226, 1) = Chr$(1)
Mid$(Ico$, 227, 1) = Chr$(209)
Mid$(Ico$, 228, 1) = Chr$(208)
Mid$(Ico$, 229, 1) = Chr$(16)
Mid$(Ico$, 230, 1) = Chr$(1)
Mid$(Ico$, 231, 1) = Chr$(17)
Mid$(Ico$, 232, 1) = Chr$(17)
Mid$(Ico$, 233, 1) = Chr$(17)
Mid$(Ico$, 234, 1) = Chr$(17)
Mid$(Ico$, 235, 1) = Chr$(17)
Mid$(Ico$, 236, 1) = Chr$(15)
Mid$(Ico$, 237, 1) = Chr$(255)
Mid$(Ico$, 238, 1) = Chr$(79)
Mid$(Ico$, 239, 1) = Chr$(64)
Mid$(Ico$, 240, 1) = Chr$(1)
Mid$(Ico$, 241, 1) = Chr$(1)
Mid$(Ico$, 242, 1) = Chr$(29)
Mid$(Ico$, 243, 1) = Chr$(29)
Mid$(Ico$, 244, 1) = Chr$(16)
Mid$(Ico$, 245, 1) = Chr$(1)
Mid$(Ico$, 246, 1) = Chr$(17)
Mid$(Ico$, 247, 1) = Chr$(17)
Mid$(Ico$, 248, 1) = Chr$(17)
Mid$(Ico$, 249, 1) = Chr$(17)
Mid$(Ico$, 250, 1) = Chr$(17)
Mid$(Ico$, 251, 1) = Chr$(17)
Mid$(Ico$, 252, 1) = Chr$(15)
Mid$(Ico$, 253, 1) = Chr$(255)
Mid$(Ico$, 254, 1) = Chr$(244)
Mid$(Ico$, 255, 1) = Chr$(64)
Mid$(Ico$, 256, 1) = Chr$(17)
Mid$(Ico$, 257, 1) = Chr$(16)
Mid$(Ico$, 258, 1) = Chr$(209)
Mid$(Ico$, 259, 1) = Chr$(209)
Mid$(Ico$, 260, 1) = Chr$(208)
Mid$(Ico$, 261, 1) = Chr$(17)
Mid$(Ico$, 262, 1) = Chr$(17)
Mid$(Ico$, 263, 1) = Chr$(17)
Mid$(Ico$, 264, 1) = Chr$(17)
Mid$(Ico$, 265, 1) = Chr$(29)
Mid$(Ico$, 266, 1) = Chr$(209)
Mid$(Ico$, 267, 1) = Chr$(17)
Mid$(Ico$, 268, 1) = Chr$(15)
Mid$(Ico$, 269, 1) = Chr$(255)
Mid$(Ico$, 270, 1) = Chr$(255)
Mid$(Ico$, 271, 1) = Chr$(240)
Mid$(Ico$, 272, 1) = Chr$(16)
Mid$(Ico$, 273, 1) = Chr$(17)
Mid$(Ico$, 274, 1) = Chr$(17)
Mid$(Ico$, 275, 1) = Chr$(29)
Mid$(Ico$, 276, 1) = Chr$(16)
Mid$(Ico$, 277, 1) = Chr$(1)
Mid$(Ico$, 278, 1) = Chr$(17)
Mid$(Ico$, 279, 1) = Chr$(17)
Mid$(Ico$, 280, 1) = Chr$(17)
Mid$(Ico$, 281, 1) = Chr$(223)
Mid$(Ico$, 282, 1) = Chr$(209)
Mid$(Ico$, 283, 1) = Chr$(17)
Mid$(Ico$, 284, 1) = Chr$(15)
Mid$(Ico$, 285, 1) = Chr$(255)
Mid$(Ico$, 286, 1) = Chr$(255)
Mid$(Ico$, 287, 1) = Chr$(240)
Mid$(Ico$, 288, 1) = Chr$(17)
Mid$(Ico$, 289, 1) = Chr$(17)
Mid$(Ico$, 290, 1) = Chr$(17)
Mid$(Ico$, 291, 1) = Chr$(17)
Mid$(Ico$, 292, 1) = Chr$(208)
Mid$(Ico$, 293, 1) = Chr$(17)
Mid$(Ico$, 294, 1) = Chr$(17)
Mid$(Ico$, 295, 1) = Chr$(17)
Mid$(Ico$, 296, 1) = Chr$(29)
Mid$(Ico$, 297, 1) = Chr$(255)
Mid$(Ico$, 298, 1) = Chr$(209)
Mid$(Ico$, 299, 1) = Chr$(17)
Mid$(Ico$, 300, 1) = Chr$(15)
Mid$(Ico$, 301, 1) = Chr$(255)
Mid$(Ico$, 302, 1) = Chr$(255)
Mid$(Ico$, 303, 1) = Chr$(240)
Mid$(Ico$, 304, 1) = Chr$(16)
Mid$(Ico$, 305, 1) = Chr$(1)
Mid$(Ico$, 306, 1) = Chr$(17)
Mid$(Ico$, 307, 1) = Chr$(29)
Mid$(Ico$, 308, 1) = Chr$(16)
Mid$(Ico$, 309, 1) = Chr$(17)
Mid$(Ico$, 310, 1) = Chr$(17)
Mid$(Ico$, 311, 1) = Chr$(17)
Mid$(Ico$, 312, 1) = Chr$(212)
Mid$(Ico$, 313, 1) = Chr$(221)
Mid$(Ico$, 314, 1) = Chr$(17)
Mid$(Ico$, 315, 1) = Chr$(17)
Mid$(Ico$, 316, 1) = Chr$(4)
Mid$(Ico$, 317, 1) = Chr$(255)
Mid$(Ico$, 318, 1) = Chr$(255)
Mid$(Ico$, 319, 1) = Chr$(240)
Mid$(Ico$, 320, 1) = Chr$(1)
Mid$(Ico$, 321, 1) = Chr$(17)
Mid$(Ico$, 322, 1) = Chr$(17)
Mid$(Ico$, 323, 1) = Chr$(17)
Mid$(Ico$, 324, 1) = Chr$(209)
Mid$(Ico$, 325, 1) = Chr$(1)
Mid$(Ico$, 326, 1) = Chr$(17)
Mid$(Ico$, 327, 1) = Chr$(17)
Mid$(Ico$, 328, 1) = Chr$(29)
Mid$(Ico$, 329, 1) = Chr$(17)
Mid$(Ico$, 330, 1) = Chr$(17)
Mid$(Ico$, 331, 1) = Chr$(17)
Mid$(Ico$, 333, 1) = Chr$(79)
Mid$(Ico$, 334, 1) = Chr$(255)
Mid$(Ico$, 335, 1) = Chr$(240)
Mid$(Ico$, 336, 1) = Chr$(16)
Mid$(Ico$, 337, 1) = Chr$(17)
Mid$(Ico$, 338, 1) = Chr$(17)
Mid$(Ico$, 339, 1) = Chr$(29)
Mid$(Ico$, 340, 1) = Chr$(221)
Mid$(Ico$, 341, 1) = Chr$(13)
Mid$(Ico$, 342, 1) = Chr$(17)
Mid$(Ico$, 343, 1) = Chr$(17)
Mid$(Ico$, 344, 1) = Chr$(17)
Mid$(Ico$, 345, 1) = Chr$(17)
Mid$(Ico$, 346, 1) = Chr$(29)
Mid$(Ico$, 347, 1) = Chr$(16)
Mid$(Ico$, 348, 1) = Chr$(16)
Mid$(Ico$, 349, 1) = Chr$(4)
Mid$(Ico$, 350, 1) = Chr$(255)
Mid$(Ico$, 351, 1) = Chr$(244)
Mid$(Ico$, 352, 1) = Chr$(1)
Mid$(Ico$, 353, 1) = Chr$(17)
Mid$(Ico$, 354, 1) = Chr$(17)
Mid$(Ico$, 355, 1) = Chr$(212)
Mid$(Ico$, 356, 1) = Chr$(253)
Mid$(Ico$, 357, 1) = Chr$(16)
Mid$(Ico$, 358, 1) = Chr$(1)
Mid$(Ico$, 359, 1) = Chr$(209)
Mid$(Ico$, 360, 1) = Chr$(209)
Mid$(Ico$, 361, 1) = Chr$(209)
Mid$(Ico$, 362, 1) = Chr$(209)
Mid$(Ico$, 364, 1) = Chr$(1)
Mid$(Ico$, 365, 1) = Chr$(16)
Mid$(Ico$, 366, 1) = Chr$(79)
Mid$(Ico$, 367, 1) = Chr$(79)
Mid$(Ico$, 368, 1) = Chr$(240)
Mid$(Ico$, 369, 1) = Chr$(17)
Mid$(Ico$, 370, 1) = Chr$(17)
Mid$(Ico$, 371, 1) = Chr$(29)
Mid$(Ico$, 372, 1) = Chr$(209)
Mid$(Ico$, 373, 1) = Chr$(16)
Mid$(Ico$, 374, 1) = Chr$(240)
Mid$(Ico$, 375, 1) = Chr$(13)
Mid$(Ico$, 376, 1) = Chr$(29)
Mid$(Ico$, 377, 1) = Chr$(16)
Mid$(Ico$, 379, 1) = Chr$(209)
Mid$(Ico$, 380, 1) = Chr$(16)
Mid$(Ico$, 381, 1) = Chr$(1)
Mid$(Ico$, 382, 1) = Chr$(4)
Mid$(Ico$, 383, 1) = Chr$(79)
Mid$(Ico$, 384, 1) = Chr$(255)
Mid$(Ico$, 386, 1) = Chr$(17)
Mid$(Ico$, 387, 1) = Chr$(17)
Mid$(Ico$, 388, 1) = Chr$(16)
Mid$(Ico$, 389, 1) = Chr$(15)
Mid$(Ico$, 390, 1) = Chr$(255)
Mid$(Ico$, 391, 1) = Chr$(64)
Mid$(Ico$, 394, 1) = Chr$(17)
Mid$(Ico$, 395, 1) = Chr$(16)
Mid$(Ico$, 396, 1) = Chr$(29)
Mid$(Ico$, 397, 1) = Chr$(29)
Mid$(Ico$, 398, 1) = Chr$(15)
Mid$(Ico$, 399, 1) = Chr$(244)
Mid$(Ico$, 400, 1) = Chr$(15)
Mid$(Ico$, 401, 1) = Chr$(255)
Mid$(Ico$, 405, 1) = Chr$(255)
Mid$(Ico$, 406, 1) = Chr$(255)
Mid$(Ico$, 407, 1) = Chr$(240)
Mid$(Ico$, 408, 1) = Chr$(32)
Mid$(Ico$, 409, 1) = Chr$(209)
Mid$(Ico$, 410, 1) = Chr$(209)
Mid$(Ico$, 411, 1) = Chr$(209)
Mid$(Ico$, 412, 1) = Chr$(209)
Mid$(Ico$, 413, 1) = Chr$(209)
Mid$(Ico$, 414, 1) = Chr$(208)
Mid$(Ico$, 415, 1) = Chr$(255)
Mid$(Ico$, 417, 1) = Chr$(255)
Mid$(Ico$, 418, 1) = Chr$(255)
Mid$(Ico$, 419, 1) = Chr$(255)
Mid$(Ico$, 420, 1) = Chr$(2)
Mid$(Ico$, 421, 1) = Chr$(255)
Mid$(Ico$, 422, 1) = Chr$(255)
Mid$(Ico$, 424, 1) = Chr$(1)
Mid$(Ico$, 425, 1) = Chr$(29)
Mid$(Ico$, 426, 1) = Chr$(17)
Mid$(Ico$, 427, 1) = Chr$(17)
Mid$(Ico$, 428, 1) = Chr$(17)
Mid$(Ico$, 429, 1) = Chr$(17)
Mid$(Ico$, 430, 1) = Chr$(16)
Mid$(Ico$, 431, 1) = Chr$(255)
Mid$(Ico$, 432, 1) = Chr$(2)
Mid$(Ico$, 434, 1) = Chr$(255)
Mid$(Ico$, 435, 1) = Chr$(255)
Mid$(Ico$, 437, 1) = Chr$(255)
Mid$(Ico$, 438, 1) = Chr$(240)
Mid$(Ico$, 440, 1) = Chr$(16)
Mid$(Ico$, 441, 1) = Chr$(1)
Mid$(Ico$, 442, 1) = Chr$(17)
Mid$(Ico$, 443, 1) = Chr$(17)
Mid$(Ico$, 444, 1) = Chr$(29)
Mid$(Ico$, 445, 1) = Chr$(17)
Mid$(Ico$, 446, 1) = Chr$(16)
Mid$(Ico$, 447, 1) = Chr$(255)
Mid$(Ico$, 448, 1) = Chr$(2)
Mid$(Ico$, 449, 1) = Chr$(34)
Mid$(Ico$, 450, 1) = Chr$(15)
Mid$(Ico$, 451, 1) = Chr$(255)
Mid$(Ico$, 452, 1) = Chr$(2)
Mid$(Ico$, 453, 1) = Chr$(79)
Mid$(Ico$, 455, 1) = Chr$(47)
Mid$(Ico$, 457, 1) = Chr$(17)
Mid$(Ico$, 458, 1) = Chr$(17)
Mid$(Ico$, 459, 1) = Chr$(17)
Mid$(Ico$, 460, 1) = Chr$(212)
Mid$(Ico$, 461, 1) = Chr$(209)
Mid$(Ico$, 462, 1) = Chr$(16)
Mid$(Ico$, 463, 1) = Chr$(255)
Mid$(Ico$, 464, 1) = Chr$(2)
Mid$(Ico$, 465, 1) = Chr$(34)
Mid$(Ico$, 466, 1) = Chr$(15)
Mid$(Ico$, 467, 1) = Chr$(255)
Mid$(Ico$, 469, 1) = Chr$(244)
Mid$(Ico$, 471, 1) = Chr$(255)
Mid$(Ico$, 473, 1) = Chr$(17)
Mid$(Ico$, 474, 1) = Chr$(17)
Mid$(Ico$, 475, 1) = Chr$(221)
Mid$(Ico$, 476, 1) = Chr$(221)
Mid$(Ico$, 477, 1) = Chr$(17)
Mid$(Ico$, 478, 1) = Chr$(16)
Mid$(Ico$, 479, 1) = Chr$(244)
Mid$(Ico$, 480, 1) = Chr$(2)
Mid$(Ico$, 481, 1) = Chr$(34)
Mid$(Ico$, 482, 1) = Chr$(15)
Mid$(Ico$, 483, 1) = Chr$(244)
Mid$(Ico$, 484, 1) = Chr$(2)
Mid$(Ico$, 485, 1) = Chr$(240)
Mid$(Ico$, 486, 1) = Chr$(15)
Mid$(Ico$, 487, 1) = Chr$(240)
Mid$(Ico$, 488, 1) = Chr$(2)
Mid$(Ico$, 489, 1) = Chr$(17)
Mid$(Ico$, 490, 1) = Chr$(29)
Mid$(Ico$, 491, 1) = Chr$(244)
Mid$(Ico$, 492, 1) = Chr$(253)
Mid$(Ico$, 493, 1) = Chr$(17)
Mid$(Ico$, 494, 1) = Chr$(15)
Mid$(Ico$, 495, 1) = Chr$(79)
Mid$(Ico$, 497, 1) = Chr$(2)
Mid$(Ico$, 498, 1) = Chr$(32)
Mid$(Ico$, 499, 1) = Chr$(79)
Mid$(Ico$, 501, 1) = Chr$(242)
Mid$(Ico$, 502, 1) = Chr$(36)
Mid$(Ico$, 503, 1) = Chr$(64)
Mid$(Ico$, 505, 1) = Chr$(17)
Mid$(Ico$, 506, 1) = Chr$(17)
Mid$(Ico$, 507, 1) = Chr$(221)
Mid$(Ico$, 508, 1) = Chr$(209)
Mid$(Ico$, 509, 1) = Chr$(16)
Mid$(Ico$, 510, 1) = Chr$(244)
Mid$(Ico$, 511, 1) = Chr$(79)
Mid$(Ico$, 512, 1) = Chr$(2)
Mid$(Ico$, 513, 1) = Chr$(34)
Mid$(Ico$, 514, 1) = Chr$(32)
Mid$(Ico$, 515, 1) = Chr$(79)
Mid$(Ico$, 516, 1) = Chr$(2)
Mid$(Ico$, 518, 1) = Chr$(255)
Mid$(Ico$, 519, 1) = Chr$(32)
Mid$(Ico$, 520, 1) = Chr$(255)
Mid$(Ico$, 521, 1) = Chr$(1)
Mid$(Ico$, 522, 1) = Chr$(17)
Mid$(Ico$, 523, 1) = Chr$(17)
Mid$(Ico$, 524, 1) = Chr$(16)
Mid$(Ico$, 525, 1) = Chr$(15)
Mid$(Ico$, 526, 1) = Chr$(255)
Mid$(Ico$, 527, 1) = Chr$(244)
Mid$(Ico$, 528, 1) = Chr$(240)
Mid$(Ico$, 529, 1) = Chr$(2)
Mid$(Ico$, 530, 1) = Chr$(32)
Mid$(Ico$, 531, 1) = Chr$(4)
Mid$(Ico$, 533, 1) = Chr$(7)
Mid$(Ico$, 535, 1) = Chr$(4)
Mid$(Ico$, 536, 1) = Chr$(255)
Mid$(Ico$, 537, 1) = Chr$(240)
Mid$(Ico$, 540, 1) = Chr$(15)
Mid$(Ico$, 541, 1) = Chr$(255)
Mid$(Ico$, 542, 1) = Chr$(255)
Mid$(Ico$, 543, 1) = Chr$(255)
Mid$(Ico$, 544, 1) = Chr$(64)
Mid$(Ico$, 545, 1) = Chr$(34)
Mid$(Ico$, 546, 1) = Chr$(2)
Mid$(Ico$, 547, 1) = Chr$(15)
Mid$(Ico$, 548, 1) = Chr$(4)
Mid$(Ico$, 551, 1) = Chr$(255)
Mid$(Ico$, 552, 1) = Chr$(79)
Mid$(Ico$, 553, 1) = Chr$(255)
Mid$(Ico$, 554, 1) = Chr$(255)
Mid$(Ico$, 555, 1) = Chr$(255)
Mid$(Ico$, 556, 1) = Chr$(79)
Mid$(Ico$, 557, 1) = Chr$(255)
Mid$(Ico$, 558, 1) = Chr$(255)
Mid$(Ico$, 559, 1) = Chr$(255)
Mid$(Ico$, 560, 1) = Chr$(244)
Mid$(Ico$, 561, 1) = Chr$(2)
Mid$(Ico$, 562, 1) = Chr$(32)
Mid$(Ico$, 566, 1) = Chr$(2)
Mid$(Ico$, 568, 1) = Chr$(4)
Mid$(Ico$, 569, 1) = Chr$(255)
Mid$(Ico$, 570, 1) = Chr$(255)
Mid$(Ico$, 571, 1) = Chr$(255)
Mid$(Ico$, 572, 1) = Chr$(244)
Mid$(Ico$, 573, 1) = Chr$(255)
Mid$(Ico$, 574, 1) = Chr$(255)
Mid$(Ico$, 575, 1) = Chr$(255)
Mid$(Ico$, 576, 1) = Chr$(244)
Mid$(Ico$, 578, 1) = Chr$(2)
Mid$(Ico$, 579, 1) = Chr$(32)
Mid$(Ico$, 580, 1) = Chr$(4)
Mid$(Ico$, 582, 1) = Chr$(34)
Mid$(Ico$, 583, 1) = Chr$(32)
Mid$(Ico$, 584, 1) = Chr$(32)
Mid$(Ico$, 585, 1) = Chr$(79)
Mid$(Ico$, 586, 1) = Chr$(255)
Mid$(Ico$, 587, 1) = Chr$(255)
Mid$(Ico$, 588, 1) = Chr$(244)
Mid$(Ico$, 589, 1) = Chr$(79)
Mid$(Ico$, 590, 1) = Chr$(255)
Mid$(Ico$, 591, 1) = Chr$(255)
Mid$(Ico$, 592, 1) = Chr$(79)
Mid$(Ico$, 593, 1) = Chr$(240)
Mid$(Ico$, 595, 1) = Chr$(32)
Mid$(Ico$, 596, 1) = Chr$(64)
Mid$(Ico$, 597, 1) = Chr$(34)
Mid$(Ico$, 598, 1) = Chr$(34)
Mid$(Ico$, 599, 1) = Chr$(2)
Mid$(Ico$, 600, 1) = Chr$(32)
Mid$(Ico$, 601, 1) = Chr$(244)
Mid$(Ico$, 602, 1) = Chr$(255)
Mid$(Ico$, 603, 1) = Chr$(255)
Mid$(Ico$, 604, 1) = Chr$(79)
Mid$(Ico$, 605, 1) = Chr$(244)
Mid$(Ico$, 606, 1) = Chr$(255)
Mid$(Ico$, 607, 1) = Chr$(244)
Mid$(Ico$, 608, 1) = Chr$(255)
Mid$(Ico$, 609, 1) = Chr$(255)
Mid$(Ico$, 610, 1) = Chr$(64)
Mid$(Ico$, 612, 1) = Chr$(240)
Mid$(Ico$, 616, 1) = Chr$(34)
Mid$(Ico$, 617, 1) = Chr$(15)
Mid$(Ico$, 618, 1) = Chr$(79)
Mid$(Ico$, 619, 1) = Chr$(244)
Mid$(Ico$, 620, 1) = Chr$(255)
Mid$(Ico$, 621, 1) = Chr$(255)
Mid$(Ico$, 622, 1) = Chr$(79)
Mid$(Ico$, 623, 1) = Chr$(79)
Mid$(Ico$, 624, 1) = Chr$(255)
Mid$(Ico$, 625, 1) = Chr$(255)
Mid$(Ico$, 626, 1) = Chr$(244)
Mid$(Ico$, 627, 1) = Chr$(79)
Mid$(Ico$, 628, 1) = Chr$(255)
Mid$(Ico$, 629, 1) = Chr$(255)
Mid$(Ico$, 630, 1) = Chr$(244)
Mid$(Ico$, 631, 1) = Chr$(64)
Mid$(Ico$, 634, 1) = Chr$(244)
Mid$(Ico$, 635, 1) = Chr$(79)
Mid$(Ico$, 636, 1) = Chr$(255)
Mid$(Ico$, 637, 1) = Chr$(255)
Mid$(Ico$, 638, 1) = Chr$(244)

Open IconFileName$ For Binary As #1
Put #1, , Ico$
Reset

Ico$ = ""

End Sub

Private Sub Form_MouseOver(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub
