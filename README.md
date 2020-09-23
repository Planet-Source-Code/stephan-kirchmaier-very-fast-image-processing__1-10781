<div align="center">

## Very fast Image Processing


</div>

### Description

This code shows how to realize fast Image Processing. I've made a small animation with it.
 
### More Info
 
Picture1 = PictureBox

Timer1 = Timer


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephan Kirchmaier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephan-kirchmaier.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephan-kirchmaier-very-fast-image-processing__1-10781/archive/master.zip)

### API Declarations

```
Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Dim PicBits() As Byte, PicInfo As BITMAP, Cnt As Long
```


### Source Code

```
Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Dim PicBits() As Byte, PicInfo As BITMAP, Cnt As Long
Private Sub CF()
  Dim k As Long
  On Error Resume Next
  Picture1.Picture = Picture1.Image
  GetObject Picture1.Image, Len(PicInfo), PicInfo
  ReDim PicBits(1 To PicInfo.bmWidth * PicInfo.bmHeight * 3) As Byte
  GetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
  For Cnt = 2 To UBound(PicBits) + 1
    k = PicBits(Cnt - 1) + PicBits(Cnt + 1)
    k = k \ 2
    PicBits(Cnt) = k
  Next Cnt
  SetBitmapBits Picture1.Image, UBound(PicBits), PicBits(1)
  Picture1.Refresh
End Sub
Private Sub Timer1_Timer()
  Call CF
End Sub
```

