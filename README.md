<div align="center">

## Bitblt Load Striaght to Mem \[Animation\]

<img src="PIC200212123668269.gif">
</div>

### Description

This code will load a picture in mem on form load. Its show how also to do an animation with bitblt. Loading it in memory not only makes fps faster but the program itself.
 
### More Info
 
Satafaction


<span>             |<span>
---                |---
**Submitted On**   |2002-01-21 19:56:26
**By**             |[Josh Nixon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/josh-nixon.md)
**Level**          |Intermediate
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Bitblt\_Loa503321212002\.zip](https://github.com/Planet-Source-Code/josh-nixon-bitblt-load-striaght-to-mem-animation__1-31059/archive/master.zip)

### API Declarations

```
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 Const SRCAND = &H8800C6
 Const SRCPAINT = &HEE0086
```





