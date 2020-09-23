VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Sub Form_Load()
    Dim b() As Byte, pic As StdPicture, DrawDirectlyOnForm As Boolean
    
    AutoRedraw = True
    DrawDirectlyOnForm = False
    
    'Load a picture into the byte array
    b = LoadFile(App.Path & "\full color.jpg")
    
    'Create a StdPicture object (bitmap object) from the bytestream
    Set pic = PictureFromByteStream(b)
    If pic Is Nothing Then
        MsgBox "Unable to load bitmap! Check filename"
        Exit Sub
    End If
    
    'Now, there are two ways to display the picture. You can either:
    If DrawDirectlyOnForm = True Then
        'Assign it directly to the picture property of the form
        Set Me.Picture = pic
    Else
        'Or select it into a DC and do other manipulations to it
        DoItTheHardWay pic
    End If

    'Destroy it when you're done.
    Set pic = Nothing
    
End Sub

Private Sub DoItTheHardWay(ByRef pic As StdPicture)
    Dim TempDC As Long, hBmp As Long, w As Long, h As Long, bmpInfo As BITMAP
    
    'Determine the width and height of the bitmap
    GetObject pic.Handle, Len(bmpInfo), bmpInfo
    w = bmpInfo.bmWidth
    h = bmpInfo.bmHeight
    
    'Create a DC compatible with the bitmap
    TempDC = CreateCompatibleDC(0)
    
    'Select the bitmap into it
    hBmp = SelectObject(TempDC, pic.Handle)
    
    'Blit it to the form
    BitBlt Me.hdc, 0, 0, w, h, TempDC, 0, 0, vbSrcCopy
    
    'Clean up
    hBmp = SelectObject(TempDC, hBmp)
    DeleteDC TempDC

End Sub
