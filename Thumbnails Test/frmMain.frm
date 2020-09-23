VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Quick Thumbnails"
   ClientHeight    =   9960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   664
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   841
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HS 
      Height          =   255
      Left            =   10440
      Max             =   5
      Min             =   1
      TabIndex        =   5
      Top             =   9600
      Value           =   5
      Width           =   2055
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   9015
      Left            =   120
      ScaleHeight     =   597
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   821
      TabIndex        =   2
      Top             =   480
      Width           =   12375
      Begin VB.VScrollBar VS 
         Enabled         =   0   'False
         Height          =   8955
         Left            =   12060
         Max             =   10000
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
      Begin VB.FileListBox flImg 
         BackColor       =   &H008B6425&
         Height          =   1350
         Left            =   8280
         Pattern         =   "*.bmp;*.jpg;*.gif;*.png"
         TabIndex        =   6
         Top             =   7440
         Visible         =   0   'False
         Width           =   3000
      End
   End
   Begin VB.PictureBox picProg 
      AutoRedraw      =   -1  'True
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   797
      TabIndex        =   1
      Top             =   0
      Width           =   12015
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   5250
      TabIndex        =   8
      Top             =   9630
      Width           =   60
   End
   Begin VB.Label lblIS 
      AutoSize        =   -1  'True
      Caption         =   "Image Size"
      Height          =   210
      Left            =   9480
      TabIndex        =   4
      Top             =   9630
      Width           =   780
   End
   Begin VB.Label lblInf 
      AutoSize        =   -1  'True
      Caption         =   "Total : 0 item(s)"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   9630
      Width           =   1110
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By siguri92

Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion           As Long    'Must be 1 for Gdi+ v1.0, the current version as of this writing.
    DebugEventCallback       As Long    'Ignored on free builds
    SuppressBackgroundThread As Long    'FALSE unless you're prepared to call the hook/unhook functions properly
    SuppressExternalCodecs   As Long    'FALSE unless you want Gdi+ only to use its internal image codecs.
End Type

' GDI Plus API
Private Declare Sub GdiplusShutdown Lib "Gdiplus" (ByVal Token As Long)

Private Declare Function GdiplusStartup Lib "Gdiplus" (Token As Long, _
    InputBuffer As GdiplusStartupInput, Optional ByVal OutputBuffer As Long = 0) As Long

Private Declare Function GdipCreateFromHDC Lib "Gdiplus" (ByVal hDC As Long, Graphics As Long) As Long

Private Declare Function GdipDeleteGraphics Lib "Gdiplus" (ByVal Graphics As Long) As Long

Private Declare Function GdipLoadImageFromFile Lib "Gdiplus" (ByVal FileName As String, Image As Long) As Long

Private Declare Function GdipGetImageWidth Lib "Gdiplus" (ByVal Image As Long, Width As Long) As Long

Private Declare Function GdipGetImageHeight Lib "Gdiplus" (ByVal Image As Long, Height As Long) As Long

' For thumbnail
Private Declare Function GdipGetImageThumbnail Lib "Gdiplus" (ByVal Image As Long, _
    ByVal ThumbWidth As Long, ByVal ThumbHeight As Long, ThumbImage As Long, _
    Optional ByVal CallBack As Long = 0, Optional ByVal CallBackData As Long = 0) As Long
    
' Draw image
Private Declare Function GdipDrawImageRect Lib "Gdiplus" (ByVal Graphics As Long, _
    ByVal Image As Long, ByVal X As Single, ByVal Y As Single, _
    ByVal Width As Single, ByVal Height As Single) As Long

' Free memory
Private Declare Function GdipDisposeImage Lib "Gdiplus" (ByVal Image As Long) As Long

Private Type IMGRECT
    Left                    As Long
    Top                     As Long
    Width                   As Long
    Height                  As Long
End Type

Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function GetTickCount& Lib "kernel32" ()

' Image type
Private Type TYPEGDITHUMBNAILSIMAGE
    strFile                 As String
    mDC                     As Long
    hImage                  As Long
    cRECT                   As IMGRECT
    orgWidth                As Long
    orgHeight               As Long
    bSelected               As Boolean
End Type

' Def thumb size
Private Const m_def_Size& = 170

' Variables
Private mItem()             As TYPEGDITHUMBNAILSIMAGE
Private lImg()              As Long
Private nCount              As Long
Private TotalHeight         As Double
' For GDI startup and shutdown
Private Token               As Long


Private Sub cmdBrowse_Click()
On Error Resume Next
    Dim tmp$, mRect As RECT, i&
    tmp = BrowseFolder(ssfDESKTOP)
    If tmp = "" Then Exit Sub
    flImg.Path = tmp
    nCount = flImg.ListCount
    lblInf = "Total : " & flImg.ListCount & " item(s)"
    picProg.Cls
    SetRect mRect, 0, 0, picProg.ScaleWidth, picProg.ScaleHeight
    DrawText picProg.hDC, tmp, -1, mRect, &H20 Or &H4 Or &H1
    
    ReDim mItem(nCount - 1)
    
    For i = 0 To nCount - 1
        mItem(i).strFile = flImg.Path & "\" & flImg.List(i)
    Next
    
    GenerateThumb
End Sub

Private Sub GenerateThumb()
    Dim t&
    Dim i&, prog%

    t = GetTickCount
    
    ReDim lImg(nCount - 1)
    
    ' Load image data
    For i = 0 To nCount - 1
        Call GdipLoadImageFromFile(StrConv(mItem(i).strFile, vbUnicode), lImg(i))
        With mItem(i)
            Call GdipGetImageWidth(lImg(i), .orgWidth)
            Call GdipGetImageHeight(lImg(i), .orgHeight)
            ' Calculate image resize
            CalculateImageRect m_def_Size, .orgWidth, .orgHeight, .cRECT
        End With
    Next
    
    CalculateRect
    
    For i = 0 To nCount - 1
        Call GdipGetImageThumbnail(lImg(i), mItem(i).cRECT.Width, mItem(i).cRECT.Height, mItem(i).hImage)
        If i Mod 10 = 0 Then               ' Draw each 10 pics
            Draw
        End If
        lblLoading = "Generating " & i & "/" & (nCount - 1)
        DoEvents
    Next
    
    lblLoading = ""
    
    For i = 0 To nCount - 1
        Call GdipDisposeImage(lImg(i))
    Next
    
    MsgBox "Complete in " & (GetTickCount - t) / 1000 & " s"
    
    Me.Refresh
End Sub

' Calculate rect for arrange image
Private Sub CalculateRect()
    Dim i&, mSpc&
    Dim mL&, mT&, mS&
    mSpc = (HS.Value / HS.Max) * m_def_Size
    
    CalculateImageRect mSpc, _
        mItem(0).orgWidth, mItem(0).orgHeight, mItem(0).cRECT
        
    mItem(0).cRECT.Left = 0
    mItem(0).cRECT.Top = 0
    For i = 1 To nCount - 1
        ' Calculate again
        CalculateImageRect mSpc, _
            mItem(i).orgWidth, mItem(i).orgHeight, mItem(i).cRECT
        ' Spacing item
        mL = mItem(i - 1).cRECT.Left + mSpc + 20
        mT = mItem(i - 1).cRECT.Top
        If mL + mSpc > VS.Left Then                     ' Move down
            mL = 0
            mT = mItem(i - 1).cRECT.Top + mSpc + 20     ' New top
        End If
        ' Set item rect
        mItem(i).cRECT.Left = mL
        mItem(i).cRECT.Left = mL
        mItem(i).cRECT.Top = mT
    Next
    ' Get total height
    TotalHeight = (mItem(nCount - 1).cRECT.Top + mSpc + 20)
    VS.Enabled = CBool(TotalHeight > picDraw.ScaleHeight)
End Sub

Private Sub Draw()
On Error Resume Next
    Dim i&, lStart&, mGraphic&, mSpc&
    Dim StartPoint%, lL&, lT&
    Dim dH&
    Dim mRect As RECT
    
    dH = TotalHeight - picDraw.ScaleHeight
    ' Get start point
    lStart = IIf(VS.Enabled, (VS.Value / 10000) * dH, 0)
    
    mSpc = (HS.Value / HS.Max) * m_def_Size + 20
    
    picDraw.Cls
    
    Call GdipCreateFromHDC(picDraw.hDC, mGraphic)
    
    ' Start index
    StartPoint = Fix(lStart / mSpc) * Fix((picDraw.ScaleWidth - VS.Width) / mSpc)

    For i = StartPoint To nCount - 1
        With mItem(i)
            If .cRECT.Top - lStart > picDraw.ScaleHeight Then Exit For
            If .cRECT.Top - lStart > -mSpc Then
                ' Center rect
                lL = (mSpc - .cRECT.Width) / 2
                lT = ((mSpc - 20) - .cRECT.Height) / 2
                If .bSelected Then
                    FillRectEx picDraw.hDC, .cRECT.Left, .cRECT.Top - lStart, mSpc, mSpc, &H8B6425
                    FillRectEx picDraw.hDC, .cRECT.Left + 1, .cRECT.Top - lStart + 1, mSpc - 2, mSpc - 2, &HFDF2EA
                End If
                ' Draw image frame
                ' Outer
                FillRectEx picDraw.hDC, .cRECT.Left + lL - 4, .cRECT.Top - lStart + lT - 4, _
                                        .cRECT.Width + 8, .cRECT.Height + 8, &H808080
                ' Center
                FillRectEx picDraw.hDC, .cRECT.Left + lL - 3, .cRECT.Top - lStart + lT - 3, _
                                        .cRECT.Width + 6, .cRECT.Height + 6, vbWhite
                ' Inner
                FillRectEx picDraw.hDC, .cRECT.Left + lL - 1, .cRECT.Top - lStart + lT - 1, _
                                        .cRECT.Width + 2, .cRECT.Height + 2, &H808080
                ' Draw image
                Call GdipDrawImageRect(mGraphic, .hImage, .cRECT.Left + lL, _
                                        .cRECT.Top - lStart + lT, _
                                        .cRECT.Width, .cRECT.Height)
                SetRect mRect, .cRECT.Left, .cRECT.Top - lStart + mSpc - 15, _
                                .cRECT.Left + mSpc, (.cRECT.Top - lStart) + mSpc
                DrawText picDraw.hDC, flImg.List(i), -1, mRect, &H20 Or &H4 Or &H1
            End If
        End With
    Next
    
    picDraw.Refresh
    Call GdipDeleteGraphics(mGraphic)
End Sub

' Calculate for best ratio
Private Sub CalculateImageRect(lSize&, lWidth&, lHeight&, mRect As IMGRECT)
On Error Resume Next
    Dim lRatio As Double
    ' If not neccessary
    If (lWidth < lSize) And (lHeight < lSize) Then
        mRect.Width = lWidth
        mRect.Height = lHeight
        Exit Sub
    End If
    lRatio = lWidth / lHeight
    If lWidth >= lHeight Then            ' If width>height
        mRect.Width = lSize
        mRect.Height = lSize / lRatio
    Else                                 ' Else
        mRect.Height = lSize
        mRect.Width = lSize / lRatio
    End If
End Sub

Private Sub Form_Load()
    Dim mInput   As GdiplusStartupInput
    mInput.GdiplusVersion = 1
    If GdiplusStartup(Token, mInput) <> 0 Then          ' Unable to load GDI+
        MsgBox "Error loading GDIPlus!", vbExclamation + vbOKOnly
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Shutdown GDI
    GdiplusShutdown Token
End Sub

Private Sub HS_Scroll()
    CalculateRect
    Draw
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> vbLeftButton Then Exit Sub
    
    Dim i&, lStart&, mSpc&
    Dim dH&
    Dim mRect As RECT
    
    dH = TotalHeight - picDraw.ScaleHeight
    ' Get start point
    lStart = IIf(VS.Enabled, (VS.Value / 10000) * dH, 0)
    
    mSpc = (HS.Value / HS.Max) * m_def_Size + 20
    
    For i = 0 To nCount - 1
        mItem(i).bSelected = False
    Next
    
    For i = 0 To nCount - 1
        With mItem(i)
            If .cRECT.Top - lStart >= picDraw.ScaleHeight Then Exit For
            If X >= .cRECT.Left And X <= .cRECT.Left + mSpc Then                    ' Item selected
                If Y > .cRECT.Top - lStart And Y < .cRECT.Top - lStart + mSpc Then
                    mItem(i).bSelected = True
                    Exit For
                End If
            End If
        End With
    Next
    Draw
End Sub

Private Sub VS_Scroll()
    Draw
End Sub

' Draw filled rectangle
Private Sub FillRectEx(ByVal hDC&, X&, Y&, Width&, Height&, lColor&)
    Dim hBrush&, hResult&, mRect As RECT
    hBrush = CreateSolidBrush(lColor)
    SetRect mRect, X, Y, X + Width, Y + Height
    hResult = FillRect(hDC, mRect, hBrush)
    ' Clean up
    DeleteObject hBrush
    DeleteObject hResult
End Sub
