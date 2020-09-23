VERSION 5.00
Begin VB.UserControl XPSlider 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000006&
   ScaleHeight     =   70
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   205
   ToolboxBitmap   =   "XPSlider.ctx":0000
   Begin VB.CommandButton cmdRight 
      BackColor       =   &H00000080&
      Height          =   195
      Left            =   2670
      TabIndex        =   2
      Top             =   45
      Width           =   240
   End
   Begin VB.CommandButton cmdLeft 
      BackColor       =   &H00000080&
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   45
      Width           =   240
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H008080FF&
      Height          =   120
      Left            =   420
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   0
      Top             =   90
      Width           =   2265
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   165
         Left            =   1140
         TabIndex        =   3
         Top             =   -30
         Width           =   375
      End
      Begin VB.Image imgKnob 
         Height          =   120
         Left            =   0
         Picture         =   "XPSlider.ctx":0312
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "XPSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMin              As Long         'Minimum value range
Private mMax              As Long         'Maximum value range
Private mValue            As Long         'Current Value
Private mSliderWH      As Long
Private mBaseColor          As OLE_COLOR
Private mMidColor          As OLE_COLOR

Event Changed()

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
'------------------------------------------------------------
'draw and set rectangular area of the control
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'draw by pixel or by line
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'select and delete created objects
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'create regions of pixels and remove them to make the control transparent
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF As Long = 4

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Const m_def_ValueVis = True
Private Const m_def_ValueCol = vbBlack

Private m_ValueVis As Boolean
Private m_ValueCol As OLE_COLOR

Private rc As RECT
Private W As Long, H As Long
Private regMain As Long, rgn1 As Long
Private R As Long, l As Long, t As Long, B As Long

Private Sub UserControl_Initialize()
   m_ValueVis = m_def_ValueVis
   m_ValueCol = m_def_ValueCol
End Sub

Private Sub UserControl_Resize()
  GetClientRect UserControl.hwnd, rc
  With rc
    R = .Right - 1: l = .Left: t = .Top: B = .Bottom
    W = .Right: H = .Bottom
  End With
  
  UserControl.Cls
  UserControl.Height = 306
  DrawButton
  DrawRail
  'set position and size of controls
  cmdRight.Left = UserControl.ScaleWidth - 30
  cmdRight.Top = 3
  cmdLeft.Left = 12
  cmdLeft.Top = 3
  pic1.Left = 28
  pic1.Top = 6
  pic1.Width = cmdRight.Left - (cmdRight.Width * 2 - 15)
  Label1.FontName = "Tahoma"
  Label1.FontSize = 7
  Label1.Left = imgKnob.Left + imgKnob.Width + 5
End Sub

Private Sub UserControl_InitProperties()
    mMin = 0
    mMax = 100
    mValue = 0
    mSliderWH = 400
    mBaseColor = vbBlue
    mMidColor = &HFFFFFE
    PositionSlider
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
        mMin = .ReadProperty("Min", 0)
        mMax = .ReadProperty("Max", 100)
        mValue = .ReadProperty("Value", 50)
        mSliderWH = .ReadProperty("SliderWid_Height", 315)
        BaseColor = .ReadProperty("BaseColor", vbBlue)
        MidColor = .ReadProperty("MidColor", mMidColor)
        ValueVis = .ReadProperty("ValueVis", m_def_ValueVis)
        ValueCol = .ReadProperty("ValueCol", m_def_ValueCol)
    End With
    PositionSlider
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Min", mMin, 0
        .WriteProperty "Max", mMax, 100
        .WriteProperty "Value", mValue, 50
        .WriteProperty "SliderWid_Height", mSliderWH, 315
        .WriteProperty "BaseColor", mBaseColor, vbBlue
        .WriteProperty "MidColor", mMidColor, vbWhite
        .WriteProperty "ValueVis", m_ValueVis, m_def_ValueVis
        .WriteProperty "ValueCol", m_ValueCol, m_def_ValueCol
    End With
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngPos                  As Long
Dim sglScale                As Single
    
    With imgKnob
            lngPos = ((x - mSliderWH / 2) \ 15) * 16
            If lngPos < 0 Then
                lngPos = 0
            ElseIf lngPos > pic1.ScaleWidth - mSliderWH Then
                lngPos = pic1.ScaleWidth - mSliderWH
            End If
            
            .Left = lngPos
            sglScale = (pic1.ScaleWidth - mSliderWH) / (mMax - mMin)
            mValue = (lngPos / sglScale) + mMin
            RaiseEvent Changed
    End With
    PositionSlider
    If ValueVis = False Then Exit Sub
      If mValue <= Max / 2 Then
         Label1.Left = imgKnob.Left + imgKnob.Width + 5
      Else
         Label1.Left = imgKnob.Left - imgKnob.Width - (Label1.Width / 2)
       End If
     Label1.Caption = mValue
   
End Sub

Private Sub imgKnob_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim lngPos                  As Long
Dim sglScale                As Single

    If Button = vbLeftButton Then
        With imgKnob
                lngPos = ((.Left + x - mSliderWH / 2) \ 15) * 16
                If lngPos < 0 Then
                    lngPos = 0
                ElseIf lngPos > pic1.ScaleWidth - mSliderWH Then
                    lngPos = pic1.ScaleWidth - mSliderWH
                End If
                             
                .Left = lngPos
                sglScale = ((pic1.ScaleWidth - mSliderWH)) / (mMax - mMin)
                mValue = (lngPos / sglScale) + mMin
                RaiseEvent Changed
                If ValueVis = False Then Exit Sub
                   If mValue <= Max / 2 Then
                      Label1.Left = imgKnob.Left + imgKnob.Width + 5
                   Else
                      Label1.Left = imgKnob.Left - imgKnob.Width - (Label1.Width / 2)
                   End If
                   Label1.Caption = mValue
                
        End With
    End If
    
End Sub

Private Function PositionSlider()

Dim sglScale                As Single

    With imgKnob
        If mMax - mMin <> 0 Then
            sglScale = (pic1.ScaleWidth - mSliderWH) / (mMax - mMin)
            .Left = (mValue - mMin) * sglScale
        End If
    End With
End Function

Private Sub DrawButton()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim I As Long, ColorR As Long, ColorG As Long, ColorB As Long
Dim hBrush As Long

  With UserControl
  
    hBrush = CreateSolidBrush(RGB(0, 60, 116))
    FrameRect UserControl.hDC, rc, hBrush
    DeleteObject hBrush
    
    'Left top corner
    SetPixel .hDC, l, t + 1, RGB(122, 149, 168)
    SetPixel .hDC, l + 1, t + 1, RGB(37, 87, 131)
    SetPixel .hDC, l + 1, t, RGB(122, 149, 168)
    
    'right top corner
    SetPixel .hDC, R - 1, t, RGB(122, 149, 168)
    SetPixel .hDC, R - 1, t + 1, RGB(37, 87, 131)
    SetPixel .hDC, R, t + 1, RGB(122, 149, 168)
    
    'left bottom corner
    SetPixel .hDC, l, B - 2, RGB(122, 149, 168)
    SetPixel .hDC, l + 1, B - 2, RGB(37, 87, 131)
    SetPixel .hDC, l + 1, B - 1, RGB(122, 149, 168)
    
    'right bottom corner
    SetPixel .hDC, R, B - 2, RGB(122, 149, 168)
    SetPixel .hDC, R - 1, B - 2, RGB(37, 87, 131)
    SetPixel .hDC, R - 1, B - 1, RGB(122, 149, 168)
  End With
  
  DeleteObject regMain
  regMain = CreateRectRgn(0, 0, W, H)
  rgn1 = CreateRectRgn(0, 0, 1, 1)            'Left top coner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(0, H - 1, 1, H)      'Left bottom corner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(W - 1, 0, W, 1)      'Right top corner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(W - 1, H - 1, W, H) 'Right bottom corner
  CombineRgn regMain, regMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  SetWindowRgn UserControl.hwnd, regMain, True
  
  'draw screws
  UserControl.DrawWidth = 1
  UserControl.Circle (8, UserControl.ScaleHeight - 10), 3, vbBlack        'left screw bottom
  UserControl.Line (8, UserControl.ScaleHeight - 12)-(9, UserControl.ScaleHeight - 6), &H404040
  UserControl.Circle (UserControl.ScaleWidth - 10, UserControl.ScaleHeight - 10), 3, vbBlack        'right screw bottom
  UserControl.Line (UserControl.ScaleWidth - 12, UserControl.ScaleHeight - 12)-(UserControl.ScaleWidth - 7, UserControl.ScaleHeight - 7), &H404040
End Sub

Private Sub DrawRail()
Dim x As Integer
pic1.ScaleMode = 3
For x = 0 To 7
   Select Case x
      Case 0, 7
         pic1.ForeColor = BlendColors(BaseColor, MidColor, 45)
      Case 1, 6
         pic1.ForeColor = BlendColors(BaseColor, MidColor, 62)
      Case 2, 5
         pic1.ForeColor = BlendColors(BaseColor, MidColor, 72)
      Case 3
         pic1.ForeColor = BlendColors(BaseColor, MidColor, 100)
      Case 4
         pic1.ForeColor = BlendColors(BaseColor, MidColor, 82)
   End Select
   pic1.Line (0, x)-(pic1.ScaleWidth, x)
Next x
   pic1.Line (0, 0)-(0, 8)
   pic1.Line (pic1.ScaleWidth - 1, 0)-(pic1.ScaleWidth - 1, 8)
   pic1.ScaleMode = 1
End Sub

Private Sub cmdRight_Click()
   Value = Value + 1

   If ValueVis = False Then Exit Sub
      If mValue <= Max / 2 Then
         Label1.Left = imgKnob.Left + imgKnob.Width + 5
      Else
         Label1.Left = imgKnob.Left - imgKnob.Width - (Label1.Width / 2)
       End If
     Label1.Caption = mValue
   
End Sub

Private Sub cmdLeft_Click()
   Value = Value - 1
 
   If ValueVis = False Then Exit Sub
      If mValue <= Max / 2 Then
         Label1.Left = imgKnob.Left + imgKnob.Width + 5
      Else
         Label1.Left = imgKnob.Left - imgKnob.Width - (Label1.Width / 2)
       End If
     Label1.Caption = mValue
   
End Sub

Public Sub GetRGB(R As Integer, G As Integer, B As Integer, ByVal Color As Long)
    Dim TempValue As Long
    
    TranslateColor Color, 0, TempValue
    
    R = TempValue And &HFF&
    G = (TempValue And &HFF00&) / 2 ^ 8
    B = (TempValue And &HFF0000) / 2 ^ 16
End Sub

Public Function BlendColors(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Percentage As Single) As Long
    On Error Resume Next
    
    Dim R(2) As Integer, G(2) As Integer, B(2) As Integer
    Dim fPercentage(2) As Single
    Dim DAmt(2) As Single
    
    Percentage = SetBound(Percentage, 0, 100)
    
    GetRGB R(0), G(0), B(0), Color1
    GetRGB R(1), G(1), B(1), Color2
    
    DAmt(0) = R(1) - R(0): fPercentage(0) = (DAmt(0) / 100) * Percentage
    DAmt(1) = G(1) - G(0): fPercentage(1) = (DAmt(1) / 100) * Percentage
    DAmt(2) = B(1) - B(0): fPercentage(2) = (DAmt(2) / 100) * Percentage
    
    R(2) = R(0) + fPercentage(0)
    G(2) = G(0) + fPercentage(1)
    B(2) = B(0) + fPercentage(2)
    
    BlendColors = RGB(R(2), G(2), B(2))
End Function

Private Function SetBound(ByVal Num As Single, ByVal MinNum As Single, ByVal MaxNum As Single) As Single

    If Num < MinNum Then
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        SetBound = MaxNum
    Else
        SetBound = Num
    End If
End Function

Public Property Get BaseColor() As OLE_COLOR
   BaseColor = mBaseColor
End Property

Public Property Let BaseColor(NewValue As OLE_COLOR)
   mBaseColor = NewValue
   PropertyChanged "BaseColor"
   UserControl_Resize
   DrawRail
End Property

Public Property Get MidColor() As OLE_COLOR
   MidColor = mMidColor
End Property

Public Property Let MidColor(NewValue As OLE_COLOR)
   mMidColor = NewValue
   If mMidColor = vbWhite Then mMidColor = &HFFFFFE  ' Does'nt like vbWhite (HFFFFFF),I think its because of something in the BlendColors Procedure
   PropertyChanged "MidColor"
   UserControl_Resize
   DrawRail
End Property

Public Property Get Min() As Long
    Min = mMin
End Property

Public Property Let Min(NewValue As Long)

    If NewValue <= mMax Then
        mMin = NewValue
        If mValue < mMin Then
            mValue = mMin
            PropertyChanged "Value"
        End If
        PositionSlider
        PropertyChanged "Min"
    End If
End Property

Public Property Get Max() As Long
    Max = mMax
End Property

Public Property Let Max(NewValue As Long)
    If NewValue > mMin Then
        mMax = NewValue
        If mValue > mMax Then
            mValue = mMax
            PropertyChanged "Value"
        End If
        PositionSlider
        PropertyChanged "Max"
    End If
End Property

Public Property Get SliderWid_Height() As Long
    SliderWid_Height = mSliderWH
End Property

Public Property Let SliderWid_Height(NewValue As Long)

   mSliderWH = NewValue
   pic1.Width = mSliderWH
   pic1.Height = UserControl.Height
   PositionSlider
   PropertyChanged "SliderWid_Height"
   UserControl_Resize
End Property

Public Property Get Value() As Long
    Value = mValue
End Property

Public Property Let Value(NewValue As Long)

    If NewValue < mMin Then
        NewValue = mMin
    
    ElseIf NewValue > mMax Then
        NewValue = mMax
    End If
    
    mValue = NewValue
    PositionSlider
    
    PropertyChanged "Value"
    RaiseEvent Changed
    
End Property

Public Property Get ValueVis() As Boolean
   ValueVis = m_ValueVis
End Property

Public Property Let ValueVis(NewValueVis As Boolean)
  m_ValueVis = NewValueVis
  Label1.Visible = m_ValueVis
  PropertyChanged "ValueVis"
  UserControl_Resize
End Property

Public Property Get ValueCol() As OLE_COLOR
   ValueCol = m_ValueCol
End Property

Public Property Let ValueCol(NewValueCol As OLE_COLOR)
  m_ValueCol = NewValueCol
  Label1.ForeColor = m_ValueCol
  PropertyChanged "ValueCol"
  UserControl_Resize
End Property

