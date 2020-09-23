VERSION 5.00
Begin VB.UserControl MySlider 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   KeyPreview      =   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   58
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   144
   ToolboxBitmap   =   "MySlider.ctx":0000
   Begin VB.PictureBox picThumbVertical 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   3900
      Picture         =   "MySlider.ctx":0312
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   435
      Width           =   150
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   465
      Picture         =   "MySlider.ctx":03F4
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   0
      Top             =   495
      Width           =   75
   End
   Begin VB.Image imgBottom 
      Height          =   30
      Left            =   2235
      Picture         =   "MySlider.ctx":04D6
      Top             =   1740
      Width           =   60
   End
   Begin VB.Image imgTop 
      Appearance      =   0  'Flat
      Height          =   30
      Left            =   1995
      Picture         =   "MySlider.ctx":0530
      Top             =   870
      Width           =   60
   End
   Begin VB.Image imgMidVertical 
      Appearance      =   0  'Flat
      Height          =   135
      Left            =   1710
      Picture         =   "MySlider.ctx":058A
      Stretch         =   -1  'True
      Top             =   1095
      Width           =   60
   End
   Begin VB.Image imgRight 
      Height          =   60
      Left            =   2670
      Picture         =   "MySlider.ctx":0638
      Top             =   75
      Width           =   30
   End
   Begin VB.Image imgLeft 
      Height          =   60
      Left            =   360
      Picture         =   "MySlider.ctx":069A
      Top             =   225
      Width           =   30
   End
   Begin VB.Image imgMid 
      Height          =   60
      Left            =   555
      Picture         =   "MySlider.ctx":06FC
      Stretch         =   -1  'True
      Top             =   255
      Width           =   2295
   End
End
Attribute VB_Name = "MySlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private mMin As Long
Private mMax As Long
Private mValue As Long
Private mOrientation As sldOrientation
Private mDragging As Boolean
Private Const clrWhite As Long = 16777215
Private Const clrBlack As Long = 0
Private Const clrLtGrey As Long = 13160660
Private Const clrDkGrey As Long = 8421504
Private mLastX As Long
Private mLastY As Long


Public Event Change(ByVal NewValue As Long)



Public Enum sldOrientation
    sldHorizontal = 0
    sldVertical = 1
End Enum


Private Sub imgLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show what the position would be if the user were to click
    imgLeft.ToolTipText = mMin
End Sub


Private Sub imgLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Set value to min
    Value = mMin
End Sub


Private Sub imgTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show what the position would be if the user were to click
    imgTop.ToolTipText = mMin
End Sub


Private Sub imgTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Set value to min
    Value = mMin
End Sub


Private Sub imgRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show what the position would be if the user were to click
    imgRight.ToolTipText = mMax
End Sub


Private Sub imgRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Set value to max
    Value = mMax
End Sub


Private Sub imgBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show what the position would be if the user were to click
    imgBottom.ToolTipText = mMax
End Sub


Private Sub imgBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Set value to max
    Value = mMax
End Sub


Private Sub imgMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show what the position would be if the user were to click
    Dim Max As Long
    Max = ScaleWidth - picThumb.Width
    imgMid.ToolTipText = Int((ScaleX(X, vbTwips, vbPixels) * (mMax - mMin)) / Max) + mMin
End Sub


Private Sub imgMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Set value via click position on bar
    Dim Max As Long
    Max = ScaleWidth - picThumb.Width
    Value = Int((ScaleX(X, vbTwips, vbPixels) * (mMax - mMin)) / Max) + mMin
End Sub


Private Sub imgMidVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show what the position would be if the user were to click
    Dim Max As Long
    Max = ScaleHeight - picThumbVertical.Height
    imgMidVertical.ToolTipText = Int((ScaleY(Y, vbTwips, vbPixels) * (mMax - mMin)) / Max) + mMin
End Sub


Private Sub imgMidVertical_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Set value via click position on bar
    Dim Max As Long
    Max = ScaleHeight - picThumbVertical.Height
    Value = Int((ScaleY(Y, vbTwips, vbPixels) * (mMax - mMin)) / Max) + mMin
End Sub


Private Sub picThumb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Start dragging the thumb
    mDragging = True
    mLastX = X
End Sub


Private Sub picThumb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Drag the thumb horizontally
    Dim NewX As Long
    Dim Max As Long
    Max = ScaleWidth - picThumb.Width
    If mDragging = True Then
        If X <> mLastX Then
            mLastX = X
            NewX = picThumb.Left + X
            If NewX < 0 Then
                NewX = 0
            ElseIf NewX > Max Then
                NewX = Max
            End If
            picThumb.Left = NewX
            SetValueFromThumb
        End If
    End If
End Sub


Private Sub picThumb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Stop dragging the thumb
    mDragging = False
End Sub


Private Sub picThumbVertical_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Start dragging the thumb
    mDragging = True
    mLastY = Y
End Sub

Private Sub picThumbVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Drag the thumb vertically
    Dim NewY As Long
    Dim Max As Long
    Max = ScaleHeight - picThumbVertical.Height
    If mDragging = True Then
        If Y <> mLastY Then
            mLastY = Y
            NewY = picThumbVertical.Top + Y
            If NewY < 0 Then
                NewY = 0
            ElseIf NewY > Max Then
                NewY = Max
            End If
            picThumbVertical.Top = NewY
            SetValueFromThumb
        End If
    End If
End Sub


Private Sub picThumbVertical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Stop dragging the thumb
    mDragging = False
End Sub


Private Sub SetValueFromThumb()
    ' Set value from thumb position
    Dim Max As Long
    If mOrientation = sldVertical Then
        Max = ScaleHeight - picThumbVertical.Height
        Value = (picThumbVertical.Top * (mMax - mMin)) / Max + mMin
    Else
        Max = ScaleWidth - picThumb.Width
        Value = (picThumb.Left * (mMax - mMin)) / Max + mMin
    End If
End Sub


Private Sub SetThumbFromValue()
    ' Set thumb position from value
    Dim Max As Long
    Dim X As Long
    Dim Y As Long
    
    ' Set vertical thumb
    Max = ScaleHeight - picThumbVertical.Height
    picThumbVertical.Top = (mValue - mMin) / (mMax - mMin) * Max
    picThumbVertical.ToolTipText = mValue
    
   ' Set horizontal thumb
    Max = ScaleWidth - picThumb.Width
    picThumb.Left = (mValue - mMin) / (mMax - mMin) * Max
    picThumb.ToolTipText = mValue
End Sub


Private Sub UserControl_InitProperties()
    ' Set default properties (executed once when control is added to form)
    MinValue = 0
    MaxValue = 100
    Orientation = sldHorizontal
    Value = 50
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
        Value = Value - 1
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
        Value = Value + 1
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Load the properties
    MinValue = PropBag.ReadProperty("MinValue", 0)
    MaxValue = PropBag.ReadProperty("MaxValue", 100)
    Orientation = PropBag.ReadProperty("Orientation", sldHorizontal)
    Value = PropBag.ReadProperty("Value", 50)
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Save the properties
    PropBag.WriteProperty "MinValue", mMin, 0
    PropBag.WriteProperty "MaxValue", mMax, 100
    PropBag.WriteProperty "Value", mValue, 50
    PropBag.WriteProperty "Orientation", mOrientation, sldHorizontal
End Sub


Private Sub UserControl_Resize()
    ' Resize the slider bar
    Dim X As Long
    Dim Y As Long
    Dim w As Long
    Dim h As Long
    Dim pos As Long
    
    ' Draw the horizontal bar
    X = 0
    Y = Int(ScaleHeight / 2)
    w = ScaleWidth
    h = imgMid.Height
    ' Draw the line
    imgMid.Move X, Y, w, h
    imgLeft.Move X, Y
    imgRight.Move w - 2, Y
    ' Draw the thumb
    pos = mValue * Int((w - picThumb.Width)) / mMax
    picThumb.Move pos - Int(picThumb.Width / 2) + 2, Y - picThumb.Height / 2 + 2
    
    ' Draw the vertical bar
    X = Int(ScaleWidth / 2)
    Y = 0
    w = imgMidVertical.Width
    h = ScaleHeight
    ' Draw the line
    imgMidVertical.Move X, Y, w, h
    imgTop.Move X, Y
    imgBottom.Move w, Y - 2
    ' Draw the thumb
    pos = mValue * Int((h - picThumbVertical.Height)) / mMax
    picThumbVertical.Move X - picThumbVertical.Width / 2 + 2, pos - Int(picThumbVertical.Height / 2) + 2
    
    'SetThumbFromValue
End Sub


Public Property Let MinValue(ByVal Value As Long)
    ' Set the minimum value (no negative numbers)
    If Value < 0 Then
        Value = 0
    End If
    mMin = Value
    PropertyChanged "MinValue"
    If mValue < mMin Then
        mValue = mMin
        PropertyChanged "Value"
    End If
    If mMax < mMin + 1 Then
        mMax = mMin + 1
        PropertyChanged "MaxValue"
    End If
End Property


Public Property Get MinValue() As Long
    ' Get the minimum value
    MinValue = mMin
End Property


Public Property Let MaxValue(ByVal Value As Long)
    ' Set the maximum value (must be greater than min)
    If Value < mMin + 1 Then
        Value = mMin + 1
    End If
    mMax = Value
    PropertyChanged "MaxValue"
    If mValue > mMax Then
        mValue = mMax
        PropertyChanged "Value"
    End If
End Property


Public Property Get MaxValue() As Long
    ' Get the maximum value
    MaxValue = mMax
End Property


Public Property Let Value(ByVal NewValue As Long)
    ' Set the value (must be >= min and <= max)
    If NewValue = mValue Then
        Exit Property
    End If
    If NewValue < mMin Then
        NewValue = mMin
    ElseIf NewValue > mMax Then
        NewValue = mMax
    End If
    mValue = NewValue
    SetThumbFromValue
    PropertyChanged "Value"
    RaiseEvent Change(NewValue)
End Property


Public Property Get Value() As Long
    ' Get the value
    Value = mValue
End Property


Public Property Let Orientation(ByVal Value As sldOrientation)
    ' Set the orientation, horizontal or vertical
    mOrientation = Value
    
    picThumbVertical.Visible = False
    imgTop.Visible = False
    imgBottom.Visible = False
    imgMidVertical.Visible = False
    
    picThumb.Visible = False
    imgLeft.Visible = False
    imgRight.Visible = False
    imgMid.Visible = False
    
    If mOrientation = sldVertical Then
        picThumbVertical.Visible = True
        imgTop.Visible = True
        imgBottom.Visible = True
        imgMidVertical.Visible = True
    Else
        picThumb.Visible = True
        imgLeft.Visible = True
        imgRight.Visible = True
        imgMid.Visible = True
    End If
    
    PropertyChanged "Orientation"
End Property


Public Property Get Orientation() As sldOrientation
    ' Get the orientation, horizontal or vertical
    Orientation = mOrientation
End Property
