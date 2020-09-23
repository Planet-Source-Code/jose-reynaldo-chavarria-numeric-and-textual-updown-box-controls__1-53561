VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.UserControl JRNumericUpDown 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1575
   PropertyPages   =   "JRNumericUpDown.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   1575
   ToolboxBitmap   =   "JRNumericUpDown.ctx":0023
   Begin ComCtl2.UpDown udHorizontal 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   327681
      Max             =   -8
      Orientation     =   1
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown udVertical 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Max             =   -8
      Enabled         =   -1  'True
   End
   Begin VB.TextBox TxtInside 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtBody 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numeric UpDown"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "JRNumericUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'/////////////////////////////////////////////////////////////////
'///             Numeric Up Down Control
'///                 (JRUpDown.ctl)
'///______________________________________________________________
'/// Numeric updown box.
'///______________________________________________________________
'/// Last modification  : Apr/26/2004
'/// Last modified by   : Jose Reynaldo Chavarria
'/// Modification reason: Creation
'/// Author: Jose Reynaldo Chavarria Q. (jchavarria@agrolibano.hn)
'/// Tested on: Windows XP SP1, Windows 98, Windows NT 4.0 SP6.0
'/// Agrolibano S.A. de C.V.
'/// Web Site: http://www.agrolibano.hn/
'/////////////////////////////////////////////////////////////////

Dim mDirection As Integer, m_Direction As Integer '0=Down, 1=Up
Dim m_ReadOnly As Boolean
Dim m_Max As Double
Dim m_Min As Double
Dim m_Value As Double
Dim m_Increment  As Double
Dim m_Wrap As Boolean
Dim m_UseArrowKeys As Boolean
Dim m_Enabled As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_DecimalPlaces As Integer
Dim m_Orientation As jrOrientationConstants
Dim m_ShowFocusRect As Boolean
Dim m_UpDownAlignment As jrUDAlignConstants
Dim WithEvents m_UpDown As UpDown
Attribute m_UpDown.VB_VarHelpID = -1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Uncoment this enums if use this
'control alone as part of your Project

'Public Enum jrOrientationConstants
'    jroVertical = 0
'    jroHorizontal = 1
'End Enum
'Public Enum jrUDAlignConstants
'    jrRighty = 0
'    jrLefty = 1
'End Enum
'Public Enum jrBorderConstants
'    jrDouble3D = 0
'    jrSingle3D = 1
'    jrNone = 2
'End Enum

Const GWL_STYLE = (-16)
Const ES_NUMBER = &H2000&

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long

'UX Theme API
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long

'OS Version API
Private Declare Function GetVersionEx& Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) 'As Long
    
'Private Const VER_PLATFORM_WIN32_NT = 2
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
'Private Const VER_PLATFORM_WIN32s = 0

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 'Maintenance string for PSS usage
End Type

Public Event Change(ByVal PreviousValue As Double, ByVal NewValue As Double)

Private Sub SetNumber(NumberText As TextBox, Flag As Boolean)
    Dim Curstyle As Long, Newstyle As Long
    'retrieve the window style
    Curstyle = GetWindowLong(NumberText.hWnd, GWL_STYLE)
    If Flag Then
       Curstyle = Curstyle Or ES_NUMBER
    Else
       Curstyle = Curstyle And (Not ES_NUMBER)
    End If
    'Set the new style
    Newstyle = SetWindowLong(NumberText.hWnd, GWL_STYLE, Curstyle)
    'refresh
    NumberText.Refresh
End Sub

'========================================
'PROPERTIES
'========================================
Public Property Get EdithWnd() As Long
    EditHandle = TxtBody.hWnd
End Property

Public Property Get UpDownhWnd() As Long
    If Not m_UpDown Is Nothing Then
        UpDownhWnd = m_UpDown.hWnd
    End If
End Property

Public Property Let UpDownAlignment(NewValue As jrUDAlignConstants)
    m_UpDownAlignment = NewValue
    UserControl_Resize
    PropertyChanged "UpDownAlignment"
End Property
Public Property Get UpDownAlignment() As jrUDAlignConstants
    UpDownAlignment = m_UpDownAlignment
End Property

Public Property Let ShowFocusRect(NewValue As Boolean)
    m_ShowFocusRect = NewValue
    If Not NewValue Then
        'clear focus in case there is
        TxtInside.Refresh
    End If
    PropertyChanged "ShowFocusRect"
End Property
Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let Orientation(NewValue As jrOrientationConstants)
    m_Orientation = NewValue
    If NewValue = jroVertical Then
        Set m_UpDown = udVertical
        udVertical.Max = udHorizontal.Max
        udVertical.Min = udHorizontal.Min
        udVertical.Wrap = udHorizontal.Wrap
        udVertical.Enabled = udHorizontal.Enabled
        UpdateUpDownValues
        UserControl_Resize
        udHorizontal.Visible = False
        udVertical.Visible = True
    Else
        Set m_UpDown = udHorizontal
        udHorizontal.Max = udVertical.Max
        udHorizontal.Min = udVertical.Min
        udHorizontal.Wrap = udVertical.Wrap
        udHorizontal.Enabled = udVertical.Enabled
        UpdateUpDownValues
        UserControl_Resize
        udVertical.Visible = False
        udHorizontal.Visible = True
    End If
    PropertyChanged "Orientation"
End Property
Public Property Get Orientation() As jrOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = TxtInside.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    TxtInside.Alignment = New_Alignment
    SetNumber TxtInside, True
    PropertyChanged "Alignment"
End Property

Public Property Get Font() As Font
    Set Font = TxtInside.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set TxtInside.Font = New_Font
    Set UserControl.Font = New_Font
    Set TxtBody.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = TxtInside.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    TxtInside.ForeColor = New_ForeColor
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Let ReadOnly(NewValue As Boolean)
    m_ReadOnly = NewValue
    If NewValue Then
        TxtInside.MousePointer = vbArrow
    Else
        TxtInside.MousePointer = vbIbeam
    End If
    TxtInside.Locked = NewValue
    DisplayCaret m_ReadOnly
    PropertyChanged "ReadOnly"
End Property
Public Property Get ReadOnly() As Boolean
    ReadOnly = m_ReadOnly
    TxtInside.Locked = m_ReadOnly
End Property

Public Property Let Max(NewValue As Double)
    m_Max = NewValue
    UpdateUpDownValues
    PropertyChanged "Max"
End Property
Public Property Get Max() As Double
    Max = m_Max
End Property

Public Property Let Min(NewValue As Double)
    m_Min = NewValue
    UpdateUpDownValues
    PropertyChanged "Min"
End Property
Public Property Get Min() As Double
    Min = m_Min
End Property

Public Property Let Increment(NewValue As Double)
    m_Increment = NewValue
    UpdateUpDownValues
    PropertyChanged "Increment"
End Property
Public Property Get Increment() As Double
    Increment = m_Increment
End Property

Public Property Let Value(NewValue As Double)
    Dim OldVal As Double
    Select Case NewValue
        Case Is > m_Max
            NewValue = m_Max
        Case Is < m_Min
            NewValue = m_Min
        Case Else
            NewValue = m_Increment * CLng(NewValue / m_Increment)
    End Select
    OldVal = m_Value
    m_Value = NewValue
    
    If m_UpDown.Value <> CUDVal(NewValue) Then
        m_UpDown.Value = CUDVal(NewValue)
    End If
    TxtInside.Text = Format(m_Value, "#,##0" & IIf(m_DecimalPlaces > 0, "." & String(m_DecimalPlaces, "0"), ""))
    RaiseEvent Change(OldVal, NewValue)
    PropertyChanged "Value"
End Property
Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Wrap(NewValue As Boolean)
    m_Wrap = NewValue
    m_UpDown.Wrap = NewValue
    PropertyChanged "Wrap"
End Property
Public Property Get Wrap() As Boolean
    Wrap = m_UpDown.Wrap
End Property

Public Property Let UseArrowKeys(NewValue As Boolean)
    m_UseArrowKeys = NewValue
    PropertyChanged "UseArrowKeys"
End Property
Public Property Get UseArrowKeys() As Boolean
    UseArrowKeys = m_UseArrowKeys
End Property

Public Property Let Enabled(NewValue As Boolean)
    m_Enabled = NewValue
    TxtInside.Enabled = NewValue
    m_UpDown.Enabled = NewValue
    PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let BackColor(NewValue As OLE_COLOR)
    m_BackColor = NewValue
    TxtBody.Enabled = True
    TxtBody.BackColor = NewValue
    TxtInside.BackColor = NewValue
    TxtBody.Enabled = False
    PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let DecimalPlaces(NewValue As Integer)
    m_DecimalPlaces = NewValue
    TxtInside.Text = Format(m_Value, "#,##0" & IIf(m_DecimalPlaces > 0, "." & String(m_DecimalPlaces, "0"), ""))
    PropertyChanged "DecimalPlaces"
End Property
Public Property Get DecimalPlaces() As Integer
    DecimalPlaces = m_DecimalPlaces
End Property

Private Sub TxtInside_Change()
    DisplayCaret m_ReadOnly
    If m_ReadOnly Then
        DrawTextFocusRect TxtInside
    End If
End Sub

Private Sub TxtInside_Click()
    If m_ReadOnly Then
        DrawTextFocusRect TxtInside
    End If
End Sub

Private Sub TxtInside_DblClick()
    If m_ReadOnly Then
        DrawTextFocusRect TxtInside
    End If
End Sub

'========================================
'END Properties
'========================================

Private Sub TxtInside_GotFocus()
    If m_ReadOnly Then
        TxtInside.BackColor = vbHighlight
        TxtInside.ForeColor = vbHighlightText
'        DrawFocus
        DrawTextFocusRect TxtInside
    Else
        SelectAll
    End If
    DisplayCaret m_ReadOnly
End Sub

Private Sub TxtInside_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeySubtract, 189
        If m_UpDown.Min < 0 Or m_UpDown.Min < 0 Then
            TxtInside.SelText = "-"
        End If
    Case vbKeyDecimal, 190
        If m_DecimalPlaces > 0 Then
            TxtInside.SelText = "."
        End If
        
    Case vbKeyUp, vbKeyPageUp
        If m_UseArrowKeys And (m_Orientation = jroVertical) Then
            If Val(TxtInside.Text) <> m_UpDown.Value Then
                TxtInside_Validate False
            End If
            GoUP
        End If
        
    Case vbKeyDown, vbKeyPageDown
        If m_UseArrowKeys And (m_Orientation = jroVertical) Then
            If Val(TxtInside.Text) <> m_UpDown.Value Then
                TxtInside_Validate False
            End If
            GoDown
        End If
        
    Case vbKeyRight
        If m_UseArrowKeys And (m_Orientation = jroHorizontal) Then
            If Val(TxtInside.Text) <> m_UpDown.Value Then
                TxtInside_Validate False
            End If
            GoUP
        End If
        
    Case vbKeyLeft
        If m_UseArrowKeys And (m_Orientation = jroHorizontal) Then
            If Val(TxtInside.Text) <> m_UpDown.Value Then
                TxtInside_Validate False
            End If
            GoDown
        End If
        
    End Select
End Sub

Private Sub TxtInside_LostFocus()
    TxtInside.BackColor = m_BackColor 'vbWindowBackground
    TxtInside.ForeColor = m_ForeColor ' vbWindowText
    DisplayCaret m_ReadOnly
    'DrawTextFocusRect TxtInside
End Sub

Private Sub TxtInside_Validate(Cancel As Boolean)
    Dim OldVal As Double
    
    TxtInside.Text = Val(TxtInside.Text)
    Select Case Val(TxtInside.Text)
        Case Is > m_Max
            TxtInside.Text = m_Max
        Case Is < m_Min
            TxtInside.Text = m_Min
        Case Else
            TxtInside.Text = m_Increment * CLng(Val(TxtInside.Text) / m_Increment)
    End Select
    OldVal = m_Value
    
    m_Value = Val(TxtInside.Text)
    If m_UpDown.Value <> CUDVal(Val(TxtInside.Text)) Then
        m_UpDown.Value = CUDVal(Val(TxtInside.Text))
    End If
    TxtInside.Text = Format(m_Value, "#,##0" & IIf(m_DecimalPlaces > 0, "." & String(m_DecimalPlaces, "0"), ""))
    RaiseEvent Change(OldVal, m_Value)
End Sub

Private Sub m_UpDown_Change()
    If CUDVal(m_Value) <> m_UpDown.Value Then
        Select Case mDirection
            Case 0 'Down
                GoDown
            Case 1 'Up
                GoUP
            Case Else
                'Not defined
        End Select
    Else
        If Not m_UpDown.Wrap Then
            If (m_UpDown.Value = m_UpDown.Max) And CUDVal(Val(TxtInside.Text)) = m_UpDown.Max Then
                If mDirection = 1 Then
                    Beep 'Already at the end
                End If
            End If
            If (m_UpDown.Value = m_UpDown.Min) And CUDVal(Val(TxtInside.Text)) = m_UpDown.Min Then
                If mDirection = 0 Then
                    Beep 'Already at the begining
                End If
            End If
        End If
    End If
End Sub

Private Sub m_UpDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mDirection = m_Direction
End Sub

Private Sub m_UpDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim UpperSection As Single
    If m_UpDown.Orientation = cc2OrientationVertical Then
        UpperSection = m_UpDown.Height / 2
        m_Direction = Abs(CInt(Y < UpperSection))
    Else
        UpperSection = m_UpDown.Width / 2
        m_Direction = Abs(CInt(X > UpperSection))
    End If
End Sub

Private Sub m_UpDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'I had to do this because UpDown change event fires
    'multiple times and sometimes not when its suposed to
    mDirection = -1
End Sub

Private Sub UserControl_Initialize()
    Set m_UpDown = udVertical
    udVertical.Visible = True
    udVertical.Width = 255
    SetNumber TxtInside, True
End Sub

Private Sub UserControl_InitProperties()
    m_UseArrowKeys = False
    m_ReadOnly = False
    m_Max = 10
    m_Min = 0
    m_Increment = 1
    m_Value = 0
    m_Enabled = True
    m_BackColor = vbWindowBackground
    Orientation = jroVertical
    m_ShowFocusRect = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UpDownAlignment = .ReadProperty("UpDownAlignment", jrRighty)
        ShowFocusRect = .ReadProperty("ShowFocusRect", True)
        Orientation = .ReadProperty("Orientation", jroVertical)
        UseArrowKeys = .ReadProperty("UseArrowKeys", False)
        ReadOnly = .ReadProperty("ReadOnly", False)
        Increment = .ReadProperty("Increment", 1)
        DecimalPlaces = .ReadProperty("DecimalPlaces", 0)
        Max = .ReadProperty("Max", 10)
        Min = .ReadProperty("Min", 0)
        Value = .ReadProperty("Value", Min)
        Enabled = .ReadProperty("Enabled", True)
        BackColor = .ReadProperty("BackColor", vbWindowBackground)
        ForeColor = .ReadProperty("ForeColor", vbWindowText)
        Font = .ReadProperty("Font", Ambient.Font)
        Alignment = .ReadProperty("Alignment", vbRightJustify)
        Wrap = .ReadProperty("Wrap", False)
    End With
End Sub

Private Sub UserControl_Show()
    'I had to do this to force the updown control
    'recalculate its right size
    If m_Orientation = jroHorizontal Then
        udHorizontal.Visible = False
        udHorizontal.Visible = True
    Else
        udVertical.Visible = False
        udVertical.Visible = True
    End If
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "UpDownAlignment", m_UpDownAlignment, jrRighty
        .WriteProperty "ShowFocusRect", m_ShowFocusRect, True
        .WriteProperty "Orientation", m_Orientation, jroVertical
        .WriteProperty "UseArrowKeys", m_UseArrowKeys, False
        .WriteProperty "ReadOnly", m_ReadOnly, False
        .WriteProperty "Increment", m_Increment, 1
        .WriteProperty "DecimalPlaces", m_DecimalPlaces, 0
        .WriteProperty "Max", m_Max, 10
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Value", m_Value, m_Min
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "BackColor", m_BackColor, vbWindowBackground
        .WriteProperty "ForeColor", TxtInside.ForeColor, vbWindowText
        .WriteProperty "Font", TxtInside.Font, Ambient.Font
        .WriteProperty "Alignment", TxtInside.Alignment, vbRightJustify
        .WriteProperty "Wrap", m_Wrap, False
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If UserControl.Width < (m_UpDown.Width + 12 * Screen.TwipsPerPixelX) Then
        UserControl.Width = m_UpDown.Width + 12 * Screen.TwipsPerPixelX
    End If
    'Position Body
    TxtBody.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    If m_UpDownAlignment = jrRighty Then
        If Not IsUsingXPTheme Then
            'for clasic style or non XP Systems
            
            'Position EditBox
            TxtInside.Move 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, TxtBody.Width - m_UpDown.Width - 6 * Screen.TwipsPerPixelX, _
            TxtBody.Height - 6 * Screen.TwipsPerPixelY
            
            'Position UpDown
            m_UpDown.Move TxtInside.Left + TxtInside.Width + Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY
            m_UpDown.Height = TxtBody.Height - 4 * Screen.TwipsPerPixelY
        Else
            'for XP style
            
            'Position EditBox
            TxtInside.Move 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, TxtBody.Width - m_UpDown.Width - 5 * Screen.TwipsPerPixelX, _
            TxtBody.Height - 6 * Screen.TwipsPerPixelY
            
            'Position UpDown
            m_UpDown.Move TxtInside.Left + TxtInside.Width + Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
            m_UpDown.Height = TxtBody.Height - 2 * Screen.TwipsPerPixelY
        End If
    Else
        If Not IsUsingXPTheme Then
            'for clasic style or non XP Systems
            'Position EditBox
            TxtInside.Move m_UpDown + m_UpDown.Width + 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, TxtBody.Width - m_UpDown.Width - 6 * Screen.TwipsPerPixelX, _
            TxtBody.Height - 6 * Screen.TwipsPerPixelY
            
            'Position UpDown
            m_UpDown.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY
            m_UpDown.Height = TxtBody.Height - 4 * Screen.TwipsPerPixelY
        Else
            'for XP style
            'Position EditBox
            TxtInside.Move m_UpDown + m_UpDown.Width + 2 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, TxtBody.Width - m_UpDown.Width - 5 * Screen.TwipsPerPixelX, _
            TxtBody.Height - 6 * Screen.TwipsPerPixelY
            
            'Position UpDown
            m_UpDown.Move 1 * Screen.TwipsPerPixelX, Screen.TwipsPerPixelY
            m_UpDown.Height = TxtBody.Height - 2 * Screen.TwipsPerPixelY
        End If
    End If
    
    If TxtBody.Height > UserControl.Height Then
        UserControl.Height = TxtBody.Height
    End If
End Sub

Private Sub GoUP()
    Dim NextValue As Double
    NextValue = m_Value + m_Increment
    If (NextValue > m_Max) Then
        If m_Wrap Then
            NextValue = m_Min
        Else
            NextValue = m_Max
            Beep
        End If
    End If
    Value = NextValue
    If Not m_ReadOnly Then
        SelectAll
    End If
End Sub

Private Sub GoDown()
    Dim NextValue As Double
    NextValue = m_Value - m_Increment
    If (NextValue < m_Min) Then
        If m_Wrap Then
            NextValue = m_Max
        Else
            NextValue = m_Min
            Beep
        End If
    End If
    Value = NextValue
    If Not m_ReadOnly Then
        SelectAll
    End If
End Sub

Private Function CPos(Value)
    If Value < 0 Then
        CPos = 0
    Else
        CPos = Value
    End If
End Function

Public Sub SelectAll()
    TxtInside.SelStart = 0
    TxtInside.SelLength = Len(TxtInside.Text)
End Sub

Private Sub UpdateUpDownValues()
    If m_Increment = 0 Then Exit Sub
    m_UpDown.Max = CLng(Abs(m_Max - m_Min) / m_Increment)
    m_UpDown.Min = 0
    m_UpDown.Increment = 1
End Sub

Private Function CUDVal(Value As Double) As Long
    CUDVal = CLng((Value - m_Min) / m_Increment)
End Function

Public Function SetRange(dMax As Double, dMin As Double, dIncrement As Double, Optional dValue As Double) As Boolean
    If (dIncrement <= 0) Or (dMax <= dMin) Then
        Exit Function
    End If
    m_Increment = dIncrement
    m_Max = dMax
    m_Min = dMin
    UpdateUpDownValues
    Value = dValue
    SetRange = True
End Function

Private Sub DisplayCaret(bHide As Boolean)
    If Not bHide Then
        ShowCaret TxtInside.hWnd
    Else
        HideCaret TxtInside.hWnd
    End If
End Sub

Private Sub DrawTextFocusRect(objTextBox As TextBox, Optional lOffset As Long = 0)
    Dim lhDC As Long
    Dim Rec As RECT
    
    'Clear Text
    objTextBox.Refresh
    If m_ShowFocusRect Then
        'Draw focus rect
        SetRect Rec, lOffset, lOffset, (objTextBox.Width \ Screen.TwipsPerPixelX) - lOffset, (objTextBox.Height \ Screen.TwipsPerPixelY) - lOffset
        lhDC = GetWindowDC(objTextBox.hWnd)
        DrawFocusRect lhDC, Rec
    End If
End Sub

Private Function IsUsingXPTheme() As Boolean
    'Detects if you are using win xp and
    'if you have selected a XP theme
    Dim lhTheme As Long
    If IsXP Then
        On Error Resume Next
        lhTheme = OpenThemeData(UserControl.hWnd, StrPtr("Spin"))
        If lhTheme <> 0 Then
            IsUsingXPTheme = True
        End If
    End If
End Function

Private Function IsXP() As Boolean
    Dim lRet As Long
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = 148
    lRet = GetVersionEx(osvi)

    If lRet <> 0 Then
        If (osvi.dwPlatformId = 2) _
        And (osvi.dwMajorVersion = 5) _
        And (osvi.dwMinorVersion >= 1) _
        Then
            IsXP = True
        End If
    End If
End Function
