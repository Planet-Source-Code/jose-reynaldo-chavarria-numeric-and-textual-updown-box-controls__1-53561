VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.UserControl JRTextUpDown 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   PropertyPages   =   "JRTextUpDown.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   1245
   ToolboxBitmap   =   "JRTextUpDown.ctx":0030
   Begin VB.Timer TmrSearch 
      Interval        =   900
      Left            =   1320
      Top             =   960
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
   Begin ComCtl2.UpDown udHorizontal 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   327681
      Max             =   -8
      Orientation     =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox TxtInside 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox TxtBody 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label LblName 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Text UpDown"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "JRTextUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'////////////////////////////////////////////////////////
'///             Textual Up Down Control
'///               (JRTextUpDown.ctl)
'///_____________________________________________________
'/// Textual updown box.
'///_____________________________________________________
'/// Last modification  : Apr/26/2004
'/// Last modified by   : Jose Reynaldo Chavarria
'/// Modification reason: Creation
'/// Author: Jose Reynaldo Chavarria Q. (jchavarria@agrolibano.hn)
'/// Tested on: Windows XP SP1, Windows 98, Windows NT 4.0 SP6.0
'/// Agrolibano S.A. de C.V.
'/// Web Site: http://www.agrolibano.hn/
'////////////////////////////////////////////////////////

Dim mDirection As Integer, m_Direction As Integer '0=Down, 1=Up
Dim m_List As String
Dim m_aList() As String
Dim m_Value As Long
Dim m_Wrap As Boolean
Dim m_ReadOnly As Boolean
Dim m_UseArrowKeys As Boolean
Dim m_Enabled As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_ForceToList As Boolean
Dim m_ReverseOrder As Boolean
Dim m_Orientation As jrOrientationConstants
Dim m_ShowFocusRect As Boolean
Dim m_UpDownAlignment As jrUDAlignConstants
Dim WithEvents m_UpDown As UpDown
Attribute m_UpDown.VB_VarHelpID = -1
Dim m_CharBuffer As String

Public Enum jrOrientationConstants
    jroVertical = 0
    jroHorizontal = 1
End Enum
Public Enum jrUDAlignConstants
    jrRighty = 0
    jrLefty = 1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

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

Public Event Change(ByVal PreviousIndex As Long, ByVal NewIndex As Long)

'///////////////////////////////////////////////
'///Properties
'///
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

Public Property Let ItemData(ByVal Index As Long, ByVal NewValue As String)
    Dim sItem As String
    
    If m_ReverseOrder Then
        Index = UBound(m_aList) - Index
    End If
    'Get Item Text
    sItem = Split(m_aList(Index), vbTab)(0)
    
    'Update Item Data
    m_aList(Index) = sItem & vbTab & NewValue
    
    'Reload List
    List = Join(m_aList, vbCr)
End Property
Public Property Get ItemData(ByVal Index As Long) As String
    Dim sItem As String
    
    If m_ReverseOrder Then
        Index = UBound(m_aList) - Index
    End If
    
    'Get Whole Item Text
    sItem = m_aList(Index)
    If InStr(1, sItem, vbTab) > 0 Then
        'Get Item Data
        ItemData = Split(m_aList(Index), vbTab)(1)
    Else
        ItemData = ""
    End If
End Property

Public Property Let Orientation(NewValue As jrOrientationConstants)
    m_Orientation = NewValue
    If NewValue = jroVertical Then
        Set m_UpDown = udVertical
        udVertical.Max = udHorizontal.Max
        udVertical.Min = udHorizontal.Min
        udVertical.Wrap = udHorizontal.Wrap
        udVertical.Enabled = udHorizontal.Enabled
        UpdateValues
        UserControl_Resize
        udHorizontal.Visible = False
        udVertical.Visible = True
    Else
        Set m_UpDown = udHorizontal
        udHorizontal.Max = udVertical.Max
        udHorizontal.Min = udVertical.Min
        udHorizontal.Enabled = udVertical.Enabled
        UpdateValues
        UserControl_Resize
        udVertical.Visible = False
        udHorizontal.Visible = True
    End If
    PropertyChanged "Orientation"
End Property
Public Property Get Orientation() As jrOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let ReverseOrder(NewValue As Boolean)
    m_ReverseOrder = NewValue
    PropertyChanged "ReverseOrder"
    UpdateValues
End Property
Public Property Get ReverseOrder() As Boolean
Attribute ReverseOrder.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ReverseOrder = m_ReverseOrder
End Property

Public Property Let Text(NewValue As String)
    TxtInside.Text = NewValue
    PropertyChanged "Text"
End Property
Public Property Get Text() As String
    Text = TxtInside.Text
End Property

Public Property Let ListIndex(NewValue As Long)
    Dim OldVal As Long
    
    If UBound(m_aList) = -1 Then Exit Property
    Select Case NewValue
        Case Is > m_UpDown.Max
            NewValue = m_UpDown.Max
        Case Is < m_UpDown.Min
            NewValue = m_UpDown.Min
    End Select
    
    OldVal = m_Value
    m_Value = NewValue
    
    If m_UpDown.Value <> NewValue Then
        m_UpDown.Value = NewValue
    End If
    If m_ReverseOrder Then
        TxtInside.Text = Split(m_aList(UBound(m_aList) - m_Value), vbTab)(0)
    Else
        TxtInside.Text = Split(m_aList(m_Value), vbTab)(0)
    End If
    TxtInside.DataChanged = False
    
    RaiseEvent Change(OldVal, NewValue)
    PropertyChanged "ListIndex"
End Property
Public Property Get ListIndex() As Long
    If m_ReverseOrder Then
        ListIndex = UBound(m_aList) - m_Value
    Else
        ListIndex = m_Value
    End If
End Property

Public Property Get ListCount() As Long
    ListCount = UBound(m_aList) + 1
End Property

Public Property Get Alignment() As AlignmentConstants
    Alignment = TxtInside.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    TxtInside.Alignment = New_Alignment
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

Public Property Let UseArrowKeys(NewValue As Boolean)
    m_UseArrowKeys = NewValue
    PropertyChanged "UseArrowKeys"
End Property
Public Property Get UseArrowKeys() As Boolean
    UseArrowKeys = m_UseArrowKeys
End Property

Public Property Let ForceToList(NewValue As Boolean)
    m_ForceToList = NewValue
    PropertyChanged "ForceToList"
End Property
Public Property Get ForceToList() As Boolean
    ForceToList = m_ForceToList
End Property

Public Property Let ReadOnly(NewValue As Boolean)
    m_ReadOnly = NewValue
    If NewValue Then
        TxtInside.MousePointer = vbArrow
    Else
        TxtInside.MousePointer = vbIbeam
    End If
    TxtInside.Locked = NewValue
    DisplayCaret NewValue
    PropertyChanged "ReadOnly"
End Property
Public Property Get ReadOnly() As Boolean
    ReadOnly = m_ReadOnly
    TxtInside.Locked = m_ReadOnly
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

Public Property Let Enabled(NewValue As Boolean)
    m_Enabled = NewValue
    TxtInside.Enabled = NewValue
    m_UpDown.Enabled = NewValue
    PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Wrap(NewValue As Boolean)
    m_Wrap = NewValue
    m_UpDown.Wrap = NewValue
    PropertyChanged "Wrap"
End Property
Public Property Get Wrap() As Boolean
    Wrap = m_UpDown.Wrap
End Property

Public Property Let List(NewValue As String)
    m_List = NewValue
    m_aList = Split(m_List, vbCr)
    UpdateValues
    PropertyChanged "List"
End Property
Public Property Get List() As String
Attribute List.VB_ProcData.VB_Invoke_Property = "ppTextList"
Attribute List.VB_MemberFlags = "400"
    List = m_List
End Property
'                                         \\\
'                         End Of Properties\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Sub AddItem(sItem As String, Optional ItemData As String = "")
    Dim CurList As String
    CurList = m_List
    
    If CurList <> "" Then
        CurList = CurList & vbCr & sItem & vbTab & ItemData
    Else
        CurList = sItem & vbTab & ItemData
    End If
    
    List = CurList
End Sub

Public Sub RemoveItem(Item As Long)
    Dim tList() As String
    Dim i As Long
    Dim tResult As String
    
    If m_ReverseOrder Then
        Item = UBound(m_aList) - Item
    End If

    tList = Split(m_List, vbCr)
    For i = 0 To UBound(tList)
        If i <> Item Then
            If tResult <> "" Then
                tResult = tResult & vbCr & tList(i)
            Else
                tResult = tList(i)
            End If
        End If
    Next
    List = tResult
End Sub

Private Sub TmrSearch_Timer()
    m_CharBuffer = ""
    TmrSearch.Enabled = False
End Sub

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

Private Sub TxtInside_GotFocus()
    If m_ReadOnly Then
        TxtInside.BackColor = vbHighlight
        TxtInside.ForeColor = vbHighlightText
        DrawTextFocusRect TxtInside
    Else
        SelectAll
    End If
    DisplayCaret m_ReadOnly
End Sub

Private Sub TxtInside_LostFocus()
    TxtInside.BackColor = m_BackColor 'vbWindowBackground
    TxtInside.ForeColor = m_ForeColor ' vbWindowText
    DisplayCaret m_ReadOnly
    'DrawTextFocusRect TxtInside
End Sub

Private Sub TxtInside_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    Select Case KeyCode
    Case vbKeyUp, vbKeyPageUp

        If m_UseArrowKeys And (m_Orientation = jroVertical) Then
            If UBound(m_aList) = -1 Then
                Beep
                Exit Sub
            End If
            If m_ReverseOrder Then
                If TxtInside.Text <> Split(m_aList(UBound(m_aList) - m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            Else
                If TxtInside.Text <> Split(m_aList(m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            End If
            GoUP
        End If
        
    Case vbKeyDown, vbKeyPageDown
        If m_UseArrowKeys And (m_Orientation = jroVertical) Then
            If UBound(m_aList) = -1 Then
                Beep
                Exit Sub
            End If
            If m_ReverseOrder Then
                If TxtInside.Text <> Split(m_aList(UBound(m_aList) - m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            Else
                If TxtInside.Text <> Split(m_aList(m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            End If
            GoDown
        End If
        
    Case vbKeyLeft
        If m_UseArrowKeys And (m_Orientation = jroHorizontal) Then
            If UBound(m_aList) = -1 Then
                Beep
                Exit Sub
            End If
            If m_ReverseOrder Then
                If TxtInside.Text <> Split(m_aList(UBound(m_aList) - m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            Else
                If TxtInside.Text <> Split(m_aList(m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            End If
            GoDown
        End If
    Case vbKeyRight
    
        If m_UseArrowKeys And (m_Orientation = jroHorizontal) Then
            If UBound(m_aList) = -1 Then
                Beep
                Exit Sub
            End If
            If m_ReverseOrder Then
                If TxtInside.Text <> Split(m_aList(UBound(m_aList) - m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            Else
                If TxtInside.Text <> Split(m_aList(m_UpDown.Value), vbTab)(0) Then
                    TxtInside_Validate False
                End If
            End If
            GoUP
        End If
    
    Case vbKeyA To vbKeyZ, vbKey0 To vbKey9, vbKeySpace
        If m_ForceToList And m_ReadOnly Then
            m_CharBuffer = m_CharBuffer & Chr(KeyCode)
            Debug.Print "look for '" & m_CharBuffer & "'"
            For i = 0 To UBound(m_aList)
                If Left(LCase(Split(m_aList(i), vbTab)(0)), Len(m_CharBuffer)) = LCase(m_CharBuffer) And i <> ListIndex Then
                    If m_ReverseOrder Then
                        ListIndex = UBound(m_aList) - i
                    Else
                        ListIndex = i
                    End If
                End If
            Next
            TmrSearch.Enabled = True
        End If
        
    End Select
End Sub

Private Sub TxtInside_Validate(Cancel As Boolean)
    Dim i As Long
    'prevent for checking existance when read only
    If m_ReadOnly Then Exit Sub
    If Not TxtInside.DataChanged Then Exit Sub
    
    If m_ForceToList Then
        For i = 0 To UBound(m_aList)
            If m_ReverseOrder Then
                If Len(Split(m_aList(UBound(m_aList) - i), vbTab)(0)) > Len(TxtInside.Text) Then
                    If LCase(Left(Split(m_aList(UBound(m_aList) - i), vbTab)(0), Len(TxtInside.Text))) = LCase(TxtInside.Text) Then
                        ListIndex = i
                        Exit Sub
                    End If
                End If
            Else
                If Len(Split(m_aList(i), vbTab)(0)) > Len(TxtInside.Text) Then
                    If LCase(Left(Split(m_aList(i), vbTab)(0), Len(TxtInside.Text))) = LCase(TxtInside.Text) Then
                        ListIndex = i
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        'not found
    
        If UBound(m_aList) = -1 Then
            TxtInside.Text = ""
            TxtInside.DataChanged = False
        Else
            ListIndex = LBound(m_aList)
        End If
    End If
    
    'leave as it is
End Sub

Private Sub UserControl_Initialize()
    Set m_UpDown = udVertical
    udVertical.Visible = True
    udVertical.Width = 255
End Sub

Private Sub UserControl_InitProperties()
    m_ReverseOrder = False
    m_UseArrowKeys = False
    m_ReadOnly = False
    m_Enabled = True
    m_BackColor = vbWindowBackground
    m_ForeColor = vbWindowText
    m_Wrap = False
    m_ForceToList = False
    Orientation = jroVertical
    m_ShowFocusRect = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UpDownAlignment = .ReadProperty("UpDownAlignment", jrRighty)
        ShowFocusRect = .ReadProperty("ShowFocusRect", True)
        Orientation = .ReadProperty("Orientation", jroVertical)
        ReverseOrder = .ReadProperty("ReverseOrder", False)
        List = .ReadProperty("List", "")
        Text = .ReadProperty("Text", "")
        UseArrowKeys = .ReadProperty("UseArrowKeys", False)
        ReadOnly = .ReadProperty("ReadOnly", False)
        Enabled = .ReadProperty("Enabled", True)
        BackColor = .ReadProperty("BackColor", vbWindowBackground)
        ForeColor = .ReadProperty("ForeColor", vbWindowText)
        Font = .ReadProperty("Font", Ambient.Font)
        Alignment = .ReadProperty("Alignment", vbRightJustify)
        Wrap = .ReadProperty("Wrap", False)
        ForceToList = .ReadProperty("ForceToList", False)
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
        .WriteProperty "ReverseOrder", m_ReverseOrder, False
        .WriteProperty "List", m_List, ""
        .WriteProperty "Text", TxtInside.Text, ""
        .WriteProperty "UseArrowKeys", m_UseArrowKeys, False
        .WriteProperty "ReadOnly", m_ReadOnly, False
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "BackColor", m_BackColor, vbWindowBackground
        .WriteProperty "ForeColor", TxtInside.ForeColor, vbWindowText
        .WriteProperty "Font", TxtInside.Font, Ambient.Font
        .WriteProperty "Alignment", TxtInside.Alignment, vbRightJustify
        .WriteProperty "Wrap", m_Wrap, False
        .WriteProperty "ForceToList", m_ForceToList, False
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

Private Sub UpdateValues()
    
    If m_List = "" Then 'is empty
        m_UpDown.Min = 0
        m_UpDown.Max = 0
        TxtInside.Enabled = False
        TxtInside.Text = ""
        TxtInside.DataChanged = False
    Else
        m_UpDown.Min = LBound(m_aList)
        m_UpDown.Max = UBound(m_aList)
        TxtInside.Enabled = Not (LBound(m_aList) = 0 And UBound(m_aList) = 0)
        If m_Value <= UBound(m_aList) Then
            ListIndex = m_Value
        Else
            ListIndex = UBound(m_aList)
        End If
    End If
    
    m_UpDown.Increment = 1
End Sub

Private Sub m_UpDown_Change()
    If m_Value <> m_UpDown.Value Then
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
            If (m_UpDown.Value = m_UpDown.Max) And m_Value = m_UpDown.Max Then
                If mDirection = 1 Then
                    Beep 'Already at the end
                End If
            End If
            If (m_UpDown.Value = m_UpDown.Min) And m_Value = m_UpDown.Min Then
                If mDirection = 0 Then
                    Beep 'Already at the begining
                End If
            End If
        End If
    End If
    If m_ReadOnly Then
        DrawTextFocusRect TxtInside
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

Private Sub GoUP()
    Dim NextValue As Double
    NextValue = m_Value + m_UpDown.Increment
    If (NextValue > m_UpDown.Max) Then
        If m_UpDown.Wrap Then
            NextValue = m_UpDown.Min
        Else
            NextValue = m_UpDown.Max
            Beep
        End If
    End If
    ListIndex = NextValue
    If Not m_ReadOnly Then
        SelectAll
    End If
End Sub

Private Sub GoDown()
    Dim NextValue As Double
    NextValue = m_Value - m_UpDown.Increment
    If (NextValue < m_UpDown.Min) Then
        If m_UpDown.Wrap Then
            NextValue = m_UpDown.Max
        Else
            NextValue = m_UpDown.Min
            Beep
        End If
    End If
    ListIndex = NextValue
    If Not m_ReadOnly Then
        SelectAll
    End If
End Sub

Public Sub SelectAll()
    TxtInside.SelStart = 0
    TxtInside.SelLength = Len(TxtInside.Text)
End Sub

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


