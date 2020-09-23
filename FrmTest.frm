VERSION 5.00
Begin VB.Form FrmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test JRUpdown"
   ClientHeight    =   6045
   ClientLeft      =   1185
   ClientTop       =   660
   ClientWidth     =   5925
   Icon            =   "FrmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5925
   Begin UpDownTest.JRTextUpDown tud 
      Height          =   300
      Left            =   240
      TabIndex        =   54
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      List            =   $"FrmTest.frx":1042
      Text            =   "San Marcos de Colon"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   0
   End
   Begin UpDownTest.JRNumericUpDown nud 
      Height          =   300
      Left            =   480
      TabIndex        =   52
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Increment       =   0.1
      DecimalPlaces   =   1
      Max             =   5
      Min             =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "Numeric Range"
      Height          =   1215
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   1935
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   930
         Left            =   120
         ScaleHeight     =   930
         ScaleWidth      =   1575
         TabIndex        =   31
         Top             =   240
         Width           =   1575
         Begin VB.CommandButton CmdUpdate 
            Caption         =   "Set"
            Height          =   285
            Left            =   960
            TabIndex        =   35
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox TxtIncrement 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            TabIndex        =   34
            Text            =   "0.1"
            Top             =   570
            Width           =   495
         End
         Begin VB.TextBox TxtMin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            TabIndex        =   33
            Text            =   "0"
            Top             =   285
            Width           =   495
         End
         Begin VB.TextBox TxtMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            TabIndex        =   32
            Text            =   "5"
            Top             =   0
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Incr:"
            Height          =   195
            Left            =   0
            TabIndex        =   38
            Top             =   615
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min:"
            Height          =   195
            Left            =   0
            TabIndex        =   37
            Top             =   330
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max:"
            Height          =   195
            Left            =   0
            TabIndex        =   36
            Top             =   45
            Width           =   345
         End
      End
   End
   Begin VB.CommandButton CmdNudFont 
      Caption         =   "Font ..."
      Height          =   375
      Left            =   480
      TabIndex        =   25
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CmdTudFont 
      Caption         =   "Font ..."
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox TxtnudVal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox TxtItemData 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "0"
      Top             =   3600
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text UpDown Properties"
      Height          =   2895
      Left            =   2160
      TabIndex        =   7
      Top             =   3000
      Width           =   3615
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   120
         ScaleHeight     =   2535
         ScaleWidth      =   3375
         TabIndex        =   8
         Top             =   240
         Width           =   3375
         Begin VB.Frame Frame6 
            Caption         =   "Spin Alignment"
            Height          =   840
            Left            =   1560
            TabIndex        =   41
            Top             =   1560
            Width           =   1455
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   540
               Left            =   120
               ScaleHeight     =   540
               ScaleWidth      =   1215
               TabIndex        =   42
               Top             =   240
               Width           =   1215
               Begin VB.OptionButton OptTudAlign 
                  Caption         =   "Left"
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   44
                  Top             =   240
                  Width           =   855
               End
               Begin VB.OptionButton OptTudAlign 
                  Caption         =   "Right"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   43
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
            End
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Focus Rect"
            Height          =   195
            Left            =   0
            TabIndex        =   40
            ToolTipText     =   "When Read only"
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Force to List"
            Height          =   195
            Left            =   0
            TabIndex        =   16
            Top             =   1900
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Use Arrow keys"
            Height          =   195
            Left            =   0
            TabIndex        =   15
            Top             =   1120
            Width           =   1455
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Read Only"
            Height          =   195
            Left            =   0
            TabIndex        =   14
            Top             =   1380
            Width           =   1455
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Enabled"
            Height          =   195
            Left            =   0
            TabIndex        =   13
            Top             =   1640
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Wrap"
            Height          =   195
            Left            =   0
            TabIndex        =   12
            Top             =   860
            Width           =   1335
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Reverse Order"
            Height          =   195
            Left            =   0
            TabIndex        =   11
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Add Item"
            Height          =   375
            Left            =   1440
            TabIndex        =   10
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Remove Item"
            Height          =   375
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Caption         =   "Spin Orientation"
            Height          =   855
            Left            =   1560
            TabIndex        =   17
            Top             =   600
            Width           =   1455
            Begin VB.PictureBox Picture7 
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   120
               ScaleHeight     =   495
               ScaleWidth      =   1215
               TabIndex        =   49
               Top             =   240
               Width           =   1215
               Begin VB.OptionButton OptOrientation 
                  Caption         =   "Vertical"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   51
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton OptOrientation 
                  Caption         =   "Horizontal"
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   50
                  Top             =   240
                  Width           =   1095
               End
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Numeric UpDown Properties"
      Height          =   2655
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         ScaleHeight     =   2055
         ScaleWidth      =   3375
         TabIndex        =   1
         Top             =   240
         Width           =   3375
         Begin UpDownTest.JRNumericUpDown NUDDecPlac 
            Height          =   300
            Left            =   120
            TabIndex        =   53
            Top             =   1560
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            ReadOnly        =   -1  'True
            Max             =   4
            Min             =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame Frame7 
            Caption         =   "Spin Alignment"
            Height          =   840
            Left            =   1680
            TabIndex        =   45
            Top             =   960
            Width           =   1455
            Begin VB.PictureBox Picture6 
               BorderStyle     =   0  'None
               Height          =   540
               Left            =   120
               ScaleHeight     =   540
               ScaleWidth      =   975
               TabIndex        =   46
               Top             =   240
               Width           =   975
               Begin VB.OptionButton OptNudAlign 
                  Caption         =   "Right"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   48
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton OptNudAlign 
                  Caption         =   "Left"
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   47
                  Top             =   240
                  Width           =   855
               End
            End
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Focus Rect"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "When Read only"
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.Frame Frame4 
            Caption         =   "Spin Orientation"
            Height          =   855
            Left            =   1680
            TabIndex        =   26
            Top             =   0
            Width           =   1455
            Begin VB.PictureBox Picture4 
               BorderStyle     =   0  'None
               Height          =   480
               Left            =   60
               ScaleHeight     =   480
               ScaleWidth      =   1215
               TabIndex        =   27
               Top             =   195
               Width           =   1215
               Begin VB.OptionButton OptNudOrientation 
                  Caption         =   "Vertical"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   29
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton OptNudOrientation 
                  Caption         =   "Horizontal"
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1215
               End
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Enabled"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Read Only"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Use Arrow keys"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Wrap"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Decimal Places"
            Height          =   195
            Left            =   0
            TabIndex        =   6
            Top             =   1200
            Width           =   1095
         End
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Value:"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   735
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "ItemData:"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3975
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index:"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   3615
      Width           =   435
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5760
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   5750
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CmDlg As New cCommonDialog

Private Sub Check1_Click()
    nud.Wrap = CBool(Check1.Value)
End Sub

Private Sub Check10_Click()
    tud.ForceToList = CBool(Check10.Value)
End Sub

Private Sub Check11_Click()
    nud.ShowFocusRect = CBool(Check11.Value)
End Sub

Private Sub Check12_Click()
    tud.ShowFocusRect = CBool(Check12.Value)
End Sub

Private Sub Check2_Click()
    nud.UseArrowKeys = CBool(Check2.Value)
End Sub

Private Sub Check3_Click()
    nud.ReadOnly = CBool(Check3.Value)
End Sub

Private Sub Check4_Click()
    nud.Enabled = CBool(Check4.Value)
End Sub

Private Sub Check5_Click()
    tud.ReverseOrder = CBool(Check5.Value)
End Sub

Private Sub Check6_Click()
    tud.Wrap = CBool(Check6.Value)
End Sub

Private Sub Check7_Click()
    tud.Enabled = CBool(Check7.Value)
End Sub

Private Sub Check8_Click()
    tud.ReadOnly = CBool(Check8.Value)
End Sub

Private Sub Check9_Click()
    tud.UseArrowKeys = CBool(Check9.Value)
End Sub

Private Sub CmdNudFont_Click()
    Dim lColor As OLE_COLOR
    
    lColor = nud.ForeColor
    With CmDlg
        If .VBChooseFont(nud.Font, , Me.hWnd, lColor) Then
            nud.ForeColor = lColor
        End If
    End With
End Sub

Private Sub CmdTudFont_Click()
    Dim lColor As OLE_COLOR
    
    lColor = tud.ForeColor
    With CmDlg
        If .VBChooseFont(tud.Font, , Me.hWnd, lColor) Then
            tud.ForeColor = lColor
        End If
    End With
End Sub

Private Sub CmdUpdate_Click()
    If Not nud.SetRange(Val(TxtMax.Text), Val(TxtMin.Text), Val(TxtIncrement.Text)) Then
        MsgBox "Wrong Range!", vbCritical
    End If
End Sub

Private Sub Command1_Click()
     Me.tud.RemoveItem tud.ListCount - 1
End Sub

Private Sub Command2_Click()
    tud.AddItem "Newitem" & tud.ListCount, "key-" & tud.ListCount
End Sub

Private Sub OptNudAlign_Click(Index As Integer)
    nud.UpDownAlignment = Index
End Sub

Private Sub OptTudAlign_Click(Index As Integer)
    tud.UpDownAlignment = Index
End Sub

Private Sub tud_Change(ByVal PreviousIndex As Long, ByVal NewIndex As Long)
    Text1.Text = NewIndex
    TxtItemData.Text = tud.ItemData(NewIndex)
End Sub

Private Sub nud_Change(ByVal PreviousValue As Double, ByVal NewValue As Double)
    TxtnudVal.Text = NewValue
End Sub

Private Sub NUDDecPlac_Change(ByVal PrevousValue As Double, ByVal NewValue As Double)
    nud.DecimalPlaces = NUDDecPlac.Value
End Sub

Private Sub OptNudOrientation_Click(Index As Integer)
    nud.Orientation = Index
End Sub

Private Sub OptOrientation_Click(Index As Integer)
    tud.Orientation = Index
End Sub
