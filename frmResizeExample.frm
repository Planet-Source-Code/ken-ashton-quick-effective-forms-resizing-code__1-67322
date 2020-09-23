VERSION 5.00
Begin VB.Form frmResizeExample 
   Caption         =   " Resize Example"
   ClientHeight    =   3735
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3450
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   210
      Width           =   5205
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   6570
      TabIndex        =   7
      Top             =   690
      Width           =   2145
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   6600
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   3450
      TabIndex        =   5
      Top             =   750
      Width           =   2895
      Begin VB.CheckBox chkSetting 
         Caption         =   "Setting 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1290
         Width           =   1005
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Setting 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1005
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Setting 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   630
         Width           =   975
      End
      Begin VB.CheckBox chkSetting 
         Caption         =   "Setting 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4170
      TabIndex        =   4
      Top             =   2760
      Width           =   1065
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   540
      TabIndex        =   2
      Top             =   570
      Width           =   1725
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   540
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   0
      Top             =   1680
      Width           =   2355
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2820
      TabIndex        =   9
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   540
      TabIndex        =   3
      Top             =   990
      Width           =   1125
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmResizeExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' The only two things here to look at are in the Form_Load and Form_Resize events
'
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'
'  Place the two lines below in the Form_Load event of
'  the first form that gets loaded in your project.
'
'  Important: Assumes Form.ScaleMode = 1 (Twips) NOT  Form.ScaleMode = 3 (Pixels)
'
    myX = Me.TextWidth("W") * 0.2        ' Define the 'Horizontal' border constant
    myY = Me.TextHeight("W") * 0.2       ' Define the 'Vertical' border constant
'
'  Thats it. myX and myY are accessible from anywhere in the project (if required)
'  They reflect VB's interpretation of sizes at startup.
'
'  Go to the example Form_Resize event to see how we use these two values
'
End Sub


Private Sub Form_Resize()
    Dim x!, y!, h!, w!, i!
'
'   Limitations: Usage is for applications where the default VB font sizes and ScaleModes are used.
'   More sophisticated applications which tailor those properties will obviously also require some
'   appropriate tailoring of the myX and myY values, which are assigned in the Form_Load event.
'
'
'   Place any desired 'limits' here
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < Screen.Width * 0.22 Then Me.Width = Screen.Width * 0.22
    If Me.Height < Screen.Height * 0.3 Then Me.Height = Screen.Height * 0.3
'
'   Init important local vars
'
    x = myX                                         ' Grab our Global border constants defined in the
    y = myY                                         ' Form_Load event and place into local variables.
    w = Me.Width                                    ' Get current form Width
    h = Me.Height                                   ' Get current form Height
'
'   When desirable, define other vars for very common values, like 'left margin' when many controls
'
    Dim x2!, x6!                                    ' Dim some vars for commonly used values
    x2 = x * 2                                      ' in this example, times 2, and
    x6 = x * 6                                      ' also 6 times the constant.
'
'   Move controls using multiples of x and y for horizontal and vertically dependent
'   static spacing, and fractions of w, h for overall control sizing and positioning.
'
    Check1.Move x6, y * 6                           ' Pretty much the 'left column' in this example
    Check2.Move x6, y * 16
    Label1.Move x2, y * 28, w * 0.3 - x6
    List1.Move x2, y * 38, w * 0.3 - x6, h - y * 57

    Text1.Move w * 0.3, y * 5, w * 0.7 - x6         ' The 'middle' column of the display
    Label2.Move Text1.Left - Label2.Width - x * 2, Text1.Top + y
    Frame1.Move w * 0.3, y * 18, w * 0.4 - x * 4, h - y * 56
    For i = 0 To 3
        chkSetting(i).Move Frame1.Width * 0.1, (i + 1) * Frame1.Height * 0.21 - y * 2
    Next
    Command1.Move w * 0.5 - x * 18, h - y * 32
'
'   Now do what is effectively the 'right column' of the display.
'   Notice we cannot turn off the 'IntegralHeght' property or the Dir/File listboxes
'   so we will just 'put up with' the 'bottom' margin variations with form height
'
    Dir1.Move w * 0.7, y * 20, w * 0.3 - x6, h * 0.4
    File1.Move w * 0.7, h * 0.4 + y * 22, w * 0.3 - x6, h * 0.6 - y * 40
    
End Sub


Private Sub mnExit_Click()
Unload Me
End Sub


