VERSION 5.00
Begin VB.Form frm_Page_Setup 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2340
      TabIndex        =   13
      Top             =   1770
      Width           =   675
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   345
      Left            =   150
      TabIndex        =   12
      Top             =   1770
      Width           =   735
   End
   Begin VB.Frame fra_Page_Setup 
      Caption         =   "Page Setup:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3105
      Begin VB.ListBox lst_PS_Top 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "frm_Page_Setup.frx":0000
         Left            =   510
         List            =   "frm_Page_Setup.frx":052F
         TabIndex        =   6
         Top             =   360
         Width           =   795
      End
      Begin VB.ListBox lst_PS_Left 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "frm_Page_Setup.frx":12FE
         Left            =   510
         List            =   "frm_Page_Setup.frx":1830
         TabIndex        =   5
         Top             =   840
         Width           =   795
      End
      Begin VB.ListBox lst_PS_Bottom 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "frm_Page_Setup.frx":2607
         Left            =   2250
         List            =   "frm_Page_Setup.frx":2B36
         TabIndex        =   4
         Top             =   360
         Width           =   795
      End
      Begin VB.ListBox lst_PS_Right 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "frm_Page_Setup.frx":3905
         Left            =   2250
         List            =   "frm_Page_Setup.frx":3E37
         TabIndex        =   3
         Top             =   840
         Width           =   795
      End
      Begin VB.CheckBox chk_Tbl_Border 
         Caption         =   "Table Border"
         Height          =   255
         Left            =   1740
         TabIndex        =   2
         Top             =   1320
         Width           =   1245
      End
      Begin VB.ListBox lst_PS_Align 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "frm_Page_Setup.frx":4C0E
         Left            =   510
         List            =   "frm_Page_Setup.frx":4C30
         TabIndex        =   1
         Top             =   1290
         Width           =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         X1              =   1650
         X2              =   1650
         Y1              =   100
         Y2              =   1680
      End
      Begin VB.Label lbl_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Top"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   285
      End
      Begin VB.Label lbl_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Bottom"
         Height          =   195
         Index           =   1
         Left            =   1710
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Left"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   840
         Width           =   270
      End
      Begin VB.Label lbl_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Right"
         Height          =   195
         Index           =   3
         Left            =   1710
         TabIndex        =   8
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lbl_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Align"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   7
         Top             =   1320
         Width           =   345
      End
   End
End
Attribute VB_Name = "frm_Page_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''''
Private Sub cmd_Cancel_Click()
 Unload Me
End Sub

Private Sub cmd_OK_Click()
 Top_Margin = lst_PS_Top.Text
 Bottom_Margin = lst_PS_Bottom.Text
 Left_Margin = lst_PS_Left.Text
 Right_Margin = lst_PS_Right.Text
 Table_Border = chk_Tbl_Border.Value
 Paragraph_Alignment = lst_PS_Align.ListIndex
 Unload Me
End Sub


Private Sub Form_Activate()
 lst_PS_Top.Text = Top_Margin
 lst_PS_Bottom.Text = Bottom_Margin
 lst_PS_Right.Text = Right_Margin
 lst_PS_Left.Text = Left_Margin
 chk_Tbl_Border.Value = Table_Border
 lst_PS_Align.Text = Paragraph_Alignment
End Sub

Private Sub Form_Load()
 Me.Caption = EM_TITLE & "..Page Setup" & EM_MAIL
End Sub
