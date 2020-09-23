VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Address_Book 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7005
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   9600
   Icon            =   "frm_Address_Book.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frm_Address_Book.frx":382A
   ScaleHeight     =   7005
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Binary_Search 
      BackColor       =   &H00C0E0FF&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   420
      Width           =   2235
   End
   Begin VB.CheckBox chk_Name 
      Height          =   195
      Left            =   4080
      TabIndex        =   22
      Top             =   5220
      Width           =   195
   End
   Begin VB.CheckBox chk_Residence_Address 
      Height          =   195
      Left            =   6510
      TabIndex        =   21
      Top             =   5220
      Width           =   195
   End
   Begin VB.CheckBox chk_Office_Address 
      Height          =   195
      Left            =   6510
      TabIndex        =   20
      Top             =   5520
      Width           =   195
   End
   Begin VB.CheckBox chk_Residence_Phone 
      Height          =   195
      Left            =   5070
      TabIndex        =   19
      Top             =   5220
      Width           =   195
   End
   Begin VB.CheckBox chk_Office_Phone 
      Height          =   195
      Left            =   5070
      TabIndex        =   18
      Top             =   5520
      Width           =   195
   End
   Begin VB.TextBox txt_Search 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   17
      Top             =   5790
      Width           =   2205
   End
   Begin VB.CheckBox chk_Email 
      Height          =   195
      Left            =   4080
      TabIndex        =   16
      Top             =   5520
      Width           =   195
   End
   Begin VB.CheckBox chk_Search_All 
      Height          =   195
      Left            =   6720
      TabIndex        =   15
      Top             =   5880
      Width           =   195
   End
   Begin VB.CommandButton cmd_Search 
      Height          =   285
      Left            =   6330
      Picture         =   "frm_Address_Book.frx":E322
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Search"
      Top             =   5820
      Width           =   285
   End
   Begin VB.ListBox lst_Group 
      Height          =   3960
      ItemData        =   "frm_Address_Book.frx":E8AC
      Left            =   240
      List            =   "frm_Address_Book.frx":E8AE
      TabIndex        =   7
      Top             =   2010
      Width           =   735
   End
   Begin VB.ListBox lst_Address 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5340
      Left            =   1260
      TabIndex        =   2
      Top             =   840
      Width           =   2745
   End
   Begin MSComDlg.CommonDialog cdl_File 
      Left            =   2730
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".htm"
      DialogTitle     =   "Convert Data"
      Filter          =   "Doc File|*.doc"
   End
   Begin MSComctlLib.ImageList imgLst_Transparent_BG 
      Left            =   2130
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   638
      ImageHeight     =   469
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":E8B0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   3270
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   55
      ImageHeight     =   56
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   44
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":E9F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EA4B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EA969
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EACAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EAFB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EB2A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EB585
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EB92F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EBC95
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EC005
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EC355
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EC860
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":ECD95
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":ED2C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":ED84E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EDD9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EE324
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EE8BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EEC59
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EEFD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EF2EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EF61B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":EFAC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F00CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F05AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F08F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F0C09
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F0F11
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F120E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F1652
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F19E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F1D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F212D
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F2692
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F2C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F317E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F3717
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F3C97
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F422C
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F47A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F4B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F4F03
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F5251
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Address_Book.frx":F55B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image img_Labels 
      Height          =   480
      Index           =   5
      Left            =   4080
      Picture         =   "frm_Address_Book.frx":F5A5B
      ToolTipText     =   "Comments"
      Top             =   4350
      Width           =   480
   End
   Begin VB.Label lbl_Comments 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comments"
      Height          =   795
      Left            =   4800
      TabIndex        =   31
      ToolTipText     =   "Comments"
      Top             =   4230
      Width           =   2955
      WordWrap        =   -1  'True
   End
   Begin VB.Image img_AR 
      Height          =   420
      Index           =   21
      Left            =   360
      ToolTipText     =   "Group"
      Top             =   1470
      Width           =   495
   End
   Begin VB.Image img_AR 
      Height          =   390
      Index           =   20
      Left            =   630
      Picture         =   "frm_Address_Book.frx":F5BAC
      Top             =   6030
      Width           =   330
   End
   Begin VB.Image img_AR 
      Height          =   390
      Index           =   19
      Left            =   300
      Picture         =   "frm_Address_Book.frx":F5EC9
      Top             =   6030
      Width           =   330
   End
   Begin VB.Image img_AR 
      Height          =   315
      Index           =   18
      Left            =   7800
      Picture         =   "frm_Address_Book.frx":F61CE
      ToolTipText     =   "Count "
      Top             =   6000
      Width           =   360
   End
   Begin VB.Image img_AR 
      Height          =   315
      Index           =   17
      Left            =   7860
      Picture         =   "frm_Address_Book.frx":F653E
      ToolTipText     =   "Border Style"
      Top             =   5640
      Width           =   360
   End
   Begin VB.Image img_AR 
      Height          =   555
      Index           =   16
      Left            =   7920
      Picture         =   "frm_Address_Book.frx":F68CC
      ToolTipText     =   "Save Prefrence"
      Top             =   5130
      Width           =   900
   End
   Begin VB.Image img_AR 
      Height          =   555
      Index           =   15
      Left            =   7890
      Picture         =   "frm_Address_Book.frx":F6E53
      ToolTipText     =   "Search Box Background Color"
      Top             =   4530
      Width           =   900
   End
   Begin VB.Image img_AR 
      Height          =   555
      Index           =   14
      Left            =   8010
      Picture         =   "frm_Address_Book.frx":F73C8
      ToolTipText     =   "Search Box Font Color"
      Top             =   3960
      Width           =   900
   End
   Begin VB.Image img_AR 
      Height          =   555
      Index           =   13
      Left            =   8130
      Picture         =   "frm_Address_Book.frx":F7909
      ToolTipText     =   "Search Option Font Color"
      Top             =   3390
      Width           =   900
   End
   Begin VB.Image img_AR 
      Height          =   555
      Index           =   12
      Left            =   8130
      Picture         =   "frm_Address_Book.frx":F7E7F
      ToolTipText     =   "Text Background Colour"
      Top             =   2820
      Width           =   900
   End
   Begin VB.Image img_AR 
      Height          =   555
      Index           =   11
      Left            =   8010
      Picture         =   "frm_Address_Book.frx":F83A2
      ToolTipText     =   "Text Font Color "
      Top             =   2220
      Width           =   900
   End
   Begin VB.Image img_AR 
      Height          =   555
      Index           =   10
      Left            =   7920
      Picture         =   "frm_Address_Book.frx":F88C7
      ToolTipText     =   "Text Font Size"
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4050
      TabIndex        =   30
      Top             =   5010
      Width           =   675
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   4845
      X2              =   7785
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Label lbl_Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check All"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   6
      Left            =   6990
      TabIndex        =   29
      Top             =   5910
      Width           =   675
   End
   Begin VB.Label lbl_Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (O)"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   5
      Left            =   6750
      TabIndex        =   28
      Top             =   5550
      Width           =   825
   End
   Begin VB.Label lbl_Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (O)"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   5340
      TabIndex        =   27
      Top             =   5550
      Width           =   720
   End
   Begin VB.Label lbl_Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   3
      Left            =   4350
      TabIndex        =   26
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label lbl_Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address (R)"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   6750
      TabIndex        =   25
      Top             =   5220
      Width           =   825
   End
   Begin VB.Label lbl_Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone (R)"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   5310
      TabIndex        =   24
      Top             =   5250
      Width           =   720
   End
   Begin VB.Label lbl_Search 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   0
      Left            =   4320
      TabIndex        =   23
      Top             =   5220
      Width           =   420
   End
   Begin VB.Label lbl_Display 
      BackStyle       =   0  'Transparent
      Height          =   945
      Left            =   1260
      TabIndex        =   13
      Top             =   6450
      Width           =   6465
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_Address_Off 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address Office"
      Height          =   885
      Left            =   4800
      TabIndex        =   12
      ToolTipText     =   "Address (Office)"
      Top             =   3270
      Width           =   2955
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_Address_Res 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address Res"
      Height          =   795
      Left            =   4800
      TabIndex        =   11
      ToolTipText     =   "Address (Residence)"
      Top             =   2400
      Width           =   2955
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_Email 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Email"
      Height          =   405
      Left            =   4800
      TabIndex        =   10
      ToolTipText     =   "Email"
      Top             =   1920
      Width           =   2955
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_Tel_Off 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tel Off"
      Height          =   735
      Left            =   4800
      TabIndex        =   9
      ToolTipText     =   "Telephone (Office)"
      Top             =   1110
      Width           =   2955
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_Tel_Res 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tel Res"
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      ToolTipText     =   "Telephone (Residence)"
      Top             =   420
      Width           =   2955
      WordWrap        =   -1  'True
   End
   Begin VB.Image img_AR 
      Height          =   330
      Index           =   9
      Left            =   8820
      Picture         =   "frm_Address_Book.frx":F8DC2
      ToolTipText     =   "Word Page Setup"
      Top             =   1140
      Width           =   300
   End
   Begin VB.Image img_AR 
      Height          =   270
      Index           =   8
      Left            =   8310
      Picture         =   "frm_Address_Book.frx":F9102
      ToolTipText     =   "Save As"
      Top             =   1140
      Width           =   375
   End
   Begin VB.Image img_AR 
      Height          =   300
      Index           =   7
      Left            =   8820
      Picture         =   "frm_Address_Book.frx":F9462
      ToolTipText     =   "Refresh"
      Top             =   750
      Width           =   495
   End
   Begin VB.Image img_AR 
      Height          =   300
      Index           =   6
      Left            =   8160
      Picture         =   "frm_Address_Book.frx":F97B8
      ToolTipText     =   "Add New Record"
      Top             =   750
      Width           =   495
   End
   Begin VB.Image img_AR 
      Height          =   255
      Index           =   5
      Left            =   8820
      Picture         =   "frm_Address_Book.frx":F9B52
      ToolTipText     =   "Next"
      Top             =   390
      Width           =   315
   End
   Begin VB.Image img_AR 
      Height          =   255
      Index           =   4
      Left            =   8370
      Picture         =   "frm_Address_Book.frx":F9E27
      ToolTipText     =   "Previous"
      Top             =   420
      Width           =   315
   End
   Begin VB.Image img_AR 
      Height          =   390
      Index           =   2
      Left            =   8850
      Picture         =   "frm_Address_Book.frx":FA101
      ToolTipText     =   "Exit"
      Top             =   5760
      Width           =   345
   End
   Begin VB.Image img_AR 
      Height          =   375
      Index           =   3
      Left            =   8910
      Picture         =   "frm_Address_Book.frx":FA432
      ToolTipText     =   "About Enterprise Manager"
      Top             =   6270
      Width           =   255
   End
   Begin VB.Image img_AR 
      Height          =   765
      Index           =   1
      Left            =   8160
      Picture         =   "frm_Address_Book.frx":FA72E
      ToolTipText     =   "Delete Selected Record"
      Top             =   5790
      Width           =   585
   End
   Begin VB.Image img_AR 
      Height          =   840
      Index           =   0
      Left            =   180
      Picture         =   "frm_Address_Book.frx":FABD2
      ToolTipText     =   "Windows"
      Top             =   330
      Width           =   825
   End
   Begin VB.Label lbl_Deail 
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   6
      Top             =   480
      Width           =   210
   End
   Begin VB.Label lbl_Deail 
      BackStyle       =   0  'Transparent
      Caption         =   "(O)"
      Height          =   195
      Index           =   1
      Left            =   4530
      TabIndex        =   5
      Top             =   1230
      Width           =   210
   End
   Begin VB.Image img_Labels 
      Height          =   480
      Index           =   0
      Left            =   4050
      Picture         =   "frm_Address_Book.frx":FB149
      ToolTipText     =   "Telephone (Residence)"
      Top             =   390
      Width           =   480
   End
   Begin VB.Image img_Labels 
      Height          =   315
      Index           =   1
      Left            =   4020
      Picture         =   "frm_Address_Book.frx":FBE13
      ToolTipText     =   "Telephone (Office)"
      Top             =   1140
      Width           =   450
   End
   Begin VB.Image img_Labels 
      Height          =   480
      Index           =   2
      Left            =   4110
      Picture         =   "frm_Address_Book.frx":FBEE6
      ToolTipText     =   "Digital ID's (Email etc)"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image img_Labels 
      Height          =   480
      Index           =   3
      Left            =   4050
      Picture         =   "frm_Address_Book.frx":FC319
      ToolTipText     =   "Office Address"
      Top             =   3330
      Width           =   480
   End
   Begin VB.Label lbl_Deail 
      BackStyle       =   0  'Transparent
      Caption         =   "(R)"
      Height          =   195
      Index           =   2
      Left            =   4500
      TabIndex        =   4
      Top             =   2610
      Width           =   210
   End
   Begin VB.Label lbl_Deail 
      BackStyle       =   0  'Transparent
      Caption         =   "(O)"
      Height          =   195
      Index           =   4
      Left            =   4530
      TabIndex        =   3
      Top             =   3390
      Width           =   210
   End
   Begin VB.Image img_Labels 
      Height          =   525
      Index           =   4
      Left            =   4050
      Picture         =   "frm_Address_Book.frx":FC44D
      ToolTipText     =   "Residence Address"
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label lbl_Deail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Groups"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   1
      Top             =   1170
      Width           =   510
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Visible         =   0   'False
      Begin VB.Menu mnu_Windows 
         Caption         =   "Graphics Menu"
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Refresh 
      Caption         =   "&Refresh"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_Add 
      Caption         =   "&Add"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "&Delete"
      Visible         =   0   'False
      Begin VB.Menu mnu_Delete 
         Caption         =   "All"
         Index           =   0
      End
      Begin VB.Menu mnu_Delete 
         Caption         =   "Individual"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Count 
      Caption         =   "&Count"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Visible         =   0   'False
      Begin VB.Menu mnu_Format 
         Caption         =   "Font"
         Index           =   0
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Font Color"
         Index           =   1
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Back Color"
         Index           =   2
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Search Option Fore Color"
         Index           =   3
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Search Box Fore Color"
         Index           =   4
      End
      Begin VB.Menu mnu_Format 
         Caption         =   "Search Box Background Color"
         Index           =   5
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Save_Preference 
         Caption         =   "Save Preference"
      End
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnu_Save 
         Caption         =   "Save (HTML File)"
         Index           =   0
      End
      Begin VB.Menu mnu_Save 
         Caption         =   "Save (Word File)"
         Index           =   1
      End
   End
   Begin VB.Menu mnu_Word_Page_Setup 
      Caption         =   "Word Page Setup"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuGroups 
      Caption         =   "&Group"
      Visible         =   0   'False
      Begin VB.Menu mnu_Group_Configure 
         Caption         =   "Configure"
      End
      Begin VB.Menu mnu_Group_Refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnu_sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Group 
         Caption         =   "Group_Name"
         Index           =   0
      End
   End
   Begin VB.Menu mnu_About 
      Caption         =   "&About"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frm_Address_Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Option Compare Text '' compulsory requirenment for binary search...
 '''''''''''''''''''''''''''''''''''''''''
 Dim Num As Long
 Dim str_ID() As String
 Dim str_Name() As String
 Dim str_Tel_Res() As String
 Dim str_Tel_Off() As String
 Dim str_Email() As String
 Dim str_Address_Res() As String
 Dim str_Address_Off() As String
 Dim str_Comments() As String
 Dim str_Group() As String
 Dim i As Integer
 ''''''''''''''''''''''''''''''''''''''''
 Dim Last_Index As Integer
 Dim bol_Img As Boolean
 Dim old_Index As Byte
 '''''''''''''''''''''''''''''''''''''''''
 Dim BS As Integer '' Border Style ..this state would be saved into prefrence..
 '''''''''''''''''''''''''''''''''''''''''
 Dim AB_Group() As String
 Const GROUP_DLMT As String = "___"
 '''''''''''''''''''''''''''''''''''''''''
 Dim Font_Name As String
 Dim Font_Bold As String
 Dim Font_Italic As String
 Dim Font_Strikethru As String
 Dim Font_Underline  As String
 Dim Font_Size As Long
 Dim Font_Color As OLE_COLOR
 Dim Back_Color As OLE_COLOR
 Dim Search_Option_Color As OLE_COLOR
 Dim Search_Box_ForeColor As OLE_COLOR
 Dim Search_Box_BGColor As OLE_COLOR
 '''''''''''''''''''''''''''''''''''''''''
 Const Chk_BG_Color As Long = &HFFFF&
 Const UnChk_BG_Color As Long = &H8000000F
 Const offset As Byte = 23
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
Private Sub chk_Email_Click()
 'If chk_Email.Value = vbChecked Then
 ' lbl_s_Email.ForeColor = Chk_BG_Color
 'Else
 ' lbl_s_Email.ForeColor = UnChk_BG_Color
 'End If
 'Check_Search_Status
End Sub

Private Sub chk_Name_Click()
 'If chk_Name.Value = vbChecked Then
 '  lbl_Name.ForeColor = Chk_BG_Color
 'Else
 '  lbl_Name.ForeColor = UnChk_BG_Color
 'End If
 'Check_Search_Status
End Sub

Private Sub chk_Office_Address_Click()
 ' If chk_Office_Address.Value = vbChecked Then
 '  lbl_Office_Address.ForeColor = Chk_BG_Color
 'Else
 '  lbl_Office_Address.ForeColor = UnChk_BG_Color
 'End If
 'Check_Search_Status
End Sub

Private Sub chk_Office_Phone_Click()
  'If chk_Office_Phone.Value = vbChecked Then
  ' lbl_Office_Phone.ForeColor = Chk_BG_Color
 'Else
 '  lbl_Office_Phone.ForeColor = UnChk_BG_Color
 'End If
 'Check_Search_Status
End Sub

Private Sub chk_Residence_Address_Click()
 'If chk_Residence_Address.Value = vbChecked Then
 '  lbl_Residence_Address.ForeColor = Chk_BG_Color
 'Else
 '  lbl_Residence_Address.ForeColor = UnChk_BG_Color
 'End If
 'Check_Search_Status
End Sub

Private Sub Check_Search_Status()
' If chk_Name.Value = vbChecked And chk_Residence_Phone.Value = vbChecked And chk_Office_Phone.Value = vbChecked And chk_Residence_Address.Value = vbChecked And chk_Office_Address.Value = vbChecked And chk_Email.Value = vbChecked Then
'  chk_Search_All.Value = vbChecked
 'Else
'   chk_Search_All.Value = vbUnchecked
 'End If
End Sub

Private Sub chk_Residence_Phone_Click()
 ' If chk_Residence_Phone.Value = vbChecked Then
 '  lbl_Residence_Phone.ForeColor = Chk_BG_Color
 'Else
 '  lbl_Residence_Phone.ForeColor = UnChk_BG_Color
 'End If
 'Check_Search_Status
End Sub

'' This Event is invoked at the the click event of check box Search All
'' Its function is toggle
'' that is either it will set the check value of all of the other chck box control to Yes
'' or if all the chceck box are already checked then this value would be reset as False

Private Sub chk_Search_All_Click()
 Dim Ctrl As Variant
 Dim BG_Color As Long
 '''''''''''''''''''''''
 If chk_Search_All.Value = vbChecked Then
  For Each Ctrl In Controls
   If TypeOf Ctrl Is CheckBox Then
    Ctrl.Value = vbChecked
   End If
  Next
  BG_Color = Chk_BG_Color
  
 ElseIf chk_Search_All.Value = vbUnchecked Then
  For Each Ctrl In Controls
   If TypeOf Ctrl Is CheckBox Then
    Ctrl.Value = vbUnchecked
   End If
  Next
  BG_Color = UnChk_BG_Color
 End If
 
End Sub

'' will be fired during the begining of the form
'' It initlize a lot of variables
'' call the other user defined routines
'' and prompt the user with error (if any) at the display

Private Sub Form_Load()
 strMsg = vbNullString ''' must be at the top statement in load event
 On Error GoTo EH
 Load_Transparent_BG Me
 Populate_Menu_Images
 Retrieve_Preference arr_Windows(0)
 Load_Menu Me, 0
 Load_Group_Menu
 Format_Skin
' mnu_Refresh_Click
 chk_Search_All.Value = vbChecked
 chk_Search_All_Click
  strMsg = "Welcome to Address Book by" & EM_MAIL
  old_Index = 5
  Initialize_Page_Setup_Data
 Exit Sub
EH:
 lbl_Display.Caption = Err.Description
End Sub

'' an Array of Image control
'' MouseDown event will check which image control has been pressed (called)
'' on the basis of index it will decide which routine is to be called..

Private Sub img_AR_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Select Case Index
  Case 0: PopupMenu mnuWindows
  Case 1: PopupMenu mnuDelete
  Case 2: Unload Me
  Case 3: mnu_About_Click
  Case 4:
   If lst_Address.ListIndex < 1 Then Exit Sub
    lst_Address.ListIndex = lst_Address.ListIndex - 1 '' next
    
  Case 5:
   If lst_Address.ListIndex >= lst_Address.ListCount - 1 Then Exit Sub
   lst_Address.ListIndex = lst_Address.ListIndex + 1 '' previous ''
  
  Case 6: mnu_Add_Click
  Case 7: mnu_Refresh_Click
  Case 8:  PopupMenu mnuSave
  Case 9:  mnu_Word_Page_Setup_Click
  Case 10: mnu_Format_Click (0) '' Text Font Size
  Case 11: mnu_Format_Click (1) '' Text Font Color
  Case 12: mnu_Format_Click (2) '' Background Color
  Case 13:  mnu_Format_Click (3) '' Search Font Color
  Case 14: mnu_Format_Click (4) '' Search Box Fore Color
  Case 15: mnu_Format_Click (5) '' Search Box Background Color
  Case 16: Save_Preference arr_Windows(0)
  Case 17: Label_Border_Style
  Case 18: mnu_Count_Click
  Case 19: '' reserved for future use...
  Case 20: '' reserved for future use...
  Case 21: PopupMenu mnuGroups
 End Select
End Sub

'' User Defined Routine.
'' Its function is toggl between enabling and disabling the borderstyle of some label controls
'' these label controls are responsible to hold the address book individual data ...

Private Sub Label_Border_Style()
 If BS = 1 Then
  BS = 0
 Else
  BS = 1
 End If
 ''''''''''''''''''''''
 lbl_Tel_Res.BorderStyle = BS
 lbl_Tel_Off.BorderStyle = BS
 lbl_Email.BorderStyle = BS
 lbl_Address_Res.BorderStyle = BS
 lbl_Address_Off.BorderStyle = BS
 lbl_Comments.BorderStyle = BS
End Sub

'' Click Event of Address List
'' First it will get the value of listindex of the selected item
'' on the basis of this listinidex .. already populated arrays would be contacted and their corresponding data would
'' be retried...
'' 2nd part of the code is responsible to return the group names associated with this address book entry

Private Sub lst_Address_Click()
 Dim A As Variant, i As Long
 '''''''''''''''''''''''''''
 If lst_Address.ListIndex = -1 Then Exit Sub
 '''''''''''''''''''''''''''
 Num = lst_Address.ListIndex
 lbl_Display.Caption = lst_Address.Text
 lbl_Tel_Res.Caption = str_Tel_Res(Num)
 lbl_Tel_Off.Caption = str_Tel_Off(Num)
 lbl_Email.Caption = str_Email(Num)
 lbl_Address_Res.Caption = str_Address_Res(Num)
 lbl_Address_Off.Caption = str_Address_Off(Num)
 lbl_Comments.Caption = str_Comments(Num)
 
 lst_Group.Clear
 If str_Group(Num) <> vbNullString Then
  A = Split(str_Group(Num), GROUP_DLMT)
  For i = 1 To UBound(A) Step 2
   lst_Group.AddItem A(i)
  Next
  ''txt = Left(txt, Len(txt) - 2)
 End If
End Sub

'' Event Routine attached with Address List ' Double Click '
'' call to frm_Address_Book_Action for data Edit/Update

Private Sub lst_Address_DblClick()
 Action_Mode = "Update"
 Edit_ID = str_ID(lst_Address.ListIndex)
 frm_Address_Book_Action.Show 1
End Sub

'' Event Routine attached with menu Add ' Single Click '
'' call to frm_Address_Book_Action for data Addition (New Address Book Value...)

Private Sub mnu_Add_Click()
 Action_Mode = "Add"
 frm_Address_Book_Action.Show 1
End Sub

'' Event Routine attched with menu Count 'Single Click'
'' Diplay the total number of Address Book Values in the Listbox
Private Sub mnu_Count_Click()
 MsgBox "Total Address Items:- " & vbCrLf & vbCrLf & Space(7) & lst_Address.ListCount, vbInformation, EM_TITLE
End Sub

'' Event Routine attached with menu Exit 'Single Click'
'' Quit whole of the Enterprise Manager Application
Private Sub mnu_Exit_Click()
 End
End Sub

'' Event Routine attached with menu Delete 'Single Click'
'' it will delete single or all of the record from the address book database
'' on the basis of index value it will determine whether
'' all of the reocrd(s) or single record is to be erase
'' accordingly it will generate a sql query
'' In next step it will prompt the user for further authentication via Message Box dialog Box
'' If User is really intended to proceed then action woul be taken
'' Action is taken as a user defined routine Delete_Data is called
'' However if User is not intended to delete the data then the routine would be ended.

Private Sub mnu_Delete_Click(Index As Integer)
 Dim Person As String
 Dim Answer As VbMsgBoxResult
 Dim Query_1 As String, Query_2 As String
 
 If UBound(str_ID) = 0 Then Exit Sub
 
 Select Case Index
  Case 0:
    Query_1 = "delete from address"
    Query_2 = "delete from address_book_group_link"
    Person = "All Records "
  Case 1:
    If str_ID(Num) < 1 Then Exit Sub
    Person = Trim(lst_Address.Text)
    If Person = vbNullString Then Exit Sub
    Query_1 = "delete from address where id=" & str_ID(Num)
    Query_2 = "delete from address_book_group_link where person_id=" & str_ID(Num)
 End Select
   Answer = MsgBox("Are you sure to ERASE " & vbCrLf & Person & vbCrLf & " from your address book", vbCritical + vbYesNo, EM_TITLE)
   If Answer = vbNo Then Exit Sub
   Delete_Data Query_1, Query_2, Person
End Sub

'' User Defined Routine for deletion of record
'' It holds three arguements
'' First is query related with the table of Person(s) records . its name is 'address'
'' the table 'address' is linked to another table via primary <-> foreign key
'' the other table named as 'address_book_group_link' table hold the related groups info
'' this related record is delted via the 2nd parametre of the routine..
'' Third argument of the routine is concerned with messaging..

Private Sub Delete_Data(Q1 As String, Q2 As String, Person As String)
 Dim Con_Del As New ADODB.Connection
On Error GoTo EH
 '''''''''''''''''''''''''''
 Con_Del.Open strCon
 Con_Del.Execute Q1
 
 If Err.Number = 0 Then
  Con_Del.Execute Q2
 End If
 
 Set Con_Del = Nothing
 
 If Err.Number = 0 Then
  strMsg = Person & vbCrLf & " has been removed from the Guest Book successfully...."
  MsgBox strMsg, vbOKOnly, EM_TITLE
  Num = -1
  Person = vbNullString
  frm_Progress_Bar.Show vbModal
  mnu_Refresh_Click
 End If
 
 Exit Sub
EH:
 MsgBox Err.Number & vbCrLf & Err.Description, , EM_TITLE
End Sub


'' Click Event of Menu Array
'' on the basis of index value..it will evaluate which menu has been clicked
'' the relevant group related address book values would be called and these record
'' would be populated in the listbox...
'' also as this group has been clicked so this group name becomes disabled and checked
'' and all of the other group name will becomes enabled and unchecked.

Private Sub mnu_Group_Click(Index As Integer)
 Dim R1 As New ADODB.Recordset
 Dim strSql As String, ID As Long
 Dim J As Long
 '''''''''''''''
 strSql = "Select id from address_book_group where name='" & mnu_Group(Index).Caption & "'"
 R1.Open strSql, Con
 If Not R1.EOF Then ID = R1.Fields(0)
 R1.Close
 
 strSql = "Select a.id as id , a.name as name , a.telephone_residence as telephone_residence , a.telephone_office as telephone_office , a.address_residence as address_residence , a.address_office as address_office  , a.email as email , a.comments as comments from address as a inner join address_book_group_link as b on a.id = b.person_id where b.group_id=" & ID & " order by name"
 AB_SQL = strSql
 Populate_List strSql
 
 '''''''''''''''''
 For J = 0 To mnu_Group.Count - 1
  mnu_Group(J).Checked = False
  mnu_Group(J).Enabled = True
 Next
   '''''''''''''
  mnu_Group(Index).Checked = True
  mnu_Group(Index).Enabled = False
 '''''''''''''''''''''''''''''''
End Sub

'' Click Event of Menu Group Configue
'' call another form named as frm_address_book_group
'' this form is used to define the basic data related with group entries..

Private Sub mnu_Group_Configure_Click()
 frm_Address_Book_Groups.Show vbModal
End Sub

'' User Defined Routine..
'' It will refresh all of the information / data on the Application Screen
'' while pulling all of the fresh / uptodated record from the database
'' if any search criteria is provided then it would be trimed and the address book
'' would be populated according to the search list box...
'' More over it calls the user defined routine..Populate_List which is responsible to populate
'' the address list box

Private Sub mnu_Refresh_Click()
 txt_Search.Text = Trim(txt_Search.Text)
 If Len(txt_Search.Text) > 0 Then
  cmd_Search_Click
  Exit Sub
 End If
    '''''''''''''
 strSql = "Select * from address order by name "
 AB_SQL = strSql
 Populate_List strSql
 End Sub

'' Event of Menu About
'' it will call the frm_About
'' This form holds some general information about the author and the application software in general
Private Sub mnu_About_Click()
 frm_About.Show vbModal
End Sub

Private Sub mnu_Group_Refresh_Click()
 Dim i As Integer
 For i = 1 To mnu_Group.Count - 1
  Unload mnu_Group(i)
 Next
 Load_Group_Menu
End Sub


'' Click Event of Array Menu Format
'' on the basis of index value the event will determine which menu has been called

'' When Index Value is 0 then a Font Common Dialog will appear.
''this can be used to set the values of
'' font size , font underline , font italic , font strikethru , font name etc
'' These values would be used to format the skin of the application. so all of the related
'' controls's fonts would be reset accordingly.
'' it calls another user defined routine named as format_skin

'' When Index Value is 1 then it will call the Color Common Dialog Control
''it will retrieve the color value which would be set to the Font Color of the Address List Box

'' When Index Value is 2 then it will call the Color Common Dialog Control
''it will retrieve the color value which would be set to the Background Color of the Address List Box

'' When Index Value is 3 then it will call the Color Common Dialog Control
''it will retrieve the color value which would be set to the Font Color of the Search Box control

'' When Index Value is 4 then it will call the Color Common Dialog Control
''it will retrieve the color value which would be set to the Backgroun Color of the Search Box control


Private Sub mnu_Format_Click(Index As Integer)
 Dim i As Integer
 '''''''''''''''
Select Case Index
 
 Case 0
  cdl_File.Flags = cdlCFBoth
  cdl_File.ShowFont
  
  If Trim(cdl_File.FontName) = vbNullString Then Exit Sub
  Font_Name = cdl_File.FontName
  Font_Bold = cdl_File.FontBold
  Font_Italic = cdl_File.FontItalic
  Font_Size = cdl_File.FontSize
  Font_Strikethru = cdl_File.FontStrikethru
  Font_Underline = cdl_File.FontUnderline
  
  cdl_File.FontName = vbNullString
  Format_Skin
  
 Case 1:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Font_Color = cdl_File.Color
  lst_Address.ForeColor = Font_Color
 
 Case 2:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Back_Color = cdl_File.Color
  lst_Address.BackColor = Back_Color

 Case 3:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Search_Option_Color = cdl_File.Color
  For i = 0 To lbl_Search.Count - 1
   lbl_Search(i).ForeColor = Search_Option_Color
  Next
  
 Case 4:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Search_Box_ForeColor = cdl_File.Color
  txt_Search.ForeColor = Search_Box_ForeColor
 
Case 5:
  cdl_File.ShowColor
  If cdl_File.Color = 0 Then Exit Sub
  Search_Box_BGColor = cdl_File.Color
  txt_Search.BackColor = Search_Box_BGColor

End Select

End Sub

'' User Defined Routine
'' It is meant to save all of the face values of the
'' controls (Search Box , Address LIst Box , Label Controls etc.
'' These values would be first extracted and then they all are meant to be saved
'' in a database so that these preference can be re-collected again from the database
'' on the startup of the next session of the application.
'' Routine is equipped error handling label as if any error occurs while retrieving and saving
'' the prefrence then an error should be displayed at the display control

Private Sub Save_Preference(arr_Win As String)
 Dim R1 As New ADODB.Recordset
 Dim lng_RecCount As Long
 Dim S As String
  On Error GoTo EH
  '''''''''''''''''
 strSql = "Select * from preferences where em_category='" & arr_Win & "'"
With R1
 .Open strSql, Con, adOpenDynamic, adLockOptimistic
 If .EOF Then .AddNew
 .Fields(1) = arr_Win
  S = Font_Name & DLMT & Font_Bold & DLMT & Font_Italic & DLMT & Font_Underline & DLMT & Font_Strikethru & DLMT & Font_Size & DLMT & Font_Color & DLMT & Back_Color
  S = S & DLMT & BS & DLMT & Search_Option_Color & DLMT & Search_Box_ForeColor & DLMT & Search_Box_BGColor
 .Fields(2) = S
 
 .Update
 .Close
End With
 lbl_Display.Caption = arr_Win & " Preference was saved successfully...."
 Format_Skin
 Exit Sub
EH:
 lbl_Display.Caption = Err.Description
End Sub

'' User Defined Routine
'' This routine is mean to retrieve (pull) all of the user prefrences related with the
'' controls (Address Book List , Search Box , Label Controls).The values include
'' fore-color , background color and label's border property.
'' if no record is retrieved from the database (as if no prefrence was ever saved...)
'' then a default value would be returned. and these default values will be
'' set to properties the controls mentioned above

Private Sub Retrieve_Preference(arr_Win As String)
 Dim R1 As New ADODB.Recordset
 Dim lng_RecCount As Long
 Dim arr_Pref As Variant
 On Error GoTo EH
   
 strSql = "Select * from preferences where em_category='" & arr_Win & "'"
 R1.Open strSql, Con, adOpenDynamic, adLockOptimistic
 
 If Not R1.EOF Then
  arr_Pref = Split(R1.Fields(2), DLMT)
  Font_Name = arr_Pref(0)
  Font_Bold = arr_Pref(1)
  Font_Italic = arr_Pref(2)
  Font_Underline = arr_Pref(3)
  Font_Strikethru = arr_Pref(4)
  Font_Size = arr_Pref(5)
  Font_Color = arr_Pref(6)
  Back_Color = arr_Pref(7)
  BS = arr_Pref(8)
  Search_Option_Color = arr_Pref(9)
  Search_Box_ForeColor = arr_Pref(10)
  Search_Box_BGColor = arr_Pref(11)
  ''' now close the connection...
  R1.Close
Else
   Font_Name = "Arial"
   Font_Bold = False
   Font_Italic = False
   Font_Underline = False
   Font_Strikethru = False
   Font_Size = 8
   Font_Color = vbBlack
   Back_Color = vbWhite
   BS = 1
   Search_Option_Color = vbYellow
   Search_Box_ForeColor = vbBlack
   Search_Box_BGColor = vbWhite
  
End If

'''''''''''''''''''''''''''''''''''''
 Exit Sub
EH:
 lbl_Display.Caption = Err.Number & vbCrLf & Err.Description
End Sub

'' User Defined Routine
'' It will polish all of the controls (Address Book List , Label Controls , Search Box)
'' the values which are to be assigned to these controls
'' are already populated from the call of the Retrieve_Prefrence()
'' these values are related to the fore-color , backgroun color , font size , font underline etc
'' and the border property of the label controls

Private Sub Format_Skin()
 lst_Address.FontName = Font_Name
 lst_Address.FontBold = Font_Bold
 lst_Address.FontItalic = Font_Italic
 lst_Address.FontUnderline = Font_Underline
 lst_Address.FontStrikethru = Font_Strikethru
 lst_Address.FontSize = Font_Size
 lst_Address.ForeColor = Font_Color
 lst_Address.BackColor = Back_Color
 '''''''''''''''''''''''''''''''''''
 txt_Binary_Search.FontName = Font_Name
 txt_Binary_Search.ForeColor = Font_Color
 txt_Binary_Search.BackColor = Back_Color
 lbl_Tel_Res.FontName = Font_Name
 lbl_Tel_Res.ForeColor = Font_Color
 lbl_Tel_Off.FontName = Font_Name
 lbl_Tel_Off.ForeColor = Font_Color
 lbl_Email.FontName = Font_Name
 lbl_Email.ForeColor = Font_Color
 lbl_Address_Res.FontName = Font_Name
 lbl_Address_Res.ForeColor = Font_Color
 lbl_Address_Off.FontName = Font_Name
 lbl_Address_Off.ForeColor = Font_Color
 lbl_Comments.FontName = Font_Name
 lbl_Comments.ForeColor = Font_Color
 lst_Group.FontName = Font_Name
 lst_Group.ForeColor = Font_Color
 
 For i = 0 To lbl_Search.Count - 1
  lbl_Search(i).ForeColor = Search_Option_Color
 Next
 
 txt_Search.ForeColor = Search_Box_ForeColor
 txt_Search.BackColor = Search_Box_BGColor
 
 ''''''''''''''''''''''''''''''''''''
 lbl_Tel_Res.BorderStyle = BS
 lbl_Tel_Off.BorderStyle = BS
 lbl_Email.BorderStyle = BS
 lbl_Address_Res.BorderStyle = BS
 lbl_Address_Off.BorderStyle = BS
 lbl_Comments.BorderStyle = BS
End Sub

'' Single Click Event of the Array Menu Save
'' On the basis of the index value of the menu.
'' When the value of index is 0' then frm_HTML_Wizar would be called

'' When the value of the index is 1' then File Save Common Dialog Control will appear.
'' this will be used for getting a file name from user of the application
'' this file name would be used to create a word file in which while of the
'' Address List Box Records would be saved...
'' For this another user defined routine namely Write_Word_Doc would be called..

Private Sub mnu_Save_Click(Index As Integer)
 Dim File_Name As String
 Select Case Index
  Case 0: frm_HTML_Wizard.Show vbModal
  
  Case 1:
   cdl_File.FilterIndex = 1
   cdl_File.ShowSave
   File_Name = cdl_File.fileName
   cdl_File.fileName = vbNullString
   If File_Name = vbNullString Then Exit Sub
   '''''''''''''''''''''
   lbl_Display.Caption = "Writing Word File..." & File_Name
    Write_Word_DOC File_Name
   lbl_Display.Caption = "Finished Writing Word File..." & File_Name
   
 End Select
End Sub

'' Single Click Event of the menu Word Page
'' it will call the frm_Page_Setup form

Private Sub mnu_Word_Page_Setup_Click()
 frm_Page_Setup.Show vbModal
End Sub
 
'' KeyPress Event of a text Box
'' This routine will call another routine namely Binary_Search
'' However this routine will be called only when focus is at this Text Box
'' and enter key is pressed and also the Address List Box holds some records
 
Private Sub txt_Binary_Search_KeyPress(KeyAscii As Integer)
 If (lst_Address.ListCount = 0) Then Exit Sub
 If KeyAscii = 13 Then Binary_Search txt_Binary_Search, lst_Address
End Sub

'' User Defined Routine
'' This routine is used to fast search the data by using the binary algorithm technique.
'' say there are 100 records in the list box ...
'' all of the list box records are sorted
'' it will check whether the value in the middle is equal / greater / smaller in alphabetical terms
'' if it is equal then loop will stop
'' if it is greater then it means the required data is before the middle value and so half of the list box
'' is waived off for possible search. now again this half list is subjected to the same process
'' unless we reach to the required record. if record is not fonund then most relevant record would be returned..
'' in this way , this binary search algorithm acts very fastly...

Private Sub Binary_Search(T As TextBox, L As ListBox)
 Dim Hit As Boolean, l_Top As Long
 Dim l_Bottom As Long, l_Tmp As Long
 '''''''''''''''''''''''''''''''''''''
 If Trim(T.Text) = vbNullString Then
  T.SetFocus
  Exit Sub
 End If
  '''''''''''''''''''''''''''''''''''''
 Hit = False
 l_Top = L.ListCount - 1
 l_Bottom = 0
 l_Tmp = (l_Top + l_Bottom) / 2
 Do While (Not Hit) And (l_Top >= l_Bottom)
  If L.List(l_Tmp) = T.Text Then
   Hit = True
  ElseIf T.Text < L.List(l_Tmp) Then
   l_Top = l_Tmp - 1
  Else
   l_Bottom = l_Tmp + 1
  End If
   l_Tmp = (l_Top + l_Bottom) / 2
  If Not Hit Then
   'txt_Display.Text = txt_Display.Text & vbCrLf & "No match"
  Else
   'txt_Display.Text = List1.Text & " -- " & lst_Store_ID.ListIndex
  End If
  Loop
  'L.SetFocus
  If l_Tmp >= L.ListCount Then l_Tmp = L.ListCount - 1
  L.Selected(l_Tmp) = True
End Sub
 
'' KeyPress Event of Search Box
'' When Enter Key is pressed while the focus is at this text box
'' it will set the value of the command button as true. which means
'' the click event of the command button is required to be be exectured

Private Sub txt_Search_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  cmd_Search.Value = True
 End If
End Sub

'' Single Click Event of the Search Command Button
'' if not value is found in the Search Text Box then the event will be ended premature resulting nothing..
'' if some value in the search text box is found...then it will start to see
'' which check boxes are checked....
'' depending on the checked controls this routine will create a complex Query so that
'' relevant information based on the given criteria be returned..
'' This complex (dynamic) query is then passed to a user defined routine Populate_List
'' which will act according to the query passed....

Private Sub cmd_Search_Click()
 Dim strSearch As String
 On Error Resume Next
 ''''''''''''''''''''''''
 strSearch = Trim(txt_Search.Text)
 If Len(strSearch) = 0 Then
  Exit Sub
 End If
  
 strSql = "Select * from address Where "
 If chk_Name.Value = vbChecked Then
  strSql = strSql & " OR name like '%" & strSearch & "%'"
 End If
 If chk_Residence_Phone.Value = vbChecked Then
  strSql = strSql & " OR telephone_residence like '%" & strSearch & "%'"
 End If
 If chk_Office_Phone.Value = vbChecked Then
  strSql = strSql & " OR telephone_office like '%" & strSearch & "%'"
 End If
 If chk_Residence_Address.Value = vbChecked Then
  strSql = strSql & " OR address_residence like '%" & strSearch & "%'"
 End If
 If chk_Office_Address.Value = vbChecked Then
  strSql = strSql & " OR address_office like '%" & strSearch & "%'"
 End If
 If chk_Email.Value = vbChecked Then
  strSql = strSql & " OR email like '%" & strSearch & "%'"
 End If
 
 strSql = strSql & " order by name "
 
 strSql = Replace(strSql, " Where  OR ", " Where ")
 
 If strSql = "Select * from address Where  order by name " Then
  lst_Address.Clear
  Exit Sub
 End If

 Populate_List strSql
 
End Sub

'' User Defined Routine
'' 1st Step: It will initialize all of the global dynamic arrays
'' 2nd Step: Clear All of the controls including listbox , label controls etc.
'' 3rd Step: based on the query , it will fetch all of the records from database file
'' 4th Step: fill all of the dynamic arrays with the fetched records.. A counter increment is required in this process
'' 5th Step: Convey the user about the sccessfully loading of the records by displaying message at display control

Private Sub Populate_List(str_SQL As String)
 Dim R1 As New ADODB.Recordset
 Dim lng_Records As Long, i As Long
 Static Cnt As Integer
 On Error GoTo EH
 '''''''''''''''''''''
 Cnt = Cnt + 1
 If Cnt > 1 Then Exit Sub '' routine is not allowed to be invoked if already in execution...this helps in avoiding duplicate entries into the list box....
 txt_Binary_Search.Enabled = False
 lst_Address.Clear
 '''''''''''''''''
 Erase str_ID
 Erase str_Name
 Erase str_Tel_Res
 Erase str_Tel_Off
 Erase str_Address_Res
 Erase str_Address_Off
 Erase str_Email
 Erase str_Group
 '''''''''''''''''''
 
 lbl_Tel_Res.Caption = vbNullString
 lbl_Tel_Off.Caption = vbNullString
 lbl_Email.Caption = vbNullString
 lbl_Address_Res.Caption = vbNullString
 lbl_Address_Off.Caption = vbNullString
 lbl_Comments.Caption = vbNullString
 lbl_Display.Caption = vbNullString
 lst_Group.Clear
  ''''''''''''''''''
 R1.Open str_SQL, Con, 3
 lng_Records = R1.RecordCount
 ReDim str_ID(lng_Records)
 ReDim str_Name(lng_Records)
 ReDim str_Tel_Res(lng_Records)
 ReDim str_Tel_Off(lng_Records)
 ReDim str_Email(lng_Records)
 ReDim str_Address_Res(lng_Records)
 ReDim str_Address_Off(lng_Records)
 ReDim str_Comments(lng_Records)
 ReDim str_Group(lng_Records)
  
 While Not R1.EOF
  str_ID(i) = R1.Fields("id")
  str_Name(i) = R1.Fields("name")
  str_Tel_Res(i) = R1.Fields("telephone_residence")
  str_Tel_Off(i) = R1.Fields("telephone_office")
  str_Address_Res(i) = R1.Fields("address_residence")
  str_Address_Off(i) = R1.Fields("address_office")
  str_Email(i) = R1.Fields("email")
  If Not IsNull(R1.Fields("comments")) Then str_Comments(i) = R1.Fields("comments")
  Populate_Group_List R1.Fields("id"), i
  
  lst_Address.AddItem UCase(str_Name(i))
  R1.MoveNext
  i = i + 1: DoEvents
 Wend
 ''''''''''''''''''''''''''''''''''
 R1.Close
 lbl_Display.Caption = "Refreshed the Address Book"
 txt_Binary_Search.Enabled = True
 Cnt = 0
Exit Sub
EH:
 lbl_Display.Caption = Err.Description
End Sub

'' User Defined Routine
'' 1st Step: It will initialize some local variables including a recordset variable
'' 2nd Step: Fetch all of the relevant group information against the selected Address Book Record based on the query
'' 3rd Step: Populate a temporary string variable with these fetched group names
'' 4th Step: set this temporary string value into a form level global variable so that it can be easily processed

 Private Sub Populate_Group_List(Person_ID As Long, ii As Long)
  Dim R1 As New ADODB.Recordset, str_SQL As String
  Dim Txt As String
  
  str_SQL = "select a.id , a.name from address_book_group_link as l inner join address_book_group as a on a.id=l.group_id where l.person_id=" & Person_ID
  R1.Open str_SQL, Con
  While Not R1.EOF
   Txt = Txt & R1.Fields(0) & GROUP_DLMT & R1.Fields(1) & GROUP_DLMT
  R1.MoveNext
  Wend
  R1.Close
  Txt = Trim(Txt)
  If Txt <> vbNullString Then
   str_Group(ii) = Txt
  End If
 End Sub

'' User Defined Routine
'' It will initialize global variables
'' all of these variable are meant to be used in setting the page setup of the word file when word file is in process of creation
Private Sub Initialize_Page_Setup_Data()
'''''''''' page setup related information...............
 Top_Margin = "1.0''"
 Bottom_Margin = "1.0''"
 Left_Margin = "1.25''"
 Right_Margin = "1.25''"
 Table_Border = vbChecked
 Paragraph_Alignment = wdAlignParagraphLeft
End Sub

'' User Defined Routin
'' On Early Binding based Word Application are initilized and invoked
'' via defining two variables , the first variable is related with word application
'' and the second variable is dealt with the table defined in the page of the word
'' Next step are concerned with defining and desiging the table and then filling the cell values with suitable values from the recods
'' Only those records would be displayed on the word file which are visible in the Address Book List control

Public Sub Write_Word_DOC(File_Name As String)
Dim oWordApp As Word.Application
Set oWordApp = New Word.Application
Dim wTable As Table
Dim Doc As Word.Document
Dim Shp As Word.Shape '' oblject for watermarking...only valid for Word-XP classes

Dim P As Long, Q As Long
Dim A As Variant, Txt As String
Dim i As Long, ii As Long
Dim FRow As Integer, TRow As Integer, TCol As Integer
''''''''''''''''''
On Error GoTo EH
''''''''''''''''''''''''''''''''''''
lbl_Display.Caption = "Wait.." & File_Name & " InProcess.."
FRow = 2: TRow = lst_Address.ListCount + 2: TCol = 9
''''''''''''
With oWordApp
 'Create a new document
 .Visible = True
 Set Doc = .Documents.Add
 Doc.PageSetup.TopMargin = Application.InchesToPoints(Val(Top_Margin))
 Doc.PageSetup.BottomMargin = Application.InchesToPoints(Val(Bottom_Margin))
 Doc.PageSetup.LeftMargin = Application.InchesToPoints(Val(Left_Margin))
 Doc.PageSetup.RightMargin = Application.InchesToPoints(Val(Right_Margin))
 Doc.PageSetup.Orientation = wdOrientLandscape
 '''''''''''''''''''''
 
 Set wTable = Doc.Tables.Add(Doc.Range, TRow, TCol)
  'Add text to the document
 wTable.Select
 wTable.Borders.Enable = Table_Border
 wTable.Range.ParagraphFormat.Alignment = Paragraph_Alignment
 
 '''''' Heading information...  '''''''''''
 P = 1: Q = 1
 wTable.Rows(P).Range.Font.Color = wdColorGray65
 wTable.Rows(P).Range.Shading.BackgroundPatternColor = wdColorGray05
 wTable.Cell(P, Q).Merge wTable.Cell(P, TCol)
 wTable.Cell(P, Q).Range.Bold = True
 wTable.Cell(P, Q).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
 wTable.Cell(P, Q).Range.InsertAfter "Address Book " & " (" & Format(Now, "Long Date") & ")"
 wTable.Cell(P, Q).Range.InsertParagraphAfter
 wTable.AllowAutoFit = True
  
 ''''''''''Column heading....'''''''''''''''''''
 P = 2: Q = 1
 wTable.Rows(P).Range.Font.Color = wdColorGray65
 wTable.Rows(P).Range.Shading.BackgroundPatternColor = wdColorGray10
 wTable.Rows(P).Range.Bold = True
 wTable.Cell(P, Q + 0).Range.InsertAfter "Ser #"
 wTable.Cell(P, Q + 1).Range.InsertAfter "Name"
 wTable.Cell(P, Q + 2).Range.InsertAfter "Tel(Res)"
 wTable.Cell(P, Q + 3).Range.InsertAfter "Tel(Off)"
 wTable.Cell(P, Q + 4).Range.InsertAfter "Address (Res)"
 wTable.Cell(P, Q + 5).Range.InsertAfter "Address (Off)"
 wTable.Cell(P, Q + 6).Range.InsertAfter "Email"
 wTable.Cell(P, Q + 7).Range.InsertAfter "Comments"
 wTable.Cell(P, Q + 8).Range.InsertAfter "Group"
 
 ''''''''''''''''''''''''''''''''''''''''
  
 P = FRow: Q = 1
 For i = 0 To lst_Address.ListCount - 1
   P = P + 1
   wTable.Cell(P, Q + 0).Range.InsertAfter (i + 1)
   wTable.Cell(P, Q + 1).Range.InsertAfter lst_Address.List(i)
   wTable.Cell(P, Q + 2).Range.InsertAfter str_Tel_Res(i)
   wTable.Cell(P, Q + 3).Range.InsertAfter str_Tel_Off(i)
   wTable.Cell(P, Q + 4).Range.InsertAfter str_Address_Res(i)
   wTable.Cell(P, Q + 5).Range.InsertAfter str_Address_Off(i)
   wTable.Cell(P, Q + 6).Range.InsertAfter str_Email(i)
   wTable.Cell(P, Q + 7).Range.InsertAfter str_Comments(i)
 
  Txt = vbNullString
  If str_Group(i) <> vbNullString Then
   A = Split(str_Group(i), GROUP_DLMT)
   For ii = 1 To UBound(A) Step 2
    Txt = Txt & A(ii) & " , "
   Next
   Txt = Left(Txt, Len(Txt) - 2)
  End If
   wTable.Cell(P, Q + 8).Range.InsertAfter Txt
 Next
 '''''''''''''''''''''''''''''''''''''''''''''''
 
 DoEvents
 '' Water Marking ......
 '' should be placed at the end of the page....
 Txt = "naeem@email.com" & vbCrLf & "+92-333-5170489"
 Set Shp = Doc.Shapes.AddTextEffect(msoTextEffect4, Txt, "Times New Roman", 1, False, False, 0, 0)
 Shp.TextEffect.NormalizedHeight = False
 Shp.Line.Visible = False
 Shp.Fill.Visible = True
 Shp.Fill.Solid
 Shp.Fill.ForeColor.RGB = RGB(192, 192, 192)
 Shp.Fill.Transparency = 0.5
 Shp.Rotation = 315
 Shp.LockAspectRatio = True
 Shp.Height = InchesToPoints(2.82)
 Shp.Width = InchesToPoints(5.64)
 Shp.WrapFormat.AllowOverlap = True
 Shp.WrapFormat.Side = wdWrapNone
 Shp.WrapFormat.Type = 3
 Shp.RelativeHorizontalPosition = wdRelativeVerticalPositionMargin
 Shp.RelativeVerticalPosition = wdRelativeVerticalPositionMargin
 Shp.Left = wdShapeCenter
 Shp.Top = wdShapeBottom
 .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
 ''''''''end of the water marking.................
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
   'Save the document
.ActiveDocument.SaveAs fileName:=File_Name & ".doc", _
  FileFormat:=wdFormatDocument, LockComments:=False, _
  Password:="", AddToRecentFiles:=True, WritePassword _
  :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
  SaveNativePictureFormat:=False, SaveFormsData:=False, _
  SaveAsAOCELetter:=False
  
End With
'''''''''''''''
'' oWordApp.Quit
'''''''''''''''''''''
Exit Sub
EH:
MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, EM_TITLE
End Sub

Private Sub lbl_display_DblClick()
 lbl_Display.Caption = vbNullString
End Sub
  
Private Sub txt_Tel_Res_DblClick()
 Clipboard.SetText lbl_Tel_Res.Caption
End Sub

Private Sub Load_Group_Menu()
Dim str_SQL As String, i As Long
Dim R1 As New ADODB.Recordset
'''''''''''''''''''''''
i = 0
str_SQL = "Select name from address_book_group"
R1.Open str_SQL, Con
If Not R1.EOF Then
 mnu_Group(i).Caption = R1.Fields(0)
 i = i + 1
 R1.MoveNext
End If
'''''''''''''''''
While Not R1.EOF
 Load mnu_Group(i)
 mnu_Group(i).Caption = R1.Fields(0)
 i = i + 1
 R1.MoveNext
Wend
 R1.Close
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MoveForm Me, X, Y, Button
  If bol_Img = True Then
   img_AR(Last_Index - offset).picture = imgLst.ListImages(Last_Index - offset + 1).picture
   bol_Img = False
  End If
End Sub

Private Sub img_AR_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 Last_Index = Index + offset
 img_AR(Index).picture = imgLst.ListImages(Last_Index).picture
 bol_Img = True
End Sub


Private Sub Populate_Menu_Images()
 Dim i As Integer
 For i = 1 To img_AR.Count - 1
  img_AR(i).picture = imgLst.ListImages(i + 1).picture
 Next
End Sub

Private Sub mnu_windows_Click(Index As Integer)
 Load_Selected_Window Me, Index
End Sub

