VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Loan_Calculator 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8190
   Icon            =   "frm_Loan_Calculator.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic_BMP_Menu 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   705
      Left            =   9210
      ScaleHeight     =   645
      ScaleWidth      =   975
      TabIndex        =   42
      Top             =   660
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmd_HTML 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5430
      Picture         =   "frm_Loan_Calculator.frx":46E2
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Convert to HTML"
      Top             =   6360
      Width           =   315
   End
   Begin VB.CommandButton cmd_Calculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6660
      TabIndex        =   32
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Frame fra_Personal_Info 
      Caption         =   "Personal Information:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   30
      TabIndex        =   31
      Top             =   60
      Width           =   3075
      Begin VB.TextBox txt_Account 
         Height          =   285
         Left            =   1380
         TabIndex        =   40
         Top             =   1260
         Width           =   1600
      End
      Begin VB.TextBox txt_Banker 
         Height          =   285
         Left            =   1380
         TabIndex        =   39
         Top             =   930
         Width           =   1600
      End
      Begin VB.TextBox txt_NTN 
         Height          =   285
         Left            =   1380
         TabIndex        =   38
         Top             =   600
         Width           =   1600
      End
      Begin VB.TextBox txt_Name 
         Height          =   285
         Left            =   1380
         TabIndex        =   37
         Top             =   270
         Width           =   1600
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "A/C Number:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   90
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Banker Name:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   35
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "NTN Nmber:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   34
         Top             =   660
         Width           =   1125
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Name:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.Timer tmr_Courtesy 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7710
      Top             =   -150
   End
   Begin VB.TextBox txt_Display 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   6750
      Width           =   8115
   End
   Begin VB.Frame fra_Output 
      Caption         =   "Output:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   30
      TabIndex        =   6
      Top             =   2100
      Width           =   8115
      Begin MSComctlLib.ListView lvw_Loan 
         Height          =   2835
         Left            =   30
         TabIndex        =   12
         Top             =   180
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   5001
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Intallment"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Principal (Monthly)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Markup (Monthly)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Principal Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Markup Paid"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Principal Left"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Total Loan Left"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl_Total_Payable_Loan 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5700
         TabIndex        =   29
         Top             =   3810
         Width           =   2325
      End
      Begin VB.Label lbl_Total_Payable_Markup 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5700
         TabIndex        =   28
         Top             =   3420
         Width           =   2325
      End
      Begin VB.Label lbl_Total_Loan_Received 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5700
         TabIndex        =   27
         Top             =   3030
         Width           =   2325
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Total Payable Loan:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   3630
         TabIndex        =   26
         Top             =   3810
         Width           =   1770
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Total Payable Mark-Up:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   3600
         TabIndex        =   25
         Top             =   3450
         Width           =   2085
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Total Loan Received:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   3600
         TabIndex        =   24
         Top             =   3120
         Width           =   1875
      End
      Begin VB.Label lbl_Installment_Factor 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2430
         TabIndex        =   23
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label lbl_Monthly_Installment 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2430
         TabIndex        =   22
         Top             =   3390
         Width           =   1095
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Intallment Factor:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   60
         TabIndex        =   21
         Top             =   3780
         Width           =   1605
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Equal Monthly Installment:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   20
         Top             =   3420
         Width           =   2355
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "No. of Installments:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   19
         Top             =   3060
         Width           =   1770
      End
      Begin VB.Label lbl_no_of_installment 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2430
         TabIndex        =   18
         Top             =   3030
         Width           =   1095
      End
   End
   Begin VB.Frame fra_Input 
      Caption         =   "Input:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Left            =   4470
      TabIndex        =   0
      Top             =   60
      Width           =   3675
      Begin VB.ComboBox cbo_Salary 
         Height          =   315
         ItemData        =   "frm_Loan_Calculator.frx":482C
         Left            =   1470
         List            =   "frm_Loan_Calculator.frx":484E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   870
         Width           =   795
      End
      Begin VB.ComboBox cbo_Year 
         Height          =   315
         ItemData        =   "frm_Loan_Calculator.frx":4871
         Left            =   1470
         List            =   "frm_Loan_Calculator.frx":4893
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1230
         Width           =   705
      End
      Begin VB.ComboBox cbo_Rate 
         Height          =   315
         Left            =   1470
         TabIndex        =   5
         Top             =   1620
         Width           =   1275
      End
      Begin VB.TextBox txt_Principal 
         Height          =   285
         Left            =   1470
         TabIndex        =   2
         Text            =   "0"
         Top             =   540
         Width           =   2115
      End
      Begin VB.TextBox txt_Pay 
         Height          =   285
         Left            =   1470
         MaxLength       =   9
         TabIndex        =   1
         Text            =   "0"
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label lbl_Salary 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2460
         TabIndex        =   17
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "No. of Salaries:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   16
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lbl_Year 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2430
         TabIndex        =   14
         Top             =   1260
         Width           =   45
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Year(s):-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   13
         Top             =   1290
         Width           =   780
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "(%)"
         Height          =   195
         Index           =   3
         Left            =   3300
         TabIndex        =   11
         Top             =   1680
         Width           =   210
      End
      Begin VB.Label lbl_Rate 
         AutoSize        =   -1  'True
         Caption         =   "00.00"
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Rate (%):-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   9
         Top             =   1620
         Width           =   930
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Principal:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label_Dummy 
         AutoSize        =   -1  'True
         Caption         =   "Current Pay:-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   270
         Width           =   1185
      End
   End
   Begin MSComDlg.CommonDialog cdl_File 
      Left            =   0
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".htm"
      DialogTitle     =   "Convert Data"
      Filter          =   "HTML File|*.htm;*.html|Text File|*.txt|All Files|*.htm;*.html;*.txt"
   End
   Begin VB.Image img_CoAuthor 
      Height          =   480
      Left            =   4680
      Picture         =   "frm_Loan_Calculator.frx":48B6
      ToolTipText     =   "Khurram Ayob"
      Top             =   6270
      Width           =   480
   End
   Begin VB.Image img_Print 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      Picture         =   "frm_Loan_Calculator.frx":8F98
      ToolTipText     =   "Print Data "
      Top             =   6330
      Width           =   375
   End
   Begin VB.Label lbl_Courtesy 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "courtesy by Khurram Ayob (engrkhurram@yahoo.com)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   30
      TabIndex        =   30
      Top             =   6420
      Width           =   4605
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnu_Windows 
         Caption         =   "Graphics Menu"
         Index           =   0
      End
   End
   Begin VB.Menu mnu_Test 
      Caption         =   "Test"
   End
   Begin VB.Menu mnu_About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frm_Loan_Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''
Dim Rate As Double, Principal As Double, Pay As Double
Dim Year As Integer, Salaries As Integer
Dim i As Integer, C As Integer
Dim Interest_Portion() As Double
Dim Principal_Portion() As Double
Dim Installment As Double
Const Precesion As String = "0.00"

Dim Tot_Int_Por As Double, Tot_Pri_Por As Double
Dim Tot_Loan As Double, Tot_Loan_Left As Double, Tot_Pri_Left As Double
Dim Markup_Paid As Double, Principal_Paid As Double

Dim arr_Courtesy() As String
Const CourtesyMsg As String = "courtesy by Khurram Ayob (engrkhurram@yahoo.com)"
Dim C_Courtesy As Integer
Const Timer_Interval As Integer = 100

Private Sub cmd_Calculate_Click()
 Dim LI As ListItem
 ''''''''''''''''''''''''''''
 On Error GoTo EH
 txt_Display.Text = vbNullString
 ReDim Interest_Portion(Year * 12)
 ReDim Principal_Portion(Year * 12)
 ''''''''''''''''''''''''''''
 Purge_Data
 ''''''''''''''''''''''''''''
 For C = 1 To Year * 12
  Interest_Portion(C) = Abs(IPmt(Rate / 12, C, Year * 12, Principal))
  Principal_Portion(C) = Abs(PPmt(Rate / 12, C, Year * 12, Principal))
  Tot_Int_Por = Tot_Int_Por + Interest_Portion(C)
  Tot_Pri_Por = Tot_Pri_Por + Principal_Portion(C)
 Next
 '''''''''''''''''''''''''''''
 Tot_Loan = Tot_Int_Por + Tot_Pri_Por
 Installment = Interest_Portion(1) + Principal_Portion(1)
 lvw_Loan.ListItems.Clear
 
 For C = 1 To Year * 12
  Principal_Paid = Principal_Paid + Principal_Portion(C)
  Markup_Paid = Markup_Paid + Interest_Portion(C)
  Tot_Loan_Left = Tot_Loan - (Installment * C)
  Tot_Pri_Left = CDbl(txt_Principal.Text) - Principal_Paid
  
  Set LI = lvw_Loan.ListItems.Add(, , C)
   LI.ListSubItems.Add , , Format(Principal_Portion(C), Precesion), , "Monthly Principal Amount in Installment"
   LI.ListSubItems.Add , , Format(Interest_Portion(C), Precesion), , "Monthly Principal Amount in Installment"
   LI.ListSubItems.Add , , Format(Principal_Paid, Precesion)
   LI.ListSubItems.Add , , Format(Markup_Paid, Precesion)
   LI.ListSubItems.Add , , Format(Tot_Pri_Left, Precesion)
   LI.ListSubItems.Add , , Format(Tot_Loan_Left, Precesion)
 Next
     '''''''''''''''''''''''''''''''''''''''''''''''
 lbl_no_of_installment.Caption = Year * 12
 lbl_Monthly_Installment.Caption = Format(Installment, Precesion)
 lbl_Installment_Factor.Caption = vbNullString
 lbl_Total_Loan_Received.Caption = Format(CDbl(txt_Principal.Text), Precesion)
 lbl_Total_Payable_Markup.Caption = Format(Tot_Int_Por, Precesion)
 lbl_Total_Payable_Loan.Caption = Format(Tot_Loan, Precesion)
  ''''''''''''''''''''''''''''''''''''''''''''''''''
 lbl_Courtesy.Caption = vbNullString
  tmr_Courtesy.Enabled = True
  cmd_Calculate.Enabled = False
  img_Print.Enabled = True
  cmd_HTML.Enabled = True
  txt_Display.Text = "Finished Processing Data...."
  Exit Sub
EH:
  txt_Display.Text = Err.Description
End Sub

Private Sub Purge_Data()
 Principal_Paid = 0
 Markup_Paid = 0
 Tot_Loan_Left = 0
 Tot_Pri_Left = 0
 Tot_Loan = 0
 Tot_Int_Por = 0
 Tot_Pri_Por = 0
End Sub

Private Sub cbo_Rate_Click()
 Calculate_Rate
End Sub

Private Sub cbo_Rate_KeyPress(KeyAscii As Integer)
 If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii > 47 And KeyAscii < 58) Then
  ''' go on...continue....
 ElseIf KeyAscii = 13 Then
  Calculate_Rate
 Else
  KeyAscii = 0
 End If
End Sub

Private Sub cbo_Salary_Click()
 Salaries = cbo_Salary.Text
 lbl_Salary.Caption = Salaries
 txt_Pay_CHANGE
End Sub

Private Sub cbo_Year_Click()
 Year = cbo_Year.Text
 lbl_Year.Caption = Year & " Year(s)"
End Sub

Private Sub cmd_HTML_Click()
 Convert
End Sub

Private Sub mnu_About_Click()
 frm_About.Show vbModal
End Sub

Private Sub Form_Load()
 Load_Menu Me, 4
 Me.Caption = EM_TITLE & "..Loan Calculator" & EM_MAIL
 Populate_Data
End Sub

Private Sub Populate_Data()
 Dim dbl_C As Double
 For dbl_C = 0.1 To 100# Step 0.01
  cbo_Rate.AddItem Format(dbl_C, "00.00")
 Next
 ''''''''''''''''''''''''''''''''''''''''''
 For dbl_C = 1 To 20 Step 1
  cbo_Salary.AddItem dbl_C
 Next
 '''''''''''''''''''''''''''''''''''''''''
 Year = 5
 Salaries = 15
 Rate = 0.11
 cbo_Rate.Text = Rate
 lbl_Rate.Caption = 100 * Rate & " %"
 cbo_Year.Text = Year
 lbl_Year.Caption = Year & " Year(s)"
 cbo_Salary.Text = Salaries
 lbl_Salary.Caption = Salaries
 ''''''''''''''''''''''''''''''''''''''''''
 ReDim arr_Courtesy(Len(CourtesyMsg))
 For i = 0 To Len(CourtesyMsg)
  arr_Courtesy(i) = Mid(CourtesyMsg, i + 1, 1)
 Next
 tmr_Courtesy.Interval = Timer_Interval
 ''''''''''''''''''''''''''''''''''''''''''
 
 
End Sub

Private Sub Calculate_Rate()
 Rate = CDbl(cbo_Rate.Text) / 100
 lbl_Rate = Rate
End Sub

Private Sub mnu_Test_Click()
 txt_Name.Text = "Hamid Hussain"
 txt_NTN.Text = "7865-4"
 txt_Banker.Text = "UBL"
 txt_Account.Text = "1234-5"
 txt_Pay.Text = "20000"
End Sub

Private Sub mnu_windows_Click(Index As Integer)
 Load_Selected_Window Me, Index
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub txt_Pay_CHANGE()
 On Error Resume Next
 Pay = CDbl(txt_Pay.Text)
 If Err.Number = 13 Then Pay = 0
 Principal = Pay * Salaries
 txt_Principal.Text = Principal
End Sub

Private Sub txt_Pay_KeyPress(KeyAscii As Integer)
On Error GoTo EH
 If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii > 47 And KeyAscii < 58) Then
  ''' go on...continue....
 ElseIf KeyAscii = 13 Then
  Pay = txt_Pay.Text
 Else
  KeyAscii = 0
 End If
 ''''''''''''''''''''
 Exit Sub
EH:
 Pay = 0
 txt_Display.Text = Err.Description
End Sub

Private Sub txt_Principal_LostFocus()
 On Error Resume Next
 Principal = txt_Principal.Text
 If Principal = 0 Then Principal = 0
 Pay = CDbl(Principal / Salaries)
 txt_Pay.Text = Pay
 txt_Principal.Text = Principal
End Sub

Private Sub tmr_Courtesy_Timer()
 lbl_Courtesy.Caption = lbl_Courtesy.Caption & arr_Courtesy(C_Courtesy)
  C_Courtesy = C_Courtesy + 1
 tmr_Courtesy.Interval = tmr_Courtesy.Interval - (Timer_Interval / UBound(arr_Courtesy))
 If C_Courtesy > UBound(arr_Courtesy) Then
  tmr_Courtesy.Enabled = False
  C_Courtesy = 0
  tmr_Courtesy.Interval = Timer_Interval
  cmd_Calculate.Enabled = True
 End If
End Sub

Private Sub Convert()
  Dim Total_Rec As Long, fileName As String, strFile As String
  Dim DL As String  '' delimiter ....
  Dim Head As String
  Dim strData As String
  Dim obj_File As New FileSystemObject, obj_T As TextStream


On Error GoTo EH

If fileName = vbNullString Then
 cdl_File.ShowSave
 strFile = cdl_File.fileName
 cdl_File.fileName = vbNullString
 If Len(strFile) < 1 Then
   Exit Sub
 End If
Else
 strFile = fileName
End If

 DL = "</td><td>"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
 Head = "<tr><td colspan=7 align=center ><table border=1 bordercolor=blue>"
 Head = Head & "<tr>"
 Head = Head & "<td>Name" & DL & txt_Name.Text & "</td></tr>"
 Head = Head & "<td>NTN" & DL & txt_NTN.Text & "</td></tr>"
 Head = Head & "<td>Banker" & DL & txt_Banker.Text & "</td></tr>"
 Head = Head & "<td>A/C" & DL & txt_Account.Text & "</td></tr>"
 Head = Head & "</td></tr>"
 
 Head = Head & "<tr>"
 Head = Head & "<td>Total Installment" & DL & lbl_no_of_installment.Caption & "</td></tr>"
 Head = Head & "<td>Monthly Installment" & DL & lbl_Monthly_Installment.Caption & "</td></tr>"
 Head = Head & "<td>Total Loan" & DL & lbl_Total_Loan_Received.Caption & "</td></tr>"
 Head = Head & "<td>Total Payable Markup" & DL & lbl_Total_Payable_Markup.Caption & "</td></tr>"
 Head = Head & "<td>Total Payable Loan" & DL & lbl_Total_Payable_Loan.Caption & "</td></tr>"
 Head = Head & "</table></td>"
 Head = Head & "</tr>"
 
 
'''''''''''''''''''''''''''''''''''''''''''''''''''''
 
  Head = Head & "<tr><td><table border=1 bordercolor=green><tr bgcolor='silver'><th>Intallment #</th>"
  Head = Head & "<th>" & lvw_Loan.ColumnHeaders(2).Text & "</th>"
  Head = Head & "<th>" & lvw_Loan.ColumnHeaders(3).Text & "</th>"
  Head = Head & "<th>" & lvw_Loan.ColumnHeaders(4).Text & "</th>"
  Head = Head & "<th>" & lvw_Loan.ColumnHeaders(5).Text & "</th>"
  Head = Head & "<th>" & lvw_Loan.ColumnHeaders(6).Text & "</th>"
  Head = Head & "<th>" & lvw_Loan.ColumnHeaders(7).Text & "</th>"
  Head = Head & "</tr>"
 
  Set obj_T = obj_File.OpenTextFile(strFile, ForWriting, True)
  strData = "<body>"
  strData = strData & "<title>Loan Calculator by " & EM_MAIL & "</title>"
  strData = strData & "<table align=center border=" & 0 & ">"
  strData = strData & "<caption><h2>Bank Loan Calculator</h2></caption>"
  strData = strData & Head
    
  For C = 1 To (Year * 12)
    strData = strData & "<tr><td>" & C & DL
    strData = strData & Format(lvw_Loan.ListItems(C).ListSubItems(1), Precesion) & DL
    strData = strData & Format(lvw_Loan.ListItems(C).ListSubItems(2), Precesion) & DL
    strData = strData & Format(lvw_Loan.ListItems(C).ListSubItems(3), Precesion) & DL
    strData = strData & Format(lvw_Loan.ListItems(C).ListSubItems(4), Precesion) & DL
    strData = strData & Format(lvw_Loan.ListItems(C).ListSubItems(5), Precesion) & DL
    strData = strData & Format(lvw_Loan.ListItems(C).ListSubItems(6), Precesion) & DL
    strData = strData & "</tr>"
    Total_Rec = Total_Rec + 1
  Next
    strData = strData & "<tr><td colspan=" & lvw_Loan.ColumnHeaders.Count & " align=center bgcolor=pink>" & "Designed and maintained by <a href=mailto:" & EM_MAIL & ">Author</a> Courtesy by Khurram Ayob(engrkhurram@yahoo.com)</td></tr>"
    obj_T.WriteLine strData
    obj_T.WriteLine "</table>"
 Set obj_T = Nothing
 Set obj_File = Nothing
 If Err.Number = 0 Then strMsg = Total_Rec & " Records Converted successfully....."

Exit Sub
EH:
  strMsg = Err.Description
  MsgBox Err.Number & vbCrLf & strMsg, , EM_TITLE
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub img_Print_Click()
' Printer.CurrentX = 1440
' Printer.CurrentY = 1440
' PrintListView lvw_Loan
' Printer.EndDoc
'cdl_File.ShowPrinter
'Caption = cdl_File.Copies
End Sub

Private Sub PrintListView(lvw As ListView)
Const MARGIN = 60
Const COL_MARGIN = 240

Dim ymin As Single
Dim ymax As Single
Dim xmin As Single
Dim xmax As Single
Dim num_cols As Integer
Dim column_header As ColumnHeader
Dim list_item As ListItem
Dim i As Integer
Dim num_subitems As Integer
Dim col_wid() As Single
Dim X As Single
Dim Y As Single
Dim line_hgt As Single
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Printer.Print "Name:-" & Space(15) & txt_Name.Text
Printer.Print "NTN:-" & Space(15) & txt_NTN.Text
Printer.Print "Banker:-" & Space(15) & txt_Banker.Text
Printer.Print "Account #:-" & Space(15) & txt_Account.Text

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    xmin = Printer.CurrentX
    ymin = Printer.CurrentY

    ' ******************
    ' Get column widths.
    num_cols = lvw.ColumnHeaders.Count
    ReDim col_wid(1 To num_cols)

    ' Check the column headers.
    For i = 1 To num_cols
        col_wid(i) = _
            Printer.TextWidth(lvw.ColumnHeaders(i).Text)
    Next i

    ' Check the items.
    num_subitems = num_cols - 1
    For Each list_item In lvw.ListItems
        ' Check the item.
        If col_wid(1) < Printer.TextWidth(list_item.Text) _
            Then _
           col_wid(1) = Printer.TextWidth(list_item.Text)

        ' Check the subitems.
        For i = 1 To num_subitems
            If col_wid(i + 1) < _
                Printer.TextWidth(list_item.SubItems(i)) _
                Then _
               col_wid(i + 1) = _
                   Printer.TextWidth(list_item.SubItems(i))
        Next i
    Next list_item

    ' Add a column margin.
    For i = 1 To num_cols
        col_wid(i) = col_wid(i) + COL_MARGIN
    Next i

    ' *************************
    ' Print the column headers.
    Printer.CurrentY = ymin + MARGIN
    Printer.CurrentX = xmin + MARGIN
    X = xmin + MARGIN
    For i = 1 To num_cols
        Printer.CurrentX = X
        Printer.Print FittedText( _
            lvw.ColumnHeaders(i).Text, col_wid(i));
        X = X + col_wid(i)
    Next i
    xmax = X + MARGIN

    Printer.Print
    line_hgt = Printer.TextHeight("X")
    Y = Printer.CurrentY + line_hgt / 2
    Printer.Line (xmin, Y)-(xmax, Y)
    Y = Y + line_hgt / 2

    ' Print the rows.
    num_subitems = num_cols - 1
    For Each list_item In lvw.ListItems
        X = xmin + MARGIN

        ' Print the item.
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print FittedText( _
            list_item.Text, col_wid(1));
        X = X + col_wid(1)

        ' Print the subitems.
        For i = 1 To num_subitems
            Printer.CurrentX = X
            Printer.Print FittedText( _
                list_item.SubItems(i), col_wid(i + 1));
            X = X + col_wid(i + 1)
        Next i

        Y = Y + line_hgt * 1.5
    Next list_item
    ymax = Y

    ' Draw lines around it all.
    Printer.Line (xmin, ymin)-(xmax, ymax), , B

    X = xmin + MARGIN / 2
    For i = 1 To num_cols - 1
        X = X + col_wid(i)
        Printer.Line (X, ymin)-(X, ymax)
    Next i
End Sub

' Return as much text as will fit in this width.
Private Function FittedText(ByVal txt As String, ByVal Wid _
    As Single) As String
    Do While Printer.TextWidth(txt) > Wid
        txt = Left$(txt, Len(txt) - 1)
    Loop
    FittedText = txt
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


