VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_HTML_Wizard 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   Icon            =   "frm_HTML_Wizard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgLst_HTML 
      Left            =   5610
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   250
      ImageHeight     =   184
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_HTML_Wizard.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_HTML_Wizard.frx":21E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_HTML_Wizard.frx":43AEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra_Go 
      Caption         =   "Write HTML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   0
      TabIndex        =   22
      Top             =   3480
      Width           =   5475
      Begin VB.PictureBox pic_WB 
         BackColor       =   &H00FFFFFF&
         Height          =   2835
         Left            =   1170
         ScaleHeight     =   2775
         ScaleWidth      =   4185
         TabIndex        =   25
         Top             =   150
         Width           =   4245
         Begin SHDocVwCtl.WebBrowser WB 
            Height          =   2835
            Left            =   -30
            TabIndex        =   26
            ToolTipText     =   "Preview Area"
            Top             =   0
            Visible         =   0   'False
            Width           =   4215
            ExtentX         =   7435
            ExtentY         =   5001
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
      Begin VB.CheckBox chk_Preview 
         Caption         =   "Preview"
         Height          =   405
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Show Preview"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmd_Next 
         DownPicture     =   "frm_HTML_Wizard.frx":43C50
         Height          =   555
         Left            =   300
         Picture         =   "frm_HTML_Wizard.frx":43DA2
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Dowload File as HTML"
         Top             =   2070
         Width           =   555
      End
   End
   Begin VB.Frame fra_Style 
      Caption         =   "Choose Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   5610
      TabIndex        =   19
      Top             =   120
      Width           =   5475
      Begin VB.OptionButton opt_Style 
         Caption         =   "Tabular"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   1380
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton opt_Style 
         Caption         =   "Columnar"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   390
         Width           =   1065
      End
      Begin VB.Image img_Wizard 
         Height          =   2775
         Left            =   1290
         ToolTipText     =   "Style Preview"
         Top             =   180
         Width           =   4125
      End
   End
   Begin VB.Frame fra_Fields 
      Caption         =   "Fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   5610
      TabIndex        =   12
      Top             =   3300
      Width           =   5475
      Begin VB.CheckBox chk_Fields 
         Caption         =   "Comments"
         Height          =   195
         Index           =   5
         Left            =   270
         TabIndex        =   28
         Top             =   2640
         Width           =   1905
      End
      Begin VB.CheckBox chk_Fields 
         Caption         =   "Address Office"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   18
         Top             =   2250
         Width           =   1905
      End
      Begin VB.CheckBox chk_Fields 
         Caption         =   "Address Residence"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   17
         Top             =   1860
         Width           =   1905
      End
      Begin VB.CheckBox chk_Fields 
         Caption         =   "Email"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   16
         Top             =   1470
         Width           =   1905
      End
      Begin VB.CheckBox chk_Fields 
         Caption         =   "Telephone Office"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   15
         Top             =   1080
         Width           =   1905
      End
      Begin VB.CheckBox chk_Fields 
         Caption         =   "Telephone Residence"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   14
         Top             =   690
         Width           =   1905
      End
      Begin VB.OptionButton opt_Name 
         Caption         =   "Name"
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   270
         Value           =   -1  'True
         Width           =   2865
      End
   End
   Begin MSComctlLib.TabStrip TS_HTML 
      Height          =   3405
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   6006
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Style"
            Key             =   "Style"
            Object.Tag             =   "Style"
            Object.ToolTipText     =   "Choose Style"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Font"
            Key             =   "Font"
            Object.Tag             =   "Font"
            Object.ToolTipText     =   "Choose Font"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fields"
            Key             =   "Fields"
            Object.Tag             =   "Fields"
            Object.ToolTipText     =   "Choose Fields"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Go"
            Key             =   "Go"
            Object.Tag             =   "Go"
            Object.ToolTipText     =   "Develop the page"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra_Font 
      Caption         =   "Font && Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   0
      TabIndex        =   0
      Top             =   6630
      Width           =   5475
      Begin VB.ComboBox cbo_Border_Size 
         Height          =   315
         ItemData        =   "frm_HTML_Wizard.frx":43EC0
         Left            =   1290
         List            =   "frm_HTML_Wizard.frx":43EF1
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1020
         Width           =   705
      End
      Begin VB.CommandButton cmd_Border_Color 
         Caption         =   "..."
         Height          =   225
         Left            =   1290
         TabIndex        =   7
         ToolTipText     =   "Border Color"
         Top             =   660
         Width           =   255
      End
      Begin VB.CommandButton cmd_Font_Color 
         Caption         =   "..."
         Height          =   225
         Left            =   1290
         TabIndex        =   3
         ToolTipText     =   "Font Color"
         Top             =   270
         Width           =   255
      End
      Begin VB.CommandButton cmd_Font 
         Caption         =   "..."
         Height          =   225
         Left            =   1260
         TabIndex        =   2
         ToolTipText     =   "Font"
         Top             =   1470
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Border Size"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Border Color"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lbl_Border_Color 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   1740
         TabIndex        =   8
         Top             =   660
         Width           =   225
      End
      Begin VB.Label lbl_Font 
         Caption         =   "AaBbCcDd123"
         ForeColor       =   &H00008000&
         Height          =   735
         Left            =   1830
         TabIndex        =   6
         ToolTipText     =   "Font Preview"
         Top             =   1500
         Width           =   3360
      End
      Begin VB.Label lbl_Color 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   225
         Left            =   1740
         TabIndex        =   5
         Top             =   270
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Font Color"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Font"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   1470
         Width           =   315
      End
   End
   Begin MSComDlg.CommonDialog cdl_File 
      Left            =   5670
      Top             =   7620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".htm"
      DialogTitle     =   "Convert Data"
      Filter          =   "HTML File|*.htm;*.html|Text File|*.txt|All Files|*.htm;*.html;*.txt"
   End
End
Attribute VB_Name = "frm_HTML_Wizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''   Frame Dimensions   '''''''''''''''''''
 Const fra_Left = 0
 Const fra_Top = 330
 Const fra_Width = 5475
 Const fra_Height = 3045
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim PVW As String

Dim strData As String
Dim obj_File As New FileSystemObject, obj_T As TextStream
Dim strFile As String
Dim bol_Style As Boolean

Dim DL As String  '' delimiter ....
Dim Head As String

Dim Border_Color As OLE_COLOR ' ColorConstants
Dim Border_Size As Long
Dim Font_Color As OLE_COLOR ' ColorConstants
Dim Font_Name As String
Dim Font_Bold As String
Dim Font_Italic As String
Dim Font_Size  As Long
Dim Font_Strikethru As String
Dim Font_Underline  As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Address_Office As Boolean
Dim Address_Residence As Boolean
Dim Email As Boolean
Dim Tel_Office As Boolean
Dim Tel_Residence As Boolean
Dim Comments As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cbo_Border_Size_Change()
 Border_Size = Val(cbo_Border_Size.Text)
End Sub

Private Sub cbo_Border_Size_Click()
 Border_Size = Val(cbo_Border_Size.Text)
End Sub

Private Sub cbo_Border_Size_Scroll()
 Border_Size = Val(cbo_Border_Size.Text)
End Sub

Private Sub chk_Fields_Click(Index As Integer)
 Select Case Index
  Case 0:   Tel_Residence = Not Tel_Residence
  Case 1:   Tel_Office = Not Tel_Office
  Case 2:   Email = Not Email
  Case 3:   Address_Residence = Not Address_Residence
  Case 4:   Address_Office = Not Address_Office
  Case 5:   Comments = Not Comments
 End Select
End Sub

Private Sub cmd_Border_Color_Click()
 cdl_File.ShowColor
 Border_Color = cdl_File.Color
 lbl_Border_Color.BackColor = Border_Color
End Sub

Private Sub cmd_Font_Click()
 cdl_File.Flags = cdlCFBoth
 cdl_File.ShowFont
  
 Font_Name = cdl_File.FontName
 If Trim(Font_Name) = vbNullString Then
  Font_Name = "Arial"
 End If
 Font_Bold = cdl_File.FontBold
 Font_Italic = cdl_File.FontItalic
 Font_Size = cdl_File.FontSize
 Font_Strikethru = cdl_File.FontStrikethru
 Font_Underline = cdl_File.FontUnderline
 
 lbl_Font.FontName = Font_Name
 lbl_Font.FontBold = Font_Bold
 lbl_Font.FontItalic = Font_Italic
 lbl_Font.FontSize = Font_Size
 lbl_Font.FontStrikethru = Font_Strikethru
 lbl_Font.FontUnderline = Font_Underline
 
 lbl_Font.ForeColor = Font_Color
 
End Sub

Private Sub cmd_Font_Color_Click()
 cdl_File.ShowColor
 Font_Color = cdl_File.Color
 lbl_Color.BackColor = Font_Color
 lbl_Font.ForeColor = Font_Color
End Sub

Private Sub cmd_Next_Click()
 On Error GoTo EH
  If Convert Then
   ShellExecute Me.hWnd, vbNullString, strFile, vbNullString, "C:\", SW_NORMAL
  End If
  strFile = vbNullString
 Exit Sub
EH:
 MsgBox Err.Number & vbCrLf & Err.Description, , EM_TITLE
End Sub

Private Sub chk_Preview_Click()
 
 Select Case chk_Preview.Value
  Case vbChecked
   Convert PVW
   WB.Navigate PVW
   WB.Visible = True
   
  Case vbUnchecked
   WB.Visible = False
   
  End Select
  
End Sub

Private Sub Form_Load()
 Dim Ctrl As Variant
 frm_HTML_Wizard.Width = fra_Width + 125
 frm_HTML_Wizard.Height = fra_Height + 700
 Me.Caption = "HTML Convertion WIZARD..." & EM_MAIL
 PVW = App.Path & "\preview.htm"
 
 For Each Ctrl In Controls
   If TypeOf Ctrl Is Frame Then
    Ctrl.Left = fra_Left
    Ctrl.Top = fra_Top
    Ctrl.Width = fra_Width
    Ctrl.Height = fra_Height
   End If
 Next
 opt_Style_Click (1)
 fra_Style.ZOrder 0
 
End Sub

Private Sub opt_Style_Click(Index As Integer)
 img_Wizard.picture = imgLst_HTML.ListImages(Index + 1).picture
 bol_Style = Not bol_Style
End Sub


Private Function Convert(Optional fileName As String) As Boolean
 Dim R1 As New ADODB.Recordset
 Dim Total_Rec As Long
On Error GoTo EH
''''''''''''''''''''''
If Trim(AB_SQL) = vbNullString Then Exit Function
''''''''''''''''''''''
If fileName = vbNullString Then
 cdl_File.ShowSave
 strFile = cdl_File.fileName
 cdl_File.fileName = vbNullString
 If Len(strFile) < 1 Then
  Convert = False
  Exit Function
 End If
Else
 strFile = fileName
End If

 If bol_Style = False Then
  DL = "<br>"
  Head = "<tr bgcolor=silver><td>Name"
  If Tel_Residence = True Then Head = Head & "<br>Tel(Res)"
  If Tel_Office = True Then Head = Head & "<br>Tel(Off)"
  If Email = True Then Head = Head & "<br>Email"
  If Address_Residence = True Then Head = Head & "<br>Address (Res)"
  If Address_Office = True Then Head = Head & "<br>Address (Off)"
  If Comments = True Then Head = Head & "<br>Comments"
  Head = Head & "</td></tr>"
 Else
  DL = "</td><td>"
  Head = "<tr bgcolor=teal><th>Name</th>"
  If Tel_Residence = True Then Head = Head & "<th>Tel(Res)</th>"
  If Tel_Office = True Then Head = Head & "<th>Tel(Off)</th>"
  If Email = True Then Head = Head & "<th>Email</th>"
  If Address_Residence = True Then Head = Head & "<th>Address (Res)</th>"
  If Address_Office = True Then Head = Head & "<th>Address (Off)</th>"
  If Comments = True Then Head = Head & "<th>Comments</th>"
  Head = Head & "</tr>"
 End If

  Set obj_T = obj_File.OpenTextFile(strFile, ForWriting, True)
  strData = "<body text=#" & Resolve_Color(Font_Color) & "#>"
  strData = strData & "<title>Address Book by" & EM_MAIL & "</title>"
  strData = strData & "<table align=center border=" & Border_Size & " bordercolor=#" & Resolve_Color(Border_Color) & "#>"
  strData = strData & "<caption><h2>Address List</h2></caption>"
  strData = strData & Head
  obj_T.WriteLine strData
  
  ''strSql = "Select * from address order by name"
  strSql = AB_SQL
 
 With R1
 .Open strSql, Con
  While Not .EOF
    strData = "<tr><td>" & .Fields("name") & DL
    
  If Tel_Residence = True Then
    If IsNull(.Fields("telephone_residence")) Then
      strData = strData & "&nbsp;" & DL
    Else
      strData = strData & .Fields("telephone_residence") & DL
    End If
  End If
          '''''''''''''''''''''''''''''''''''''''''''
  If Tel_Office = True Then
    If IsNull(.Fields("telephone_office")) Then
      strData = strData & "&nbsp;" & DL
    Else
       strData = strData & .Fields("telephone_office") & DL
    End If
   End If
          '''''''''''''''''''''''''''''''''''''''''''
   If Email = True Then
    If IsNull(.Fields("email")) Then
      strData = strData & "&nbsp;" & DL
    Else
      strData = strData & .Fields("email") & DL
    End If
   End If
          '''''''''''''''''''''''''''''''''''''''''''
   If Address_Residence = True Then
     If IsNull(.Fields("address_residence")) Then
      strData = strData & "&nbsp;" & DL
    Else
       strData = strData & .Fields("address_residence") & DL
    End If
   End If
          '''''''''''''''''''''''''''''''''''''''''''
   If Address_Office = True Then
    If IsNull(.Fields("address_office")) Then
      strData = strData & "&nbsp;</td><td>"
    Else
      strData = strData & .Fields("address_office") & "</td></tr>"
    End If
   End If
          '''''''''''''''''''''''''''''''''''''''''''
   If Comments = True Then
    If IsNull(.Fields("comments")) Then
      strData = strData & "&nbsp;</td><td>"
    Else
      strData = strData & .Fields("comments") & "</td></tr>"
    End If
   End If
          '''''''''''''''''''''''''''''''''''''''''''
    obj_T.WriteLine strData
    .MoveNext
    Total_Rec = Total_Rec + 1
  Wend
    strData = "<tr><td colspan=6 align=center bgcolor=pink>" & Total_Rec & " Records  ................  Designed and maintained by <a href=mailto:" & EM_MAIL & ">Author</a></td></tr>"
    obj_T.WriteLine strData
 .Close
 End With
 obj_T.WriteLine "</table>"
 Set obj_T = Nothing
 Set obj_File = Nothing
 ''strMsg = Total_Rec & " Records Converted successfully....." '' no need infact as the IE itself invoked ...by this program
 Convert = True
Exit Function
EH:
  MsgBox Err.Number & vbCrLf & Err.Description, , EM_TITLE
End Function

Private Function Resolve_Color(ByVal Color As OLE_COLOR) As String
 Dim R As Byte, G As Byte, B As Byte
 R = Color And &HFF&
 G = (Color And &HFF00&) \ &H100&
 B = (Color And &HFF0000) \ &H10000
 Resolve_Color = R & G & B
End Function

Private Sub TS_HTML_Click()
 Select Case TS_HTML.SelectedItem.Index
  Case 1: fra_Style.ZOrder 0
  Case 2: fra_Font.ZOrder 0
  Case 3: fra_Fields.ZOrder 0
  Case 4:
        chk_Preview.Value = vbUnchecked
        fra_Go.ZOrder 0
 End Select
End Sub

Private Sub WB_DocumentComplete(ByVal pDisp As Object, url As Variant)
 On Error GoTo EH
 If obj_File.FileExists(PVW) Then
  obj_File.DeleteFile PVW, True
 End If
 Exit Sub
EH:
 MsgBox Err.Number & vbCrLf & Err.Description
End Sub

