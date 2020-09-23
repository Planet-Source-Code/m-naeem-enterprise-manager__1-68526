VERSION 5.00
Begin VB.Form frm_Address_Book_Action 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frm_Address_Book_Action.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1185
      Index           =   1
      Left            =   2460
      TabIndex        =   19
      Top             =   3690
      Width           =   225
      Begin VB.Image img_Arrow 
         Height          =   240
         Index           =   11
         Left            =   0
         Picture         =   "frm_Address_Book_Action.frx":030A
         Top             =   930
         Width           =   240
      End
      Begin VB.Image img_Arrow 
         Height          =   240
         Index           =   10
         Left            =   0
         Picture         =   "frm_Address_Book_Action.frx":0454
         Top             =   630
         Width           =   240
      End
      Begin VB.Image img_Arrow 
         Height          =   240
         Index           =   9
         Left            =   0
         Picture         =   "frm_Address_Book_Action.frx":059E
         Top             =   330
         Width           =   240
      End
      Begin VB.Image img_Arrow 
         Height          =   240
         Index           =   8
         Left            =   0
         Picture         =   "frm_Address_Book_Action.frx":06E8
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.ListBox lst_Group_Select 
      Height          =   1230
      ItemData        =   "frm_Address_Book_Action.frx":0832
      Left            =   2940
      List            =   "frm_Address_Book_Action.frx":0834
      MultiSelect     =   2  'Extended
      TabIndex        =   9
      Top             =   3660
      Width           =   2235
   End
   Begin VB.ListBox lst_Group_All 
      Height          =   1230
      ItemData        =   "frm_Address_Book_Action.frx":0836
      Left            =   0
      List            =   "frm_Address_Book_Action.frx":0838
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   3660
      Width           =   2175
   End
   Begin VB.CommandButton cmd_Action 
      Caption         =   "Add New Address"
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
      Left            =   1470
      TabIndex        =   10
      Top             =   4950
      Width           =   2385
   End
   Begin VB.Frame fra_Action 
      Caption         =   "Add New Address:-"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txt_Comments 
         Height          =   495
         Left            =   1170
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2850
         Width           =   3885
      End
      Begin VB.TextBox txt_Name 
         Height          =   285
         Left            =   720
         MaxLength       =   200
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txt_Address_Off 
         Height          =   495
         Left            =   1170
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2280
         Width           =   3885
      End
      Begin VB.TextBox txt_Address_Res 
         Height          =   495
         Left            =   1170
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1680
         Width           =   3885
      End
      Begin VB.TextBox txt_Email 
         Height          =   285
         Left            =   720
         MaxLength       =   200
         TabIndex        =   4
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txt_Tel_Off 
         Height          =   285
         Left            =   720
         MaxLength       =   200
         TabIndex        =   3
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txt_Tel_Res 
         Height          =   285
         Left            =   720
         MaxLength       =   200
         TabIndex        =   2
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label lbl_Detail 
         AutoSize        =   -1  'True
         Caption         =   "Commenst"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lbl_Detail 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lbl_Detail 
         AutoSize        =   -1  'True
         Caption         =   "Address(Off)"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   870
      End
      Begin VB.Label lbl_Detail 
         AutoSize        =   -1  'True
         Caption         =   "Address (Res)"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label lbl_Detail 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lbl_Detail 
         AutoSize        =   -1  'True
         Caption         =   "Tel (O)"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   480
      End
      Begin VB.Label lbl_Detail 
         AutoSize        =   -1  'True
         Caption         =   "Tel (R)"
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   480
      End
   End
   Begin VB.Label lbl_Display 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   17
      Top             =   5340
      Width           =   5145
   End
End
Attribute VB_Name = "frm_Address_Book_Action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''''''''''''''''''''''''''
Dim AB_Group() As String
Dim AB_Group_Select() As String

Private Sub cmd_Action_Click()
 On Error GoTo EH
 '''''''''''''''''''
 If Trim(txt_Name.Text) = vbNullString Then
  lbl_Display.Caption = "Name is mandatory field"
  Exit Sub
 End If
 '''''''''''''''''''
 Select Case Action_Mode
 Case "Add"
  Add_Address
  strMsg = txt_Name.Text
   txt_Name.Text = vbNullString
   txt_Tel_Res.Text = vbNullString
   txt_Tel_Off.Text = vbNullString
   txt_Email.Text = vbNullString
   txt_Address_Res.Text = vbNullString
   txt_Address_Off.Text = vbNullString
   txt_Comments.Text = vbNullString
 
   lbl_Display.Caption = "Successfully added:- " & strMsg

 Case "Update"
 Update_Address
 frm_Address_Book.lbl_Display.Caption = txt_Name.Text & " updated "
 Unload Me
 
 End Select
 Erase AB_Group_Select
 Exit Sub
EH:
 lbl_Display.Caption = Err.Description
End Sub

Private Sub Update_Group_List(F As Form)
 Dim i  As Integer, J As Long
 '''''''''''''''''''''''''
 If F.lst_Group.ListCount = 0 Then Exit Sub
 For i = 0 To F.lst_Group.ListCount - 1
  
  lst_Group_Select.AddItem F.lst_Group.List(i)
  For J = 0 To lst_Group_All.ListCount
   If lst_Group_All.List(J) = F.lst_Group.List(i) Then
    lst_Group_All.RemoveItem (J)
    Exit For
   End If
  Next
 Next

End Sub

Private Function Get_Group_ID() As Boolean
Dim i As Long, J As Long
Dim A As Variant
On Error GoTo EH

ReDim AB_Group_Select(lst_Group_Select.ListCount - 1)
For i = 0 To lst_Group_Select.ListCount - 1
 
 For J = 0 To UBound(AB_Group) - 1
  A = Split(AB_Group(J), "---")
  If A(0) = lst_Group_Select.List(i) Then
   AB_Group_Select(i) = A(1)
   Exit For
  End If
 Next
 
Next
Get_Group_ID = True
Exit Function
EH:
Get_Group_ID = False
End Function


Private Sub Form_Load()

strMsg = vbNullString ''' must be at the top statement in load event
Me.Caption = EM_TITLE & ".." & "Address Book" & EM_MAIL
Populate_Group


Select Case Action_Mode
Case "Add":

Case "Update":

Me.Caption = "Edit Address " & EM_MAIL
fra_Action.Caption = "Edit Address"
cmd_Action.Caption = "Update Address"

With frm_Address_Book
 txt_Name.Text = .lst_Address.Text
 txt_Tel_Res.Text = .lbl_Tel_Res.Caption
 txt_Tel_Off.Text = .lbl_Tel_Off.Caption
 txt_Email.Text = .lbl_Email.Caption
 txt_Address_Res.Text = .lbl_Address_Res.Caption
 txt_Address_Off.Text = .lbl_Address_Off.Caption
 txt_Comments.Text = .lbl_Comments.Caption
 
 Update_Group_List frm_Address_Book
End With

End Select



End Sub

Private Sub Populate_Group()
 Dim R1 As New ADODB.Recordset
 Dim i As Long
 ''''''''''''''''''
 lst_Group_All.Clear
 strSql = "select * from address_book_group order by name"
 With R1
  .Open strSql, Con, 3, 3
  ReDim AB_Group(.RecordCount)
 While Not .EOF
   lst_Group_All.AddItem .Fields(1)
   AB_Group(i) = .Fields(1) & "---" & .Fields(0)
   i = i + 1
   .MoveNext
 Wend
 .Close
 End With
End Sub

Private Sub Add_Address()
 Dim R1 As New ADODB.Recordset
 Dim P_ID As Long
 ''''''''''''''''''''''''
 strSql = "select * from address"
 With R1
  .Open strSql, Con, adOpenDynamic, adLockPessimistic
  .AddNew
   .Fields("name") = txt_Name.Text
   .Fields("telephone_residence") = txt_Tel_Res.Text + " "
   .Fields("telephone_office") = txt_Tel_Off.Text + " "
   .Fields("email") = Trim(txt_Email.Text) + " "
   .Fields("address_residence") = txt_Address_Res.Text + " "
   .Fields("address_office") = txt_Address_Off.Text + " "
   .Fields("comments") = txt_Comments.Text + " "
 .Update
 .Close
 End With
 ''''''''''''''''''''''''''''''''''''''''
 strSql = "select id from address order by id desc"
 With R1
  .Open strSql, Con
   P_ID = R1.Fields(0)
 .Close
 End With
 ''''''''''''''''''''''''''''''''''''''''
 Add_Group_ID_in_DB P_ID
End Sub

Private Sub Add_Group_ID_in_DB(Person_ID)
 Dim R1 As New ADODB.Recordset
 Dim i As Long
 ''''''''''''
If Get_Group_ID = True Then
 strSql = "select * from address_book_group_link"
 With R1
  .Open strSql, Con, adOpenDynamic, adLockPessimistic
  For i = 0 To UBound(AB_Group_Select)
  .AddNew
   .Fields("person_id") = Person_ID
   .Fields("group_id") = Val(AB_Group_Select(i))
 .Update
  Next
 .Close
 End With
 End If
End Sub


Private Sub Update_Address()
Dim R1 As New ADODB.Recordset
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 strSql = "select * from address where id = " & Edit_ID
 With R1
  .Open strSql, Con, adOpenDynamic, adLockPessimistic
   .Fields("name") = txt_Name.Text
   .Fields("telephone_residence") = txt_Tel_Res.Text + " "
   .Fields("telephone_office") = txt_Tel_Off.Text + " "
   .Fields("email") = Trim(txt_Email.Text) + " "
   .Fields("address_residence") = txt_Address_Res.Text + " "
   .Fields("address_office") = txt_Address_Off.Text + " "
   .Fields("comments") = txt_Comments.Text + " "
 .Update
 .Close
 End With
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 strSql = "delete from address_book_group_link where person_id=" & Edit_ID
 Con.Execute strSql
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Add_Group_ID_in_DB Edit_ID
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub lst_Group_All_DblClick()
 lst_Group_Select.AddItem lst_Group_All.Text
 lst_Group_All.RemoveItem (lst_Group_All.ListIndex)
End Sub
Private Sub lst_Group_select_DblClick()
 lst_Group_All.AddItem lst_Group_Select.Text
 lst_Group_Select.RemoveItem (lst_Group_Select.ListIndex)
End Sub

