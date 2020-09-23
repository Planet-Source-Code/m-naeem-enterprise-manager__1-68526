VERSION 5.00
Begin VB.Form frm_Vocabulary_Edit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   Icon            =   "frm_Vocabulary_Edit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Update 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   3
      Top             =   2010
      Width           =   3015
   End
   Begin VB.CommandButton cmd_Exit 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   195
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Width           =   75
   End
   Begin VB.TextBox txt_Meaning 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txt_Word 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3015
   End
End
Attribute VB_Name = "frm_Vocabulary_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''''''''''''''''''''''''''''
Dim strWord As String

Private Sub cmd_Exit_Click()
 Unload Me
End Sub

Private Sub cmd_Update_Click()
 Dim R1 As New ADODB.Recordset
 ''''''''''''''''
 strSql = "select * from vocabulary where word ='" & strWord & " '"
 
 With R1
  .Open strSql, Con, adOpenDynamic, adLockPessimistic
  .Fields(1).Value = Trim(txt_Word.Text)
  .Fields(2).Value = Trim(txt_Meaning.Text)
 .Update
 .Close
 End With
 '''''''''''''''''''''''
 strMsg = txt_Word.Text & " updated "
 Unload Me
 frm_Vocabulary_Master.txt_Display.Text = Err.Number & Space(2) & strMsg
End Sub

Private Sub Form_Activate()
 strWord = Trim(txt_Word)
End Sub

Private Sub Form_Load()
 Me.Caption = EM_TITLE & "..Edit" & EM_MAIL
End Sub

Private Sub txt_Meaning_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  cmd_Update.Value = True
 End If
End Sub
