VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_About 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Spy....."
   ClientHeight    =   3465
   ClientLeft      =   3630
   ClientTop       =   675
   ClientWidth     =   5745
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frm_About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   231
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList EM_ImageList 
      Left            =   420
      Top             =   2850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   109
      ImageHeight     =   154
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_About.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_About.frx":675D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H8000000B&
      Caption         =   "OK"
      Height          =   315
      Left            =   4920
      TabIndex        =   1
      Top             =   3090
      Width           =   765
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frm_About.frx":A3F4
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Timer Timer2 
      Interval        =   4
      Left            =   6120
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   120
   End
   Begin VB.Label lbl_Version 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1950
      TabIndex        =   0
      Top             =   2070
      Width           =   3705
   End
   Begin VB.Image imgAuthor 
      Height          =   2310
      Left            =   120
      ToolTipText     =   "Muhammad Naeem (naeem@email.com)"
      Top             =   360
      Width           =   1635
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Muhammad Naeem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "naeem@email.com"
      Top             =   3060
      Width           =   2010
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designed && Maintained by:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   0
      Left            =   1410
      TabIndex        =   4
      Top             =   2700
      Width           =   2985
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enterprise Manager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   3555
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim flgAuthor As Boolean, flgImgAuthor As Boolean
Dim lngCounter As Long

Private Sub cmdOK_Click()
 Timer1.Enabled = False: Timer2.Enabled = False
 Unload Me
 Set frm_About = Nothing
End Sub


Private Sub Form_Load()
'' code to draw picture and tool tip text of the author's  04 (or whatever) pix
'Dim i As Integer
'Randomize Timer
'i = Int((7 - 4 + 1) * Rnd() + 4)
'imgAuthor.ToolTipText = imgLst.ListImages(i).Tag
 imgAuthor.picture = EM_ImageList.ListImages(1).picture
'imgAuthor.Picture = LoadPicture(App.Path & "\author.gif")

   ''''''''' start some initialization for the Blazing code
 maxx = Label1.Width                          'get label width
 maxy = Label1.Height + (Label1.Height / 2)   'get label height add extra height for flame
 ReDim new_flame(maxx, maxy)                  'resize array to label
 ReDim old_flame(maxx, maxy)
   
 lbl_Version.Caption = "Version:- " & App.Major & ":" & App.Minor & ":" & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If flgAuthor = True Then
    flgAuthor = False
    lblAuthor(0).FontItalic = False
    lblAuthor(0).ForeColor = &HFFC0C0
 End If
 If flgImgAuthor = True Then
    flgImgAuthor = False
    imgAuthor.picture = EM_ImageList.ListImages(1).picture
 End If
End Sub

Private Sub imgAuthor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If flgImgAuthor = False Then
  flgImgAuthor = True
  imgAuthor.picture = EM_ImageList.ListImages(2).picture
 End If
End Sub

Private Sub lblAuthor_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 lblAuthor(Index).FontItalic = True
 lblAuthor(Index).ForeColor = vbYellow
 flgAuthor = True
End Sub

Private Sub Timer1_Timer()
  'This is the main timer,  Displays and updates the flame
  Dim X, Y As Integer    'store current x and y pos.
  Dim red, green, blue As Long     'store colours
  Dim Tmp As Long
If lngCounter > 20 Then Timer1.Enabled = False: Timer2.Enabled = False: Exit Sub
  'This part generates the flame :)
    For X = 1 To maxx - 1
     For Y = 1 To maxy - 1
       'Add up the surrounding red colours
        red = new_flame(X + 1, Y).R
        red = red + new_flame(X - 1, Y).R
        red = red + new_flame(X, Y + 1).R
        red = red + new_flame(X, Y - 1).R
            DoEvents
        'Add up the surrounding green colours
        green = new_flame(X + 1, Y).G
        green = green + new_flame(X - 1, Y).G
        green = green + new_flame(X, Y + 1).G
        green = green + new_flame(X, Y - 1).G
             DoEvents
'        blue = blue + new_flame(X + 1, Y).b    'Add up the surrounding blue colours
'        blue = blue + new_flame(X - 1, Y).b
'        blue = blue + new_flame(X, Y + 1).b
'        blue = blue + new_flame(X, Y - 1).b
        
        'uses the row above (y-1) to give the effect of moving up!
        If old_flame(X, Y - 1).C = False Then   'if pixel is part of flame update
          Tmp = (Rnd * Flame_Height)                      'pick a number from the air!
          old_flame(X, Y - 1).R = red / 4 - (Tmp) ' Average the red and decrease the colour
          old_flame(X, Y - 1).G = (green / 4) - (Tmp + 8) ' Average the green and decrease the colour
             
'         old_flame(X, Y - 1).b = blue / 4 ' Average the blue
          'Check colours haven`t gone below 0
          If old_flame(X, Y - 1).R < 0 Then old_flame(X, Y - 1).R = 0
          If old_flame(X, Y - 1).G < 0 Then old_flame(X, Y - 1).G = 0
'          If old_flame(X, Y - 1).b < 0 Then old_flame(X, Y - 1).b = 0
        End If
     Next Y
  Next X
  
  'This loop Displays and updates the array
  For X = 1 To maxx
     For Y = 1 To maxy
        new_flame(X, Y).R = old_flame(X, Y).R     ' update array
        new_flame(X, Y).G = old_flame(X, Y).G
   '  new_flame(X, Y).b = old_flame(X, Y).b
        'put the pixel!
        DoEvents
 ' Me.PSet (Label1.Left + X, Label1.Top + Y - Int(Label1.Height / 2)), RGB(new_flame(X - 1, Y).r, new_flame(X - 1, Y).g, new_flame(X - 1, Y).b)
Me.PSet (Label1.Left + X, Label1.Top + Y - Int(Label1.Height / 2)), RGB(new_flame(X - 1, Y).R, new_flame(X - 1, Y).G, new_flame(X - 1, Y).B)
     Next Y
  Next X
  lngCounter = lngCounter + 1
End Sub

Private Sub Timer2_Timer()
    'This timer only initializes the array colours
    Dim X As Long
    Dim Y As Long
      
    For X = 1 To maxx
     For Y = 1 To maxy
          If Point(Label1.Left + X, Label1.Top + Label1.Height - Y) <> 0 Then ' is there any colour at this point
           new_flame(X, maxy - Y).R = 255   ' Set colour to Yellow
           new_flame(X, maxy - Y).G = 255
           new_flame(X, maxy - Y).B = 0
           new_flame(X, maxy - Y).C = True  ' Is a permenant colour
          Else
           new_flame(X, maxy - Y).R = 0
           new_flame(X, maxy - Y).G = 0
           new_flame(X, maxy - Y).B = 0
           new_flame(X, maxy - Y).C = False ' Can be any colour
          End If
            DoEvents
          old_flame(X, maxy - Y).R = new_flame(X, maxy - Y).R  'old_flame=new_flame
          old_flame(X, maxy - Y).G = new_flame(X, maxy - Y).G
          old_flame(X, maxy - Y).B = new_flame(X, maxy - Y).B
          old_flame(X, maxy - Y).C = new_flame(X, maxy - Y).C
     Next Y
  Next X
  Label1.Visible = False
  Timer1.Enabled = True   ' Call the Fire brigade :)
  Timer2.Enabled = False  ' Turn off the taps!
End Sub
