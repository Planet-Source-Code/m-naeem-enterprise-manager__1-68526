Attribute VB_Name = "Text_BG_Pix"
Option Explicit

Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "User32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDC As Long, ByVal DX As Long, ByVal DY As Long, ByVal DWidth As Long, ByVal DHeight As Long, ByVal ShDC As Long, ByVal SX As Long, ByVal SY As Long, ByVal vbSrCopy As Long) As Long
Public Declare Function BeginPaint Lib "User32" (ByVal hWnd As Long, lPaint As Any) As Long
Public Declare Function EndPaint Lib "User32" (ByVal hWnd As Long, lPaint As Any) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function RedrawWindow Lib "User32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function SetBkMode Lib "GDI32" (ByVal hDC As Long, ByVal hMode As Long) As Long
Public Declare Function SetTextColor Lib "GDI32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function WindowFromDC Lib "User32" (ByVal hDC As Long) As Long
Public Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Const GWL_WNDPROC = -4
Public Const VBM_FIREEVENT = &H100E

Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_CANCELMODE = &H1F
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_CHAR = &H102
Public Const WM_CHARTOITEM = &H2F
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
Public Const WM_CHOOSEFONT_SETFLAGS = (&H400 + 102)
Public Const WM_CHOOSEFONT_SETLOGFONT = (&H400 + 101)
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_COMPACTING = &H41
Public Const WM_COMPAREITEM = &H39
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_CONVERTREQUESTEX = &H108
Public Const WM_COPY = &H301
Public Const WM_COPYDATA = &H4A
Public Const WM_CREATE = &H1
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_CUT = &H300
Public Const WM_DEADCHAR = &H103
Public Const WM_DELETEITEM = &H2D
Public Const WM_DESTROY = &H2
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_DRAWITEM = &H2B
Public Const WM_DROPFILES = &H233
Public Const WM_ENABLE = &HA
Public Const WM_ENDSESSION = &H16
Public Const WM_ENTERIDLE = &H121
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_ERASEBKGND = &H14
Public Const WM_EXITMENULOOP = &H212
Public Const WM_FONTCHANGE = &H1D
Public Const WM_GETDLGCODE = &H87
Public Const WM_GETFONT = &H31
Public Const WM_GETHOTKEY = &H33
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_HOTKEY = &H312
Public Const WM_HSCROLL = &H114
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_KEYUP = &H291
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_INITDIALOG = &H110
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_KEYUP = &H101
Public Const WM_KILLFOCUS = &H8
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICASCADE = &H227
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDINEXT = &H224
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDISETMENU = &H230
Public Const WM_MDITILE = &H226
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUCHAR = &H120
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCHITTEST = &H84
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCPAINT = &H85
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_PAINT = &HF
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_PAINTICON = &H26
Public Const WM_PALETTECHANGED = &H311
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_PASTE = &H302
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_POWER = &H48
Public Const WM_PSD_ENVSTAMPRECT = (&H400 + 5)
Public Const WM_PSD_FULLPAGERECT = (&H400 + 1)
Public Const WM_PSD_GREEKTEXTRECT = (&H400 + 4)
Public Const WM_PSD_MARGINRECT = (&H400 + 3)
Public Const WM_PSD_MINMARGINRECT = (&H400 + 2)
Public Const WM_PSD_PAGESETUPDLG = (&H400)
Public Const WM_PSD_YAFULLPAGERECT = (&H400 + 6)
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_QUERYOPEN = &H13
Public Const WM_QUEUESYNC = &H23
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_RENDERFORMAT = &H305
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFOCUS = &H7
Public Const WM_SETFONT = &H30
Public Const WM_SETHOTKEY = &H32
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_SYSCOMMAND = &H112
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSTIMER = &H118
Public Const WM_TIMECHANGE = &H1E
Public Const WM_TIMER = &H113
Public Const WM_UNDO = &H304
Public Const WM_USER = &H400
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_VSCROLL = &H115
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WININICHANGE = &H1A

Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Public Type POINTAPI
X As Long
Y As Long
End Type

' MEASUREITEMSTRUCT for ownerdraw
Public Type MEASUREITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemWidth As Long
ItemHeight As Long
ItemData As Long
End Type

' DRAWITEMSTRUCT for ownerdraw
Public Type DRAWITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
ItemAction As Long
ItemState As Long
hWndItem As Long
hDC As Long
rcItem As RECT
ItemData As Long
End Type

' DELETEITEMSTRUCT for ownerdraw
Public Type DELETEITEMSTRUCT
CtlType As Long
CtlID As Long
ItemID As Long
hWndItem As Long
ItemData As Long
End Type

' COMPAREITEMSTRUCT for ownerdraw sorting
Public Type COMPAREITEMSTRUCT
CtlType As Long
CtlID As Long
hWndItem As Long
ItemID1 As Long
ItemData1 As Long
ItemID2 As Long
ItemData2 As Long
End Type

'For BeginPaint and EndPaint "erases backcolor"
'i tryed this for the textbox, but didnt work
'or i wasnt using it right
Public Type PAINTSTRUCT
hDC As Long
fErase As Long
fPaint As RECT
fRestore As Long
fIncUpdate As Long
rgbReserved(32) As Byte
End Type

'records mouse move events
Public Type MouseEvents
Button As Integer
Shift As Integer
X As Integer
Y As Integer
End Type

Global FrmPrev As Long
Global TxtPrev As Long
Global TxtColor As Long

Public Sub SubClassFrm(ByVal hWnd As Long)
FrmPrev = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf FrmProc)
End Sub

Public Sub UnHookFrm(ByVal hWnd As Long)
Call SetWindowLong(hWnd, GWL_WNDPROC, FrmPrev)
End Sub

Public Sub SubClassTxt(ByVal hWnd As Long)
TxtPrev = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf TxtProc)
End Sub

Public Sub UnHookTxt(ByVal hWnd As Long)
Call SetWindowLong(hWnd, GWL_WNDPROC, TxtPrev)
End Sub

Private Function FrmProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim Rec As RECT, LP As PAINTSTRUCT, Bu As MouseEvents
'with wm_ctrlcoloredit wParam becomes a hDC
'for all controls that are like a text on the Form
'if you have more then 1 textbox this works for all of them
'but the WindowFromDC uses wParam(hDC) and
'finds the hWnd for the controls
If uMsg = WM_CTLCOLOREDIT Then
If WindowFromDC(wParam) = frmAddress_Book.txt_Tel_Res.hWnd Then
GetWindowRect WindowFromDC(wParam), Rec
SetBkMode wParam, 1
BitBlt wParam, 0, 0, Rec.Right, Rec.Bottom, frmAddress_Book.Picture1.hDC, 0, 0, vbSrcCopy
End If
End If
FrmProc = CallWindowProc(FrmPrev, hWnd, uMsg, wParam, lParam)
End Function

Private Function TxtProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'this just refreshs the textbox
If uMsg = WM_ERASEBKGND Or uMsg = WM_CHAR Then
Call RedrawWindow(hWnd, ByVal 0, 0, &H1)
End If
TxtProc = CallWindowProc(TxtPrev, hWnd, uMsg, wParam, lParam)
End Function

Private Sub FindMsg(uMsg As Long)
'This will help you find out what messages are
'being sent to the control or window
'some of them you have to ignore like WM_MOUSEMOVE
'everytime you move the mouse over the control or window
'the debug window go crazy
Select Case uMsg
Case WM_ACTIVATE: Debug.Print "WM_ACTIVATE"
Case WM_ACTIVATEAPP: Debug.Print "WM_ACTIVATEAPP"
Case WM_ASKCBFORMATNAME: Debug.Print "WM_ASKCBFORMATNAME"
Case WM_CANCELJOURNAL: Debug.Print "WM_CANCELJOURNAL"
Case WM_CANCELMODE: Debug.Print "WM_CANCELMODE"
Case WM_CHANGECBCHAIN: Debug.Print "WM_CHANGECBCHAIN"
Case WM_CHAR: Debug.Print "WM_CHAR"
Case WM_CHARTOITEM: Debug.Print "WM_CHARTOITEM"
Case WM_CHILDACTIVATE: Debug.Print "WM_CHILDACTIVATE"
Case WM_CHOOSEFONT_GETLOGFONT: Debug.Print "WM_CHOOSEFONT_GETLOGFONT"
Case WM_CHOOSEFONT_SETFLAGS: Debug.Print "WM_CHOOSEFONT_SETFLAGS"
Case WM_CHOOSEFONT_SETLOGFONT: Debug.Print "WM_CHOOSEFONT_SETLOGFONT"
Case WM_CLEAR: Debug.Print "WM_CLEAR"
Case WM_CLOSE: Debug.Print "WM_CLOSE"
Case WM_COMMAND: Debug.Print "WM_COMMAND"
Case WM_COMPACTING: Debug.Print "WM_COMPACTING"
Case WM_COMPAREITEM: Debug.Print "WM_COMPAREITEM"
Case WM_CONTEXTMENU: Debug.Print "WM_CONTEXTMENU"
Case WM_CONVERTREQUESTEX: Debug.Print "WM_CONVERTREQUESTEX"
Case WM_COPY: Debug.Print "WM_COPY"
Case WM_COPYDATA: Debug.Print "WM_COPYDATA"
Case WM_CREATE: Debug.Print "WM_CREATE"
Case WM_CTLCOLORBTN: Debug.Print "WM_CTLCOLORBTN"
Case WM_CTLCOLORDLG: Debug.Print "WM_CTLCOLORDLG"
Case WM_CTLCOLOREDIT: Debug.Print "WM_CTLCOLOREDIT"
Case WM_CTLCOLORLISTBOX: Debug.Print "WM_CTLCOLORLISTBOX"
Case WM_CTLCOLORMSGBOX: Debug.Print "WM_CTLCOLORMSGBOX"
Case WM_CTLCOLORSCROLLBAR: Debug.Print "WM_CTLCOLORSCROLLBAR"
Case WM_CTLCOLORSTATIC: Debug.Print "WM_CTLCOLORSTATIC"
Case WM_CUT: Debug.Print "WM_CUT"
Case WM_DEADCHAR: Debug.Print "WM_DEADCHAR"
Case WM_DELETEITEM: Debug.Print "WM_DELETEITEM"
Case WM_DESTROY: Debug.Print "WM_DESTROY"
Case WM_DESTROYCLIPBOARD: Debug.Print "WM_DESTROYCLIPBOARD"
Case WM_DEVMODECHANGE: Debug.Print "WM_DEVMODECHANGE"
Case WM_DRAWCLIPBOARD: Debug.Print "WM_DRAWCLIPBOARD"
Case WM_DRAWITEM: Debug.Print "WM_DRAWITEM"
Case WM_DROPFILES: Debug.Print "WM_DROPFILES"
Case WM_ENABLE: Debug.Print "WM_ENABLE"
Case WM_ENDSESSION: Debug.Print "WM_ENDSESSION"
Case WM_ENTERIDLE: Debug.Print "WM_ENTERIDLE"
Case WM_ENTERMENULOOP: Debug.Print "WM_ENTERMENULOOP"
Case WM_ERASEBKGND: Debug.Print "WM_ERASEBKGND"
Case WM_EXITMENULOOP: Debug.Print "WM_EXITMENULOOP"
Case WM_FONTCHANGE: Debug.Print "WM_FONTCHANGE"
Case WM_GETDLGCODE: Debug.Print "WM_GETDLGCODE"
Case WM_GETFONT: Debug.Print "WM_GETFONT"
Case WM_GETHOTKEY: Debug.Print "WM_GETHOTKEY"
Case WM_GETMINMAXINFO: Debug.Print "WM_GETMINMAXINFO"
Case WM_GETTEXT: Debug.Print "WM_GETTEXT"
Case WM_GETTEXTLENGTH: Debug.Print "WM_GETTEXTLENGTH"
Case WM_HOTKEY: Debug.Print "WM_HOTKEY"
Case WM_HSCROLL: Debug.Print "WM_HSCROLL"
Case WM_HSCROLLCLIPBOARD: Debug.Print "WM_HSCROLLCLIPBOARD"
Case WM_ICONERASEBKGND: Debug.Print "WM_ICONERASEBKGND"
Case WM_IME_CHAR: Debug.Print "WM_IME_CHAR"
Case WM_IME_COMPOSITION: Debug.Print "WM_IME_COMPOSITION"
Case WM_IME_COMPOSITIONFULL: Debug.Print "WM_IME_COMPOSITIONFULL"
Case WM_IME_CONTROL: Debug.Print "WM_IME_CONTROL"
Case WM_IME_ENDCOMPOSITION: Debug.Print "WM_IME_ENDCOMPOSITION"
Case WM_IME_KEYDOWN: Debug.Print "WM_IME_KEYDOWN"
Case WM_IME_KEYLAST: Debug.Print "WM_IME_KEYLAST"
Case WM_IME_KEYUP: Debug.Print "WM_IME_KEYUP"
Case WM_IME_NOTIFY: Debug.Print "WM_IME_NOTIFY"
Case WM_IME_SELECT: Debug.Print "WM_IME_SELECT"
Case WM_IME_SETCONTEXT: Debug.Print "WM_IME_SETCONTEXT"
Case WM_IME_STARTCOMPOSITION: Debug.Print "WM_IME_STARTCOMPOSITION"
Case WM_INITDIALOG: Debug.Print "WM_INITDIALOG"
Case WM_INITMENU: Debug.Print "WM_INITMENU"
Case WM_INITMENUPOPUP: Debug.Print "WM_INITMENUPOPUP"
Case WM_KEYDOWN: Debug.Print "WM_KEYDOWN"
Case WM_KEYFIRST: Debug.Print "WM_KEYFIRST"
Case WM_KEYLAST: Debug.Print "WM_KEYLAST"
Case WM_KEYUP: Debug.Print "WM_KEYUP"
Case WM_KILLFOCUS: Debug.Print "WM_KILLFOCUS"
Case WM_LBUTTONDBLCLK: Debug.Print "WM_LBUTTONDBLCLK"
Case WM_LBUTTONDOWN: Debug.Print "WM_LBUTTONDOWN"
Case WM_LBUTTONUP: Debug.Print "WM_LBUTTONUP"
Case WM_MBUTTONDBLCLK: Debug.Print "WM_MBUTTONDBLCLK"
Case WM_MBUTTONDOWN: Debug.Print "WM_MBUTTONDOWN"
Case WM_MBUTTONUP: Debug.Print "WM_MBUTTONUP"
Case WM_MDIACTIVATE: Debug.Print "WM_MDIACTIVATE"
Case WM_MDICASCADE: Debug.Print "WM_MDICASCADE"
Case WM_MDICREATE: Debug.Print "WM_MDICREATE"
Case WM_MDIDESTROY: Debug.Print "WM_MDIDESTROY"
Case WM_MDIGETACTIVE: Debug.Print "WM_MDIGETACTIVE"
Case WM_MDIICONARRANGE: Debug.Print "WM_MDIICONARRANGE"
Case WM_MDIMAXIMIZE: Debug.Print "WM_MDIMAXIMIZE"
Case WM_MDINEXT: Debug.Print "WM_MDINEXT"
Case WM_MDIREFRESHMENU: Debug.Print "WM_MDIREFRESHMENU"
Case WM_MDIRESTORE: Debug.Print "WM_MDIRESTORE"
Case WM_MDISETMENU: Debug.Print "WM_MDISETMENU"
Case WM_MDITILE: Debug.Print "WM_MDITILE"
Case WM_MEASUREITEM: Debug.Print "WM_MEASUREITEM"
Case WM_MENUCHAR: Debug.Print "WM_MENUCHAR"
Case WM_MENUSELECT: Debug.Print "WM_MENUSELECT"
Case WM_MOUSEACTIVATE: Debug.Print "WM_MOUSEACTIVATE"
Case WM_MOUSEFIRST: ' Debug.Print "WM_MOUSEFIRST"
Case WM_MOUSELAST: Debug.Print "WM_MOUSELAST"
Case WM_MOUSEMOVE: ' DeBug.Print "WM_MOUSEMOVE"
Case WM_MOVE: Debug.Print "WM_MOVE"
Case WM_NCACTIVATE: Debug.Print "WM_NCACTIVATE"
Case WM_NCCALCSIZE: Debug.Print "WM_NCCALCSIZE"
Case WM_NCCREATE: Debug.Print "WM_NCCREATE"
Case WM_NCDESTROY: Debug.Print "WM_NCDESTROY"
Case WM_NCHITTEST: 'Debug.Print "WM_NCHITTEST"
Case WM_NCLBUTTONDBLCLK: Debug.Print "WM_NCLBUTTONDBLCLK"
Case WM_NCLBUTTONDOWN: Debug.Print "WM_NCLBUTTONDOWN"
Case WM_NCLBUTTONUP: Debug.Print "WM_NCLBUTTONUP"
Case WM_NCMBUTTONDBLCLK: Debug.Print "WM_NCMBUTTONDBLCLK"
Case WM_NCMBUTTONDOWN: Debug.Print "WM_NCMBUTTONDOWN"
Case WM_NCMBUTTONUP: Debug.Print "WM_NCMBUTTONUP"
Case WM_NCMOUSEMOVE: ' DeBug.Print "WM_NCMOUSEMOVE"
Case WM_NCPAINT: Debug.Print "WM_NCPAINT"
Case WM_NCRBUTTONDBLCLK: Debug.Print "WM_NCRBUTTONDBLCLK"
Case WM_NCRBUTTONDOWN: Debug.Print "WM_NCRBUTTONDOWN"
Case WM_NCRBUTTONUP: Debug.Print "WM_NCRBUTTONUP"
Case WM_NEXTDLGCTL: Debug.Print "WM_NEXTDLGCTL"
Case WM_PAINT: Debug.Print "WM_PAINT"
Case WM_PAINTCLIPBOARD: Debug.Print "WM_PAINTCLIPBOARD"
Case WM_PAINTICON: Debug.Print "WM_PAINTICON"
Case WM_PALETTECHANGED: Debug.Print "WM_PALETTECHANGED"
Case WM_PALETTEISCHANGING: Debug.Print "WM_PALETTEISCHANGING"
Case WM_PARENTNOTIFY: Debug.Print "WM_PARENTNOTIFY"
Case WM_PASTE: Debug.Print "WM_PASTE"
Case WM_PENWINFIRST: Debug.Print "WM_PENWINFIRST"
Case WM_PENWINLAST: Debug.Print "WM_PENWINLAST"
Case WM_POWER: Debug.Print "WM_POWER"
Case WM_PSD_ENVSTAMPRECT: Debug.Print "WM_PSD_ENVSTAMPRECT"
Case WM_PSD_FULLPAGERECT: Debug.Print "WM_PSD_FULLPAGERECT"
Case WM_PSD_GREEKTEXTRECT: Debug.Print "WM_PSD_GREEKTEXTRECT"
Case WM_PSD_MARGINRECT: Debug.Print "WM_PSD_MARGINRECT"
Case WM_PSD_MINMARGINRECT: Debug.Print "WM_PSD_MINMARGINRECT"
Case WM_PSD_PAGESETUPDLG: Debug.Print "WM_PSD_PAGESETUPDLG"
Case WM_PSD_YAFULLPAGERECT: Debug.Print "WM_PSD_YAFULLPAGERECT"
Case WM_QUERYDRAGICON: Debug.Print "WM_QUERYDRAGICON"
Case WM_QUERYENDSESSION: Debug.Print "WM_QUERYENDSESSION"
Case WM_QUERYNEWPALETTE: Debug.Print "WM_QUERYNEWPALETTE"
Case WM_QUERYOPEN: Debug.Print "WM_QUERYOPEN"
Case WM_QUEUESYNC: Debug.Print "WM_QUEUESYNC"
Case WM_QUIT: Debug.Print "WM_QUIT"
Case WM_RBUTTONDBLCLK: Debug.Print "WM_RBUTTONDBLCLK"
Case WM_RBUTTONDOWN: Debug.Print "WM_RBUTTONDOWN"
Case WM_RBUTTONUP: Debug.Print "WM_RBUTTONUP"
Case WM_RENDERALLFORMATS: Debug.Print "WM_RENDERALLFORMATS"
Case WM_RENDERFORMAT: Debug.Print "WM_RENDERFORMAT"
Case WM_SETCURSOR: 'Debug.Print "WM_SETCURSOR"
Case WM_SETFOCUS: Debug.Print "WM_SETFOCUS"
Case WM_SETFONT: Debug.Print "WM_SETFONT"
Case WM_SETHOTKEY: Debug.Print "WM_SETHOTKEY"
Case WM_SETREDRAW: Debug.Print "WM_SETREDRAW"
Case WM_SETTEXT: Debug.Print "WM_SETTEXT"
Case WM_SHOWWINDOW: Debug.Print "WM_SHOWWINDOW"
Case WM_SIZE: Debug.Print "WM_SIZE"
Case WM_SIZECLIPBOARD: Debug.Print "WM_SIZECLIPBOARD"
Case WM_SPOOLERSTATUS: Debug.Print "WM_SPOOLERSTATUS"
Case WM_SYSCHAR: Debug.Print "WM_SYSCHAR"
Case WM_SYSCOLORCHANGE: Debug.Print "WM_SYSCOLORCHANGE"
Case WM_SYSCOMMAND: Debug.Print "WM_SYSCOMMAND"
Case WM_SYSDEADCHAR: Debug.Print "WM_SYSDEADCHAR"
Case WM_SYSKEYDOWN: Debug.Print "WM_SYSKEYDOWN"
Case WM_SYSKEYUP: Debug.Print "WM_SYSKEYUP"
Case WM_SYSTIMER: Debug.Print "WM_SYSTIMER"
Case WM_TIMECHANGE: Debug.Print "WM_TIMECHANGE"
Case WM_TIMER: Debug.Print "WM_TIMER"
Case WM_UNDO: Debug.Print "WM_UNDO"
Case WM_USER: Debug.Print "WM_USER"
Case WM_VKEYTOITEM: Debug.Print "WM_VKEYTOITEM"
Case WM_VSCROLL: Debug.Print "WM_VSCROLL"
Case WM_VSCROLLCLIPBOARD: Debug.Print "WM_VSCROLLCLIPBOARD"
Case WM_WINDOWPOSCHANGED: Debug.Print "WM_WINDOWPOSCHANGED"
Case WM_WINDOWPOSCHANGING: Debug.Print "WM_WINDOWPOSCHANGING"
Case WM_WININICHANGE: Debug.Print "WM_WININICHANGE"
Case Else
Debug.Print uMsg
End Select
End Sub
