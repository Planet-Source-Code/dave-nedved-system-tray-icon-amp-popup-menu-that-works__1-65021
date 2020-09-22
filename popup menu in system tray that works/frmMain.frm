VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Popup Menu That Works!"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   MaxButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":0000
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Menu mPopupSys 
      Caption         =   "&SysTray"
      Visible         =   0   'False
      Begin VB.Menu mPopAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mPopHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mSysBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mPopRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mSysBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mPopExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'the form must be fully visible before calling Shell_NotifyIcon
Me.Show
Me.Refresh
 With nid
  .cbSize = Len(nid)
  .hwnd = Me.hwnd
  .uId = vbNull
  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  .uCallBackMessage = WM_MOUSEMOVE
  .hIcon = Me.Icon
  .szTip = "This is a Sample Tool Tip" & vbNullChar
  End With
Shell_NotifyIcon NIM_ADD, nid
App.TaskVisible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
 Dim Result As Long
 Dim msg As Long
'the value of X will vary depending upon the scalemode setting
  If Me.ScaleMode = vbPixels Then
   msg = X
  Else
   msg = X / Screen.TwipsPerPixelX
  End If
  Select Case msg
   Case WM_LBUTTONUP        '514 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
   Case WM_LBUTTONDBLCLK    '515 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
   Case WM_RBUTTONUP        '517 display popup menu
    Result = SetForegroundWindow(Me.hwnd)
    Me.PopupMenu Me.mPopupSys
  End Select
End Sub

Private Sub Form_Resize()
 'this is necessary to assure that the minimized window is hidden
 If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'this removes the icon from the system tray
  Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mPopAbout_Click()
 MsgBox "A simple system tray icon, that can show a popup menu," & vbNewLine & "and will hide the menu when clicking on the desktop" & vbNewLine & vbNewLine & "By Dave Nedved, API's from MSDN"
End Sub

Private Sub mPopExit_Click()
'called when user clicks the popup menu Exit command
 Unload Me
End Sub

Private Sub mPopHelp_Click()
 MsgBox "Sample Menu"
End Sub

Private Sub mPopRestore_Click()
 'called when the user clicks the popup menu Restore command
  Dim Result As Long
  Me.WindowState = vbNormal
  Result = SetForegroundWindow(Me.hwnd)
  Me.Show
End Sub
