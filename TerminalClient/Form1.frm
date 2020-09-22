VERSION 5.00
Object = "{FC7C887E-70BD-4ADB-8BED-8681D74F36D1}#1.0#0"; "msrdp.ocx"
Begin VB.Form frmMain 
   Caption         =   "Terminal Server Client"
   ClientHeight    =   2085
   ClientLeft      =   7005
   ClientTop       =   5835
   ClientWidth     =   3765
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   Begin VB.CheckBox chkUserPass 
      Caption         =   "Use Password"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cboRes 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "Form1.frx":0442
      Left            =   840
      List            =   "Form1.frx":0458
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.CheckBox chkBitmap 
      Caption         =   "Enable Cache Bitmaps"
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CheckBox chkCompress 
      Caption         =   "Enable Data Compression"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
   End
   Begin MSTSCLibCtl.MsTscAx msts 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
      Server          =   ""
      Domain          =   ""
      UserName        =   ""
      FullScreen      =   ""
      StartConnected  =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Resolution"
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "UserName"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Using the Microsoft Terminal Server Control
'Coded by Chris Peneguy
'Updated Date Dec. 10 2001
'http://www.secureinsights.com
'chris@secureinsights.com
'Use this code any which way you like
'####################################################################################
' The auto Password code only works when you set the configuration on the Server-side.
' Log on to the Terminal Server as an administrator
' Start\Programs\Administrative Tools\Terminal Services Configuration
' Click on Connections
' On the Right Pane, right-click on RDP-Tcp and choose Properties
' Click on the "Logon Settings" Tab
' Uncheck "Always prompt for password" and click OK
'####################################################################################
'Update
' Removed the full screen even at smaller resolutions
' Added Bitmap Caching to speed up image display
' Added compression to speed up data tranfers
' Added Automatic Password entry
'
Option Explicit
Dim Server As String    'Server Address
Dim UserName As String  'User Login Name
Dim Password As String  'User Password
Dim resWidth As String  'Resolution Size - Width
Dim resHeight As String 'Resolution Size - Height
Dim Reso As String
Const FullScreenWarnTxt1 = "Your current security settings do not allow automatically switching to fullscreen mode."
Const FullScreenWarnTxt2 = "You can use ctrl-alt-pause to toggle your terminal services session to fullscreen mode"
Const FullScreenTitleTxt = "Terminal Services Connection "
Const ErrMsgText = "Error connecting to terminal server: "
Private Obj As IMsTscNonScriptable

Private Sub Form_Load()

msts.Visible = False
txtPassword.PasswordChar = "*"
End Sub

Private Sub cmdGo_Click()


Server = txtServer.Text       'Server Address
UserName = txtUserName.Text   'User Login Name
Password = txtPassword.Text   'User Password
Reso = cboRes.Text            'Resolution Temp


If Server = "" Then
  MsgBox "Please enter a Server address"
  txtServer.SetFocus
 ElseIf UserName = "" Then
    MsgBox "Please enter a UserName"
    txtUserName.SetFocus
  ElseIf chkUserPass = "1" And txtPassword.Text = "" Then
    MsgBox "Please enter a Password"
   ElseIf chkUserPass = "0" And txtPassword.Text <> "" Then
     MsgBox "Please check the Use Password Check Box"
    ElseIf Reso = "" Then
      MsgBox "Please choose a Resolution"
      cboRes.SetFocus
   Else
    Res
    Connect
End If

End Sub

Sub Res()
 'Sets the Resolution size for the terminal based
 'on the resolution chosen in the combo box
  Dim a() As String
  If Reso = "Full-Screen " Then
    resWidth = Screen.Width \ Screen.TwipsPerPixelX    'Converts from Twips to Pixels
    resHeight = Screen.Height \ Screen.TwipsPerPixelY  'Convers from Twis to Pixels
    msts.SecuredSettings.FullScreen = 1
  Else
    a = Split(Reso, " x ")
    resWidth = a(0)
    resHeight = a(1)
    msts.SecuredSettings.FullScreen = 0
  End If
  
End Sub


Sub Connect()
'Connecting to the Terminal Server
  Dim intCnt As Integer
  Set Obj = msts.Object
  frmMain.Top = 15
  frmMain.Left = 15
  frmMain.Width = resWidth * Screen.TwipsPerPixelX
  frmMain.Height = resHeight * Screen.TwipsPerPixelY
  msts.Width = resWidth '
  msts.Height = resHeight '
  
  If chkBitmap = 1 Then
    msts.AdvancedSettings.BitmapPeristence = 1
  Else
    msts.AdvancedSettings.BitmapPeristence = 0
  End If
  If chkCompress.Value = 1 Then
    msts.AdvancedSettings.Compress = 1
  Else
    msts.AdvancedSettings.Compress = 0
  End If
    
  For intCnt = 0 To frmMain.Controls.Count - 1
    If frmMain.Controls(intCnt).Name <> "tmr" Then
      frmMain.Controls(intCnt).Visible = False
    End If
  Next intCnt
  
  msts.Visible = True
  If txtPassword.Text <> "" Then
  Obj.ClearTextPassword = Password
  End If
  msts.Server = Server
  msts.UserName = UserName
  msts.DesktopHeight = resHeight
  msts.DesktopWidth = resWidth
  
  
  msts.Connect
 
  
End Sub


Private Sub msts_OnDisconnected(ByVal DisconnectReason As Long)
'Resets the form back to original state
  Dim intCnt As Integer
  frmMain.Height = 2595
  frmMain.Width = 3885
  For intCnt = 0 To frmMain.Controls.Count - 1
   If frmMain.Controls(intCnt).Name <> "tmr" Then
      frmMain.Controls(intCnt).Visible = True
   End If
  Next intCnt
  msts.Visible = False

End Sub





