Attribute VB_Name = "Module1"
Option Explicit

' General Declares:
Public TaskbarHwnd As Long ' Containts Windows Taskbar visibility status.
Public KBDhwnd As Long ' Disable Windows key codes (Alt+TAB, Ctrl+ESC)

' Declarations to enable frmMain to cover entire screen, and for hiding/displaying the Windows Taskbar, and frmDebug.
Public Declare Function GetSystemMetrics Lib "user32" (ByVal vIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' Supporting Constants. (interspersed)
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80

' Stuff to keep CTRL+TAB, ALT+TAB, CTRL+ESC, and the Windows System keys from being used.
Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const HC_ACTION = 0
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_DELETE = &H2E
Public Const WH_KEYBOARD_LL = 13
Public Const LLKHF_ALTDOWN = &H20
Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    FLAGS As Long
    time As Long
    dwExtraInfo As Long
End Type
Dim KBD As KBDLLHOOKSTRUCT


' Put this in a module:
Public Function CatchKeys(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' This function captures the Windows system commands:
' ALT+TAB
' CTRL+TAB
' CTRL+ESC
' Windows Flag Key (between CTRL and ALT)
'
' Note that Windows 2000 does *NOT ALLOW* CTRL+ALT+DEL to be intercepted. Period. This is by
' design by Microsoft for security purposes, so any machine that this test will be administered MUST
' incorporate a Group Policy (or User Policy) to disable all of the features available
' on the dialog box that comes up when you hit CTRL+ALT+DEL.
' -----------------------------------------------------------------------------------------
Dim bEatIt As Boolean: bEatIt = False

    If (nCode = HC_ACTION) Then
        If wParam = WM_KEYDOWN Or wParam = WM_SYSKEYDOWN Or wParam = WM_KEYUP Or wParam = WM_SYSKEYUP Then
            CopyMemory KBD, ByVal lParam, Len(KBD)
            bEatIt = ((KBD.vkCode = VK_ESCAPE) And ((GetKeyState(VK_CONTROL) And &H8000) <> 0)) Or _
                     ((KBD.vkCode = VK_TAB) And ((KBD.FLAGS And LLKHF_ALTDOWN) <> 0)) Or _
                     ((KBD.vkCode = VK_ESCAPE) And ((KBD.FLAGS And LLKHF_ALTDOWN) <> 0)) Or KBD.vkCode = 91
        End If
    End If
     
    If bEatIt Then
        CatchKeys = -1
    Else
        CatchKeys = CallNextHookEx(0, nCode, wParam, ByVal lParam)
    End If

End Function

' You can use this to set the form the whole screen:
Public Function CoverWholeScreen(frm As Object) As Boolean
' Sets the form to the size of the entire screen. 800x600 is
' the optimal size for this application.
Dim iXPos As Long: iXPos = 0
Dim iYPos As Long: iYPos = 0
Dim iRet As Long: iRet = 0

On Error GoTo ErrorHandler

    frm.WindowState = vbNormal

    iXPos = GetSystemMetrics(SM_CXSCREEN)
    iYPos = GetSystemMetrics(SM_CYSCREEN)

    iRet = SetWindowPos(frm.hwnd, HWND_TOP, 0, 0, iXPos, iYPos, SWP_SHOWWINDOW)

    CoverWholeScreen = iRet <> 0

ErrorHandler:
End Function


