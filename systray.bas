Attribute VB_Name = "Module2"
Option Explicit
'shellapi.h


'Declared Functions

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
 Public Const WM_LBUTTONUP = &H202
 Public Const WM_RBUTTONUP = &H205

    'Types
    
    Public NOTIFYICONDATA As NOTIFYICONDATA
    Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
    End Type
    

    'Constants

    Public Const NIF_ICON = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_TIP = &H4
    Public Const NIF_STATE = &H8
    Public Const NIF_INFO = &H10

    Public Const NIM_ADD = &H0
    Public Const NIM_DELETE = &H2
    Public Const NIM_MODIFY = &H1
    Public Const NIM_SETFOCUS = &H3
    Public Const NIM_SETVERSION = &H4

