VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Lock"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4440
   Icon            =   "protect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Hide"
      Height          =   390
      Left            =   2760
      TabIndex        =   11
      Top             =   6120
      Width           =   1440
   End
   Begin VB.Timer Timer3 
      Left            =   3510
      Top             =   3525
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   390
      Left            =   1380
      TabIndex        =   10
      Top             =   6120
      Width           =   1230
   End
   Begin MSComDlg.CommonDialog com1 
      Left            =   3465
      Top             =   6285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   390
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   8
      Top             =   4335
      Width           =   4185
   End
   Begin VB.ListBox file1 
      Height          =   1620
      Left            =   105
      TabIndex        =   4
      Top             =   375
      Width           =   4200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3390
      TabIndex        =   3
      Top             =   2835
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   3375
      TabIndex        =   2
      Top             =   2430
      Width           =   975
   End
   Begin VB.ListBox protect 
      Height          =   1425
      ItemData        =   "protect.frx":1272
      Left            =   105
      List            =   "protect.frx":1274
      TabIndex        =   1
      Top             =   2415
      Width           =   3210
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3480
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4200
      Top             =   2535
   End
   Begin VB.ListBox lstTasks 
      Height          =   255
      Left            =   315
      TabIndex        =   0
      Top             =   870
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4440
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Protected files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   3975
      Width           =   4155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Protected Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   105
      TabIndex        =   6
      Top             =   2055
      Width           =   4200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Window List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   45
      Width           =   4185
   End
   Begin VB.Menu mnuop 
      Caption         =   "&Folder Options"
      Begin VB.Menu mnulock 
         Caption         =   "&Lock All"
      End
      Begin VB.Menu mnuunlock 
         Caption         =   "&Unlock All"
      End
      Begin VB.Menu mnulist 
         Caption         =   "&Add from List"
      End
      Begin VB.Menu mnuhide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnulockop 
      Caption         =   "L&ock "
      Begin VB.Menu mnucon 
         Caption         =   "Control Panel"
      End
      Begin VB.Menu mnubin 
         Caption         =   "Recycle Bin"
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "Protected files"
      Begin VB.Menu mnulfile 
         Caption         =   "Lock All"
      End
      Begin VB.Menu mnuunlfile 
         Caption         =   "Unlock All"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tmp As String
Public apiError As Long
Dim x, X1 As Integer
Dim filenumber As Integer


Private Sub Command1_Click()
Dim tmpPath As String
Dim file1 As String
tmpPath = BrowseForFolder(tmpPath)
If tmpPath = "" Then
    Exit Sub
Else
file1 = GetFileTitle(tmpPath)
protect.AddItem LCase(file1)
End If
ListSave protect, "c:\folder.txt"
End Sub

Private Sub Command2_Click()
protect.RemoveItem protect.ListIndex
ListSave protect, "c:\folder.txt"
End Sub

Private Sub Command3_Click()
com1.Filter = "All files | *.*"
com1.FilterIndex = 2
com1.ShowOpen
If com1.FileName <> "" Then
List1.AddItem com1.FileName
Else
End If
Module1.ListSave List1, "c:\file.txt"
mnuunlfile_Click
mnulfile_Click
End Sub

Private Sub pGetTasks()

    Call fEnumWindows(lstTasks)
    On Error Resume Next
    lstTasks.ListIndex = -1
    
End Sub

Private Sub Command4_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
Module1.ListSave List1, "c:\file.txt"
mnuunlfile_Click
mnulfile_Click
End Sub

Private Sub Command5_Click()
Me.Hide
End Sub

Private Sub Form_Load()
App.TaskVisible = False
Module1.ListOpen List1, "c:\file.txt"
Module1.ListOpen protect, "c:\folder.txt"
Me.Hide


With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = Me.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = 512
        .hIcon = Me.Icon
        .szTip = "Protect" & Chr(0) 'Tooltip text
    End With
    apiError = Shell_NotifyIcon(NIM_ADD, NOTIFYICONDATA)
mnulfile_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Subclass callback
    Dim Password As String
    Dim tmpLong As Single
    tmpLong = x / Screen.TwipsPerPixelX
    
    Select Case tmpLong
        Case WM_LBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd)
            
            Password = InputBox("Password")
            If Password = "password" Then
            Me.show
            Else
            MsgBox "Sorry Wrong password"
            End If
        Case WM_RBUTTONUP
            apiError = SetForegroundWindow(Me.hwnd)
            Password = InputBox("Password", "Password")
            If Password = "finn" Then
            Me.show
            Else
            MsgBox "Sorry Wrong password"
            End If
        
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
With NOTIFYICONDATA
        .cbSize = Len(NOTIFYICONDATA)
        .hwnd = Me.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = vbNull
        .hIcon = Me.Icon
        .szTip = Chr(0)
    End With
    apiError = Shell_NotifyIcon(NIM_DELETE, NOTIFYICONDATA)
End Sub


Private Sub mnubin_Click()
protect.AddItem "recycle bin"
ListSave protect, "c:\folder.txt"
End Sub

Private Sub mnuclose_Click()
Unload Me
End Sub

Private Sub mnucon_Click()
protect.AddItem "control panel"
ListSave protect, "c:\folder.txt"
End Sub



Private Sub mnuhide_Click()
Me.Hide
End Sub

Public Sub mnulfile_Click()
Dim fileLock As String

    Open "C:\file.txt" For Input As #1


    Do While Not EOF(1)
        Line Input #1, fileLock
        filenumber = FreeFile
        Open fileLock For Binary Shared As #filenumber
        Lock #filenumber
    Loop
    Close #1
End Sub

Private Sub mnulist_Click()
Dim num As Integer
num = InputBox("Please pick from Window list", "Window Picker")
num = Val(num) - 1
protect.AddItem file1.List(num)
End Sub

Public Sub mnulock_Click()
Dim lock1 As Integer
For lock1 = 0 To protect.ListCount - 1
protect.ItemData(lock1) = 0
Next lock1

End Sub

Public Sub mnuunlfile_Click()

Open "C:\file.txt" For Input As #1
    For x = 1 To FreeFile - 1
    Close #x
    Next x
    Close #1
    
End Sub

Public Sub mnuunlock_Click()
Dim unlock1 As Integer
For unlock1 = 0 To protect.ListCount - 1
protect.ItemData(unlock1) = 1
Next unlock1

End Sub

Private Sub protect_Click()
If protect.ItemData(protect.ListIndex) = 0 Then
Me.Caption = protect.List(protect.ListIndex) & " = Locked"
Else
Me.Caption = protect.List(protect.ListIndex) & " = Unlocked"
End If
End Sub

Private Sub Timer1_Timer()
Call pGetTasks
showfile
End Sub

Private Sub Timer2_Timer()
Dim winname, pro As String
winname = file1.List(0)
If protect.ListCount = 0 Then
Exit Sub
Else
For X1 = 0 To protect.ListCount - 1
pro = protect.List(X1)
If winname = pro Then
access
Exit Sub
ElseIf winname = "exploring - " + pro Then
access
Exit Sub
End If
Next X1
End If
End Sub

Private Sub access()
Dim x As Integer
Dim Y As Integer
Dim pass As String
Dim z As Integer
Dim hid As Long
hid = lstTasks.ItemData(0)
Y = protect.ItemData(X1)
If Y = 0 Then
HideWin hid
pass = InputBox("Please enter a Password", "Password")
If pass = "password" Then
protect.ItemData(X1) = 1
ShowWin hid
Else
MsgBox "Invalid Password The file or folder you tried to open will now close", , "Password"
CloseWin hid
protect.ItemData(X1) = 0
End If
End If
End Sub

Public Function BrowseForFolder(selectedPath As String) As String
Dim Browse_for_folder As BROWSEINFOTYPE
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
With Browse_for_folder
    .hOwner = Me.hwnd
    .lpszTitle = "Browse for folders with directory pre-selection"
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr)
    selectedPathPointer = LocalAlloc(LPTR, Len(selectedPath) + 1)
    CopyMemory ByVal selectedPathPointer, ByVal selectedPath, Len(selectedPath) + 1
    .lParam = selectedPathPointer
End With
itemID = SHBrowseForFolder(Browse_for_folder)
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1)
    End If
    Call CoTaskMemFree(itemID)
End If
Call LocalFree(selectedPathPointer)
End Function


Public Sub showfile()
Dim show As Integer
file1.Clear
For show = 0 To lstTasks.ListCount - 1
file1.AddItem LCase(lstTasks.List(show))
Next show
End Sub
