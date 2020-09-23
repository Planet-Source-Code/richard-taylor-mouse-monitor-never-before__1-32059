VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mouse Monitor"
   ClientHeight    =   1455
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   4545
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2.566
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   8.017
   Begin VB.CheckBox Hiding 
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picHook 
      Height          =   375
      Left            =   2400
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Millimetres"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Centimetres"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Pixels"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Created By Richard Taylor"
      Height          =   615
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Y 
      Caption         =   "0"
      Height          =   1215
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label X 
      Caption         =   "0"
      Height          =   1095
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "0 pixels"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''IF YOU LIKE WHAT IVE DONE IN THIS PROGRAM PLEASE PLEASE PLEASE PLEASE PLEASE VOTE ON THE WEB SITE
'''I WILL PROBABLY MAKE ANOTHER WITH MORE UNITS SOMETIME WITH A LOG AND THINGS, BUT FOR NOW I THINK
'''THAT THIS WILL SATISFY YOURE PROGRAMMING NEEDS :)
'''(PLEASE TELL ME IF THIS HAS BEEN DONE BEFORE BECAUSE IVE NOT SEEN IT)
'''
'''ENJOY
'''RICHARD TAYLOR
'''UK
'////////START THE CODE\\\\\\\\'
'JUST SOME LITTLE API CALLS AND THINGS TO GET IT WORKING
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim P As POINTAPI, C As Integer, OldX As Single, OldY As Single, Chan As Integer, SysTray As NOTIFYICONDATA
Private Sub Form_Activate()
'JUST MAKE SURE THAT THERE ARNT 2 PROGRAMS RUNNING AT THE SAME TIME SO WE DONT GET AN ERROR!
If App.PrevInstance = True Then
MsgBox "ERROR: You can only run the program once at a time!", vbCritical, "Error"
End
End If
'MOVE THE FORM TO THE TOP LEFT OF THE SCREEN
Me.Move 0, 0
'SET THE SYSTEM TRAY ICON SETTINGS
With SysTray
.cbSize = Len(SysTray)
.hIcon = Me.Icon
.hwnd = picHook.hwnd
.szTip = "Mouse Monitor                        Created By Richard Taylor"
.ucallbackMessage = WM_MOUSEMOVE
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uId = 1&
End With
'PUT THE ICON IN THE SYSTEM TRAY
Shell_NotifyIcon NIM_ADD, SysTray
'MAKE THE COMPUTER KNOW THAT WERE NOT QUITING YET WERE ONLY GOING TO HIDE IT
Hiding.Value = 1
End Sub

Private Sub Form_Load()
'STOP THE TIMER SO WE DONT GET AN ERROR OR STRANGE NUMBERS
Timer1.Enabled = False
'GET THE CURRENT DISTANCE TAKEN
Label1.Caption = GetSetting("Mouse Monitor", "Milage", "Distance Up To Now")
If Label1.Caption = "" Then
'IF THE CURRENT DISTANCE IS "" THEN CREATE A NEW ONE
SaveSetting "Mouse Monitor", "Milage", "Distance Up To Now", "0 pixels"
Label1.Caption = GetSetting("Mouse Monitor", "Milage", "Distance Up To Now")
End If
'IF THE CURRENT DISTANCE IS IN PIXELS THEN SELECT OPTION 1, AND SO ON FOR THE OTHERS :)
If InStr(1, Label1.Caption, "pixels") Then
Option1.Value = True
ElseIf InStr(1, Label1.Caption, "cm") Then
Option2.Value = True
ElseIf InStr(1, Label1.Caption, "mm") Then
Option3.Value = True
End If
'MAKE SURE THAT ITS STILL NOT QUITTING YET :), JUST INCASE
Hiding.Value = 1
'START THE TIMER AND START MONITORING
Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Hiding.Value = 0 Then
'IF WE ARNT HIDING THEN WE MUST CLOSE THE PROGRAM (ONLY THE EXIT MENU CAN DO THIS)
SysTray.cbSize = Len(SysTray)
SysTray.hwnd = picHook.hwnd
SysTray.uId = 1&
Shell_NotifyIcon NIM_DELETE, SysTray
'SAVE THE CURRENT MILAGE
SaveSetting "Mouse Monitor", "Milage", "Distance Up To Now", Label1.Caption
End
Else
'STOP QUITTING
Cancel = 1
'ONLY HIDE =D
Me.Hide
End If
End Sub

Private Sub mnuAbout_Click()
frmabout.Show
End Sub

Private Sub mnuExit_Click()
'MAKE IT KNOW WE ARNT HIDING NOW, WE ARE QUITTING
Hiding.Value = 0
Unload Me
End Sub
Private Sub mnuReset_Click()
'RESET THE CAPTION TO 0, AND CHECK THE UNITS AFTERWARDS
If Option1.Value = True Then
Label1.Caption = "0 pixels"
ElseIf Option2.Value = True Then
Label1.Caption = "0 cm"
ElseIf Option3.Value = True Then
Label1.Caption = "0 mm"
End If
End Sub
Private Sub mnuShow_Click()
Me.Show
End Sub
Private Sub Option1_Click()
'IF THE CAPTION HAS CM IN IT THEN TO CHANGE IT TO PIXELS WE DO THIS :
If InStr(1, Label1.Caption, "cm") Then
Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2)
Label1.Caption = Val(Label1.Caption) * 37.7952389628211
Label1.Caption = Int(Label1.Caption)
Label1.Caption = Label1.Caption & " pixels"
'IF THE CAPTION HAS MM IN IT THEN TO CHANGE IT TO PIXELS WE DO THIS :
ElseIf InStr(1, Label1.Caption, "mm") Then
Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2)
Label1.Caption = Val(Label1.Caption) * 3.77952389628211
Label1.Caption = Int(Label1.Caption)
Label1.Caption = Label1.Caption & " pixels"
End If
End Sub
Private Sub Option2_Click()
'THE SAME AS THE ABOVE BUT WITH DIFFERENT VALUES :)
If InStr(1, Label1.Caption, "pixels") Then
Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 7)
Label1.Caption = Val(Label1.Caption) / 37.7952389628211
Label1.Caption = Int(Label1.Caption)
Label1.Caption = Label1.Caption & " cm"
ElseIf InStr(1, Label1.Caption, "mm") Then
Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2)
Label1.Caption = Val(Label1.Caption) / 10
Label1.Caption = Int(Label1.Caption)
Label1.Caption = Label1.Caption & " cm"
End If
End Sub
Private Sub Option3_Click()
'THE SAME AS THE ABOVE 2 BUT USING DIFFERENT UNITS
If InStr(1, Label1.Caption, "pixels") Then
Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 7)
Label1.Caption = Val(Label1.Caption) / 3.77952389628211
Label1.Caption = Int(Label1.Caption)
Label1.Caption = Label1.Caption & " mm"
ElseIf InStr(1, Label1.Caption, "cm") Then
Label1.Caption = Left(Label1.Caption, Len(Label1.Caption) - 2)
Label1.Caption = Val(Label1.Caption) * 10
Label1.Caption = Int(Label1.Caption)
Label1.Caption = Label1.Caption & " mm"
End If
End Sub

Private Sub picHook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'THE TIME THAT THE MOUSE HAS CLICKED THE ICON IN THE SYSTEM TRAY
Static rec As Boolean, msg As Long
msg = X / Screen.TwipsPerPixelX
If rec = False Then
rec = True
Select Case msg
'IF THE MOUSE IS THE RIGHT MOUSE BUTTON AND THE BUTTON HAS LET GO, THEN BRING UP THE MENU WHERE THE MOUSE IS
Case WM_RBUTTONUP:
Me.PopupMenu mnuFile
'IF THE MOUSE HAS DOUBLE CLICKED THE ICON THEN LOAD UP THE FORM TO SAVE TIME WITH THE MENU :)
Case WM_LBUTTONDBLCLK:
Me.Show
End Select
rec = False
End If
End Sub

Private Sub Timer1_Timer()
'SET "P" SO THE POINTAPI MOUSE POSITION
GetCursorPos P
'IF THE MOUSE HAS CHANGED THEN FLAG C IS 1 SO IT CAN CHANGE THE MILAGE
If Str(P.X) <> X.Caption Then
OldX = X
X.Caption = P.X
C = 1
End If
If Str(P.Y) <> Y.Caption Then
OldY = Y
Y.Caption = P.Y
C = 2
End If
If C = 1 Or 2 Then
'IF PIXELS IS CHOSEN THEN, AND SO ON etc. - I CANT COMMENT THE REST OF THIS SUB BECAUSE ITS TO HARD TO EXPLAIN YET SIMPLE TO UNDERSTAND.. IF U NO WHT I MEAN
If Option1.Value = True Then
If P.X > OldX Then
Chan = P.X - OldX
Label1.Caption = Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 7) + Chan)
Label1.Caption = Label1.Caption & " pixels"
End If
If P.X < OldX Then
Chan = OldX - P.X
Label1.Caption = Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 7) + Chan)
Label1.Caption = Label1.Caption & " pixels"
End If
If P.Y > OldY Then
Chan = P.Y - OldY
Label1.Caption = Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 7) + Chan)
Label1.Caption = Label1.Caption & " pixels"
End If
If P.Y < OldY Then
Chan = OldY - P.Y
Label1.Caption = Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 7) + Chan)
Label1.Caption = Label1.Caption & " pixels"
End If
ElseIf Option2.Value = True Then
If P.X > OldX Then
Chan = (P.X - OldX) / 37.7952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " cm"
End If
If P.X < OldX Then
Chan = (OldX - P.X) / 37.7952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " cm"
End If
If P.Y > OldY Then
Chan = (P.Y - OldY) / 37.7952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " cm"
End If
If P.Y < OldY Then
Chan = (OldY - P.Y) / 37.7952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " cm"
End If
ElseIf Option3.Value = True Then
If C <> 2 Then Exit Sub
If P.X > OldX Then
Chan = (P.X - OldX) / 3.77952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " mm"
End If
If P.X < OldX Then
Chan = (OldX - P.X) / 3.77952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " mm"
End If
If P.Y > OldY Then
Chan = (P.Y - OldY) / 3.77952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " mm"
End If
If P.Y < OldY Then
Chan = (OldY - P.Y) / 3.77952389628211
Label1.Caption = Int(Val(Mid(Label1.Caption, 1, Len(Label1.Caption) - 2) + Chan))
Label1.Caption = Label1.Caption & " mm"
End If
End If
End If
End Sub
