VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "******** VIRIIO Port Listener ********"
   ClientHeight    =   3435
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4440
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   2880
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Timer OFF"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Timer ON"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox secs 
      Height          =   285
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "60"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Scan All Possible Ports!"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox logport 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":0442
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "65530"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Ports Every :"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":0466
      Top             =   0
      Width           =   4425
   End
   Begin VB.Image Image1 
      Height          =   1785
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":7226
      Top             =   1680
      Width           =   4425
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim portnum As Long
Dim start As String
Dim TimerOnOff As String


Private Sub Command1_Click()
Command2.Enabled = True
If Text1.Text = "" Then
MsgBox "You must enter a number int the 'FROM' text box!"
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "You must enter a number in the 'TO' text box!"
Exit Sub
End If

Text1.Locked = True
Text2.Locked = True
Command1.Enabled = False
Winsock1.Close
start = True
logport.Text = " ***** VIRIIO Port Listener *****"

Call scanningports
logport.Text = logport.Text & vbCrLf & "Ports " & Text1.Text & "- " & Text3.Text & "  have been scanned Successfully."
End Sub

Sub scanningports()
Dim porttwo As Long
portnum = Text1.Text
porttwo = Text2.Text
Command2.Enabled = True
On Error GoTo viriio
Do
portnum = portnum + 1
DoEvents
If start = True Then
Winsock1.Close 'Close current winsock
DoEvents 'Slows program down but no errors
Winsock1.LocalPort = portnum
DoEvents
Text3.Text = portnum
Winsock1.Listen 'listen on port
DoEvents
Else 'If someone clicks STOP!
portnum = 0
Command1.Enabled = True
Text1.Locked = False
Text2.Locked = False
Exit Sub
End If
Winsock1.Close
DoEvents
   Loop Until portnum >= porttwo
   'if it finishes with no errors do this stuff
portnum = 0
Command1.Enabled = True
logport.Text = logport.Text & vbCrLf & "Scanning Ports Done!" & vbCrLf
Text1.Locked = False
Text2.Locked = False
viriio:

If Err.Number = 10048 Then
logport.Text = logport.Text & vbCrLf & "Port " & Winsock1.LocalPort & " in Use!"
Resume Next ' go back to where the error happened
End If

End Sub

Private Sub Command2_Click()
Command2.Enabled = False
start = False
End Sub

Private Sub Command3_Click()
Text1.Text = "1"
Text2.Text = "65530"

Text1.Locked = True
Text2.Locked = True
Command1.Enabled = False
Winsock1.Close
start = True
logport.Text = " ***** VIRIIO Port Listener *****"

Call scanningports
End Sub

Private Sub Command4_Click()
TimerOnOff = True
Command4.Enabled = False
Timer1.Interval = secs.Text * 1000
logport.Text = logport.Text & vbCrLf & "Timer is ON"
Command5.Enabled = True
Timer1.Enabled = True
End Sub

Private Sub Command5_Click()
TimerOnOff = False
Command4.Enabled = True
logport.Text = logport.Text & vbCrLf & "Timer is OFF"
Command5.Enabled = False
Timer1.Enabled = False
End Sub

Private Sub mnufile_Click()

End Sub

Private Sub Timer1_Timer()
If TimerOnOff = True Then
TimerOnOff = False
Call Command1_Click
End If
End Sub
