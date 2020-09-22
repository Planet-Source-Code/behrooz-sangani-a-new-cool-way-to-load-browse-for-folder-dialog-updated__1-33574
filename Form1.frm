VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Using Shell32.dll  developed on WindowsXP   by Behrooz Sangani <bs20014@yahoo.com>"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Explorer"
      Height          =   2775
      Left            =   6240
      TabIndex        =   25
      Top             =   4080
      Width           =   3495
      Begin VB.CommandButton Command15 
         Caption         =   "Go To My Website"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Explore"
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Open"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "c:\"
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Enter a folder path"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "System"
      Height          =   3855
      Left            =   6240
      TabIndex        =   12
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command12 
         Caption         =   "Show Shut Down Dialog"
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   3120
         Width           =   2655
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Show System Time Properties"
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Show Taskbar Properties"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   2400
         Width           =   2655
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Show Computer Search Dialog"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Show File Search Dialog"
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Show Help"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Show Run Dialog"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Load System Properties"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Window Functions"
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   6015
      Begin VB.CommandButton Command6 
         Caption         =   "Undo Minimize All"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Tile Windows Vertically"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Tile Windows Horizentally"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   2280
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Minimize All Windows"
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cascade All Windows"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Browse For Folder"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   960
         Width           =   5775
      End
      Begin VB.TextBox Text2 
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1560
         Width           =   5775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Select a folder"
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Folder Details:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================
'Using shell32.dll
'======================================
'
'By: Behrooz Sangani
'Email: bs20014@yahoo.com
'Web: http://www.geocities.cm/bs20014/
'
'Use and modify for free!
'======================================
'Some Info:
'I have developed this code on WindowsXP and I am
'not sure if it works on earlier versions of Windows.
'I will test this in a few days but since it works perfectly
'on XP I decided to share it. Browse the code for
'yourself and remember the old one page module needed
'to load Browse For Folder Dialog and ...
'Here it just needs a reference to shell32.dll located
'in system directory.
'
'Sorry if the code is not well organized. :) After all it
'has been written yesterday.
'
'Please leave comments and vote for this!
'Thanks!
'======================================

Dim SH As New Shell  'reference to shell32.dll class
Dim ShBFF As Folder  'Shell Browse For Folder

Private Sub Command1_Click() 'Show BFF Dialog
On Error Resume Next
'set object
Set ShBFF = SH.BrowseForFolder(hWnd, "Hey this is a sample, " & _
            "please choose a folder and click OK!", 1)
With ShBFF.Items.Item
   'get folder props
   Text1 = .Path
   Text2 = "Name: " & .Name & vbCrLf & _
           "Type: " & .Type & vbCrLf & _
           "Last Modified: " & .ModifyDate & vbCrLf & _
           "Parent: " & .Parent & vbCrLf
End With

End Sub

Private Sub Command10_Click() 'Show help
  SH.Help
End Sub

Private Sub Command11_Click() 'Show compute find dialog
  SH.FindComputer
End Sub

Private Sub Command12_Click() 'Show shut down dialog

  If MsgBox("Are you sure you want to do this!?", _
     vbQuestion + vbYesNo + vbDefaultButton2, _
     "Confirm Action!") <> vbYes Then Exit Sub
     
    SH.ShutdownWindows

End Sub

Private Sub Command13_Click() 'Open path
  SH.Open Text3.Text
End Sub

Private Sub Command14_Click() 'Explore path
  SH.Explore Text3.Text
End Sub

Private Sub Command15_Click() 'Open URL
  SH.Open "http://www.geocities.com/bs20014/"
End Sub

Private Sub Command16_Click() 'Show Taskbar & Start Menu Properties
  SH.TrayProperties
End Sub

Private Sub Command17_Click() 'Show clock dialog
  SH.SetTime
End Sub

Private Sub Command2_Click() 'Cascade windows
  SH.CascadeWindows
End Sub

Private Sub Command3_Click() 'Minimize all windows
  SH.MinimizeAll
End Sub

Private Sub Command4_Click() 'Tile windows horizentally
  SH.TileHorizontally
End Sub

Private Sub Command5_Click() 'Tile windows vertically
  SH.TileVertically
End Sub

Private Sub Command6_Click() 'Undo minimizing windows
  SH.UndoMinimizeALL
End Sub

Private Sub Command7_Click() 'Load something in Control Panel
  SH.ControlPanelItem "sysdm.cpl" 'System Properties
  'This can be easily changed
  'Search for *.cpl in your system directory
  'inetcpl.cpl    ==> Internet Options
  'appwiz.cpl     ==> Add/Remove Programs
  'and many more...
End Sub

Private Sub Command8_Click() 'Find files dialog
  SH.FindFiles
End Sub

Private Sub Command9_Click() 'Run dialog
  SH.FileRun
End Sub

Private Sub Form_Load() 'Form Load
 Text3 = App.Path
End Sub
