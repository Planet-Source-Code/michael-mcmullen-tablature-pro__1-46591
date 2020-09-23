VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tablature Pro Editor"
   ClientHeight    =   6960
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   9345
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   480
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All Files|*.*"
   End
   Begin VB.Frame fraTablature 
      Caption         =   "Tablature"
      Enabled         =   0   'False
      Height          =   6375
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   5895
      Begin VB.TextBox txtTab 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label lblSongStop 
         AutoSize        =   -1  'True
         Caption         =   "</song>"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   6120
         Width           =   600
      End
      Begin VB.Label lblSongStart 
         AutoSize        =   -1  'True
         Caption         =   "<song>"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Frame fraSongs 
      Caption         =   "Song List"
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   1560
         Width           =   735
      End
      Begin VB.ListBox lstSongs 
         Height          =   840
         ItemData        =   "frmMain.frx":0BC2
         Left            =   120
         List            =   "frmMain.frx":0BC4
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblgtp 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblgtpend 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   480
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const gtpStart As String = "<gtp>"
Const gtpEnd As String = "</gtp>"
Dim strFileName As String
Dim strTabs(0 To 500) As String
Dim intCounter As Integer

Private Sub cmdAdd_Click()
    Dim song As String
    
    song = InputBox("Enter the song name.", "Tablature Pro Editor")
    If song <> "" Then
        lstSongs.AddItem Trim(song)
        frmMain.Caption = "Tablature Pro Editor - Unsaved"
    End If
End Sub

Private Sub Form_Load()
    Call ReInit
End Sub

Private Sub Enables()
    fraSongs.Enabled = True
    fraTablature.Enabled = True
End Sub

Private Sub lstSongs_DblClick()
    fraTablature.Enabled = True
    lblSongStart.Caption = "<" & lstSongs.Text & ">"
    lblSongStop.Caption = "</" & lstSongs.Text & ">"
    intCounter = lstSongs.ListIndex
    txtTab.Text = strTabs(intCounter)
    frmMain.Caption = "Tablature Pro Editor - Unsaved"
    txtTab.Enabled = True
    txtTab.SetFocus
End Sub

Private Sub mnuAbout_Click()
    MsgBox "This program is made for Tablature Pro and will not work with any other program" _
    , vbInformation, "Tablature Pro Editor"
End Sub

Private Sub mnuFileExit_Click()
    If frmMain.Caption = "Tablature Pro Editor - Unsaved" Then
        Dim intAnswer As Integer
        intAnswer = MsgBox("You havent saved your file, do you wish to save it now?", vbYesNo)
        If intAnswer = vbYes Then
            Call mnuFileSave_Click
        End If
    End If
    End
End Sub

Private Sub ReInit()
    Dim i As Integer
    For i = 0 To 500
        strTabs(i) = ""
    Next i
    lstSongs.Clear
    txtTab.Text = ""
    lblgtp.Caption = gtpStart
    lblgtpend.Caption = gtpEnd
    lblSongStart.Caption = "<song>"
    lblSongStop.Caption = "</song>"
    fraTablature.Enabled = False
End Sub

Private Sub mnuFileNew_Click()
    On Error GoTo Cancelled
    Dialog.DialogTitle = "Enter A filename"
    Dialog.DefaultExt = "txt"
    Dialog.ShowSave
    
    Open Dialog.FileName For Output As #1
    strFileName = Dialog.FileName
    Close #1
    Call Enables
    Call ReInit
    Exit Sub
Cancelled:
End Sub


Private Sub mnuFileOpen_Click()
    On Error GoTo Cancelled
    
    Dim strSong As String
    
    Dialog.DialogTitle = "Select a Pro Tab"
    Dialog.ShowOpen
    intCounter = 0
    On Error GoTo NonGFile
    Open Dialog.FileName For Input As #1
    Input #1, strSong
    If UCase(strSong) = "<GTP>" Then
        lstSongs.Clear
        Call ReInit
        Do
            Input #1, strSong
            If UCase(strSong) <> "</GTP>" Then
                lstSongs.AddItem strSong
                intCounter = intCounter + 1
            End If
        Loop Until UCase(strSong) = "</GTP>"
    Else
        GoTo NonGFile
    End If
    Dim intRealCounter As Integer
    intRealCounter = 0
    Do
        lstSongs.Selected(intRealCounter) = True
        Do
            Input #1, strSong
        Loop Until UCase(strSong) = "<" & UCase(lstSongs.Text) & ">"
        
        Do
            Input #1, strSong
            
            If strSong = "" Then
                strTabs(intRealCounter) = strTabs(intRealCounter) & vbNewLine
            End If
            If UCase(strSong) <> "</" & UCase(lstSongs.Text) & ">" Then
                strTabs(intRealCounter) = strTabs(intRealCounter) & strSong
                strTabs(intRealCounter) = strTabs(intRealCounter) & vbNewLine
            End If
        Loop Until UCase(strSong) = "</" & UCase(lstSongs.Text) & ">"
        
        intRealCounter = intRealCounter + 1
    Loop Until intRealCounter >= intCounter
    Close #1
    strFileName = Dialog.FileName
    Call Enables
    
    Exit Sub
Cancelled:
    Exit Sub
NonGFile:
    MsgBox "Sorry this file is not in the Tab Pro Format", vbCritical
    Close #1
End Sub

Private Sub mnuFileSave_Click()
    Dim lstCounter As Integer

    Open strFileName For Output As #1
    Print #1, gtpStart
    For lstCounter = 0 To lstSongs.ListCount - 1
        lstSongs.Selected(lstCounter) = True
        Print #1, lstSongs.Text
    Next lstCounter
    Print #1, gtpEnd
    
    For lstCounter = 0 To lstSongs.ListCount - 1
        lstSongs.Selected(lstCounter) = True
        Print #1, "<" & lstSongs.Text & ">"
        Print #1, strTabs(lstCounter)
        Print #1, "</" & lstSongs.Text & ">"
    Next lstCounter
    Close #1
    frmMain.Caption = "Tablature Pro Editor - Saved"
End Sub

Private Sub txtTab_Change()
    strTabs(intCounter) = txtTab.Text
End Sub
