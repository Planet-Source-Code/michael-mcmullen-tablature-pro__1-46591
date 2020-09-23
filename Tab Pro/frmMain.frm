VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablature Pro - Unregistered"
   ClientHeight    =   6075
   ClientLeft      =   2280
   ClientTop       =   1395
   ClientWidth     =   11550
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTab 
      Caption         =   "Tablature"
      Height          =   5895
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Frame fraSong 
      Caption         =   "Song"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3135
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Grab"
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox cboSongs 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame fraFile 
      Caption         =   "Current File"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3135
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   45
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSeparater1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsFont 
         Caption         =   "Font..."
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuAboutAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFileName As String

Private Sub cboSongs_Change()
    cmdGo.SetFocus
End Sub

Private Sub cmdAddsong_Click()
    Dim strSongName As String
    strSongName = InputBox("Enter a song title", "Tablature Pro")
    cboSongs.AddItem strSongName
End Sub

Private Sub cmdGo_Click()
    If cboSongs.Text <> "" Then
        Call FindSong
        txtTab.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call ReadIni
    If blnRegistered = True Then
        frmMain.Caption = "Tablature Pro"
        mnuAboutRegister.Visible = False
    End If
End Sub

Private Sub mnuAboutAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAboutRegister_Click()
    frmRegister.Show
End Sub

Private Sub mnuFileNew_Click()
    On Error GoTo Cancelled
    Dialog.DialogTitle = "Enter a file name"
    Dialog.ShowSave
    Open Dialog.FileName For Output As #1
    
    txtTab.Text = ""
    cboSongs.Clear
    
    Close #1
    Exit Sub
Cancelled:
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo Cancelled
    
    Dim strSong As String

    Dialog.CancelError = True
    Dialog.Filter = "All Files|*.*"
    Dialog.DialogTitle = "Select a Pro Tab"
    Dialog.ShowOpen
    
    On Error GoTo NonGFile
    Open Dialog.FileName For Input As #1
    Input #1, strSong
    If UCase(strSong) = "<GTP>" Then
        cboSongs.Clear
        
        Do
            Input #1, strSong
            If UCase(strSong) <> "</GTP>" Then
                cboSongs.AddItem strSong
            End If
        Loop Until UCase(strSong) = "</GTP>"
    Else
        GoTo NonGFile
    End If
    
    Close #1
    lblFileName.Caption = Dialog.FileTitle
    strFileName = Dialog.FileName
    Exit Sub
Cancelled:
    Exit Sub
NonGFile:
    MsgBox "Sorry this file is not in the Tab Pro Format", vbCritical, "Tablature Pro"
    Close #1
End Sub
Private Sub FindSong()
    On Error GoTo CannotFind

    Dim strSongTitle As String
    Dim strTab As String
    
    Open strFileName For Input As #1
    Input #1, strSongTitle
    Do Until UCase(strSongTitle) = "</GTP>"
        Input #1, strSongTitle
    Loop
    Do Until UCase(strSongTitle) = "<" & UCase(cboSongs.Text) & ">"
        Input #1, strSongTitle
    Loop
    txtTab.Text = ""
    Do
        Input #1, strTab
        If strTab = "" Then
            txtTab.Text = txtTab.Text & vbNewLine
        End If
        If UCase(strTab) <> "</" & UCase(cboSongs.Text) & ">" Then
            txtTab.Text = txtTab.Text & strTab
            txtTab.Text = txtTab.Text & vbNewLine
        End If
    Loop Until UCase(strTab) = "</" & UCase(cboSongs.Text) & ">"
    Close #1
    Exit Sub
    
CannotFind:
    Close #1
    MsgBox "Sorry, the song " & cboSongs.Text & ", is not in this file.", vbCritical, "Tablature Pro"
End Sub

Private Sub mnuFilePrint_Click()
    If txtTab.Text <> "" Then
    On Error GoTo Cancelled
    Dim oldFont As String
    Dim oldSize As Integer
    
    Dialog.CancelError = True
    Dialog.ShowPrinter
    oldFont = Printer.Font
    oldSize = Printer.FontSize
    Printer.Font = strFont
    Printer.FontSize = intFontSize
    
    
    Printer.Print txtTab.Text
    Exit Sub
Cancelled:
End If
End Sub

Private Sub mnuOptionsFont_Click()
    On Error GoTo Cancelled

    Dialog.CancelError = True
    Dialog.DialogTitle = "Please Select A Font"
    Dialog.ShowFont
    
    strFont = Dialog.FontName
    intFontSize = Dialog.FontSize
    txtTab.Font = strFont
    txtTab.FontSize = intFontSize
    txtTab.Refresh
    WriteIni
    Exit Sub
    
Cancelled:

End Sub
