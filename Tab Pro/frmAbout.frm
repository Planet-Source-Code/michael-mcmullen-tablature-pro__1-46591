VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "About Tablature Pro"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6495
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   570
   End
   Begin VB.Label lblReg 
      AutoSize        =   -1  'True
      Caption         =   "This program is unregistered"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload frmAbout
End Sub

Private Sub Form_Load()
    If blnRegistered = True Then
        lblReg.Caption = "This program is registered to " & strUserName
    End If
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor
End Sub
