VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Register Tablature Pro"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2393
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   495
      Left            =   1073
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame fraUser 
      Caption         =   "Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblRegCode 
         AutoSize        =   -1  'True
         Caption         =   "Registration Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Full Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmRegister
End Sub

Private Sub cmdRegister_Click()
    Call CreateReg
    If UCase(strRegCode) = UCase(txtCode.Text) Then
        Call SaveSetting("Tablature Pro", "Main", "User Name", txtName.Text)
        Call SaveSetting("Tablature Pro", "Main", "Registration Code", txtCode.Text)
        frmMain.Caption = "Tablature Pro"
        Unload frmRegister
    Else
        MsgBox "Sorry your registration code is incorrect.", vbCritical, "Tablature Pro"
        txtCode.SetFocus
    End If
End Sub

Private Sub Form_Load()

End Sub

Private Sub txtCode_GotFocus()
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode)
End Sub
