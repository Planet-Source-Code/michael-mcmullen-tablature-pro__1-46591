VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUser 
      Caption         =   "Information"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Full Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   750
      End
      Begin VB.Label lblRegCode 
         AutoSize        =   -1  'True
         Caption         =   "Registration Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Get Code"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strUserName As String


Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdRegister_Click()
    strUserName = txtName.Text
    Dim strPass As String
    Dim strletter As String
    Dim i As Integer
    strPass = "TP"
    strletter = Left(strUserName, 1)
    
    For i = 1 To Len(strUserName)
        If strletter = "" Then
            Exit For
        End If
        strPass = strPass & (Str(Asc(strletter)) * Len(strUserName))
        strletter = Right(strUserName, Len(strUserName) - i)
    Next i
     txtCode = strPass
End Sub
