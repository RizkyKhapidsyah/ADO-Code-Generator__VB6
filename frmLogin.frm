VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password"
   ClientHeight    =   2175
   ClientLeft      =   3825
   ClientTop       =   3750
   ClientWidth     =   3435
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1285.063
   ScaleMode       =   0  'User
   ScaleWidth      =   3225.278
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   180
      TabIndex        =   5
      Top             =   1020
      Width           =   3174
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1140
      TabIndex        =   2
      Top             =   1680
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2295
      TabIndex        =   3
      Top             =   1680
      Width           =   1020
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1140
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "The Database is password protected, this program requires  a correct password to continue."
      Height          =   735
      Left            =   660
      TabIndex        =   4
      Top             =   180
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":014A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   210
      Index           =   1
      Left            =   165
      TabIndex        =   0
      Top             =   1260
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sPassword As String
Private bCancel   As Boolean

Public Property Let Password(sPass As String)
  sPassword = sPass
End Property

Public Property Get Password() As String
  Password = sPassword
End Property

Public Property Let NoMore(bCan As Boolean)
  bCancel = bCan
End Property

Public Property Get NoMore() As Boolean
  NoMore = bCancel
End Property


Private Sub cmdCancel_Click()
  Me.NoMore = True
  sPassword = vbNullString
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  If Len(txtPassword.Text) = 0 Then
    txtPassword.SetFocus
  Else
    Me.Password = txtPassword.Text
    Me.Hide
  End If
End Sub

Private Sub Form_Activate()
  txtPassword.SetFocus
  txtPassword.Text = vbNullString
  bCancel = False
  sPassword = vbNullString
    
End Sub

