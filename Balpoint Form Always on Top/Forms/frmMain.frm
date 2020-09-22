VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Set Form Always on Top"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFalse 
      Caption         =   "AlwaysOnTop (False)"
      Height          =   435
      Left            =   2610
      TabIndex        =   1
      Top             =   795
      Width           =   1875
   End
   Begin VB.CommandButton cmdTrue 
      Caption         =   "AlwaysOnTop (True)"
      Height          =   435
      Left            =   2610
      TabIndex        =   2
      Top             =   375
      Width           =   1875
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: #"
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   4395
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFalse_Click()
'   MAKE THE FORM ALWAYS ON TOP(FALSE)
    MakeAlwaysOnTop Me, False
    lblStatus.Caption = "Status: AlwaysOnTop(False)"
End Sub
Private Sub cmdTrue_Click()
'   SET ALWAYS ON TOP(TRUE)
    MakeAlwaysOnTop Me, True
    lblStatus.Caption = "Status: AlwaysOnTop(True)"
End Sub

Private Sub Form_Load()
'   MAKE THE FORM ALWAYS ON TOP(FALSE)
    MakeAlwaysOnTop Me, False
    lblStatus.Caption = "Status: AlwaysOnTop(False)"
End Sub
