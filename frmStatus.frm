VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmStatus"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar proceso"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1890
      TabIndex        =   3
      Top             =   960
      Width           =   1515
   End
   Begin MSComctlLib.ProgressBar prgExtracc 
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblNums 
      AutoSize        =   -1  'True
      Caption         =   "x of xxx"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   300
      Width           =   525
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Extracting messages: "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   300
      Width           =   1545
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bCancel As Boolean

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    prgExtracc.Value = 0
    Caption = App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bCancel = True
End Sub
