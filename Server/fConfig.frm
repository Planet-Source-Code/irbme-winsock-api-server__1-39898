VERSION 5.00
Begin VB.Form fConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SocketWise Server Configuration"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2235
   Icon            =   "fConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLog 
      Caption         =   "Log Activity?"
      Height          =   330
      Left            =   105
      TabIndex        =   4
      Top             =   630
      Value           =   1  'Checked
      Width           =   2010
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   540
      Left            =   1155
      TabIndex        =   3
      Top             =   1155
      Width           =   960
   End
   Begin VB.CommandButton cmdReConnect 
      Caption         =   "Reconnect"
      Height          =   540
      Left            =   105
      TabIndex        =   2
      Top             =   1155
      Width           =   960
   End
   Begin VB.TextBox txtLocalPort 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   1155
      TabIndex        =   1
      Text            =   "4000"
      Top             =   210
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Local Port"
      Height          =   330
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1485
   End
End
Attribute VB_Name = "fConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
End Sub


Private Sub cmdReConnect_Click()
    Me.Hide
    
    fServer.ServerPort = Int(txtLocalPort.Text)
    
    fServer.DisconnectServer
    fServer.ConnectServer
End Sub
