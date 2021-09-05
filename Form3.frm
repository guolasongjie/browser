VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于 Internet Browser"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6180
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form3.frx":D416
   NegotiateMenus  =   0   'False
   Picture         =   "Form3.frx":2E9C0
   ScaleHeight     =   5175
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command2 
      Caption         =   "联系作者"
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   300
      Left            =   5220
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   $"Form3.frx":4C32A
      Height          =   1275
      Left            =   1080
      MouseIcon       =   "Form3.frx":4C452
      TabIndex        =   4
      Top             =   3180
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "版权所有 (C) 2021 Alvaro Ivanov."
      Height          =   195
      Left            =   1080
      MouseIcon       =   "Form3.frx":6D9FC
      TabIndex        =   2
      Top             =   2700
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "版本 2.0.5"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   2220
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Internet Browser"
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   1740
      Width           =   1515
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Command2_Click()
Form4.Show
End Sub



Private Sub Label4_Click()
Form1.WBControl1.OpenURL "http://www.tiancao.net/blogview.asp?logID=860"
End Sub


