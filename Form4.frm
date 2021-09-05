VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "联系方式"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3945
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox 带分音符的西里尔文小写字母U 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3600
      Picture         =   "Form4.frx":D416
      ScaleHeight     =   180
      ScaleWidth      =   120
      TabIndex        =   5
      Top             =   4740
      Width           =   120
   End
   Begin VB.PictureBox 西里尔文小写字母Schwa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2880
      Picture         =   "Form4.frx":D578
      ScaleHeight     =   135
      ScaleWidth      =   120
      TabIndex        =   4
      Top             =   4740
      Width           =   120
   End
   Begin VB.PictureBox qrcode_for_gh_20239ee9aa38_258 
      AutoSize        =   -1  'True
      Height          =   3930
      Left            =   0
      Picture         =   "Form4.frx":D692
      ScaleHeight     =   3870
      ScaleWidth      =   3870
      TabIndex        =   0
      Top             =   780
      Width           =   3930
   End
   Begin VB.Line Line3 
      X1              =   3900
      X2              =   3840
      Y1              =   4860
      Y2              =   4920
   End
   Begin VB.Line Line2 
      X1              =   3900
      X2              =   3780
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line1 
      X1              =   3900
      X2              =   3720
      Y1              =   4740
      Y2              =   4920
   End
   Begin VB.Label zh_jvn 
      AutoSize        =   -1  'True
      Caption         =   "шуйбувеншачуангф  нгы"
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   4740
      Width           =   3600
   End
   Begin VB.Label Label2 
      Caption         =   "2、"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "1、通过作者邮箱alvaroivanov@qq.com"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1755
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub 西里尔文小写字母Schwa_Click()
MsgBox "w'ajj（拼音，亲属称呼）"
End Sub
