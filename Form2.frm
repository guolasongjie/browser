VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "选项"
   ClientHeight    =   3405
   ClientLeft      =   7515
   ClientTop       =   4920
   ClientWidth     =   7170
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7170
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "清除历史记录"
      Height          =   255
      Left            =   420
      TabIndex        =   7
      Top             =   2100
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "添加搜索引擎"
      Height          =   315
      Left            =   4140
      TabIndex        =   6
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CheckBox Check3 
      Caption         =   "跟随IE主页设置"
      Height          =   180
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   255
      Left            =   4380
      TabIndex        =   4
      Top             =   2940
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2940
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   6015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "不显示“脚本错误”对话框"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "主页："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub Command1_Click()
write1 = WritePrivateProfileString("mainpage", "url", Text1.Text, App.Path & "\ibrowser.ini")
  If Check3.Value = 1 Then
   write1 = WritePrivateProfileString("mainpage", "flowie", "1", App.Path & "\ibrowser.ini")
    ElseIf Check3.Value = 0 Then
   write1 = WritePrivateProfileString("mainpage", "flowie", "0", App.Path & "\ibrowser.ini")
  End If
  If Check1.Value = 1 Then
  write1 = WritePrivateProfileString("errmsg", "dontshow", "1", App.Path & "\ibrowser.ini")
  End If
  

Dim seachname As String
Dim seachurl As String
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
seachname = InputBox("输入搜索服务名称")
seachurl = InputBox("输入搜索参数地址")


End Sub

Private Sub Command4_Click()
Kill App.Path & "\History.htm"
End Sub
