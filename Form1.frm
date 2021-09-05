VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Internet Browser"
   ClientHeight    =   8070
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   14865
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   14865
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   4
      Top             =   7695
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25691
         EndProperty
      EndProperty
   End
   Begin ibrowser.WBControl WBControl1 
      Height          =   6420
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   11324
   End
   Begin VB.CommandButton Command1 
      Caption         =   "转到"
      Height          =   255
      Left            =   13920
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   34
      ImageHeight     =   34
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D416
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E238
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":F3C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":FFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":119E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":128FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":138A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14672
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   1640
      ButtonWidth     =   1217
      ButtonHeight    =   1482
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "后退"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前进"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "停止"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刷新"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "主页"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "搜索"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "收藏"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "历史"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新标签"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.ComboBox serchcombox 
         Height          =   300
         Left            =   9180
         TabIndex        =   8
         Text            =   "选择搜索引擎"
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "搜索一下"
         Height          =   300
         Left            =   13920
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   10680
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
      Begin VB.Timer Timer1 
         Left            =   6120
         Top             =   0
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   900
      TabIndex        =   0
      Top             =   900
      Width           =   12975
   End
   Begin VB.Label Label1 
      Caption         =   "地址(&D)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Menu file 
      Caption         =   "文件(&F)"
      Begin VB.Menu open 
         Caption         =   "打开(&O)…"
      End
      Begin VB.Menu save 
         Caption         =   "保存(&S)…"
      End
      Begin VB.Menu print 
         Caption         =   "打印(&P)…"
      End
      Begin VB.Menu pview 
         Caption         =   "打印预览(&V)…"
      End
      Begin VB.Menu close 
         Caption         =   "关闭标签(&C)"
      End
      Begin VB.Menu shuxing 
         Caption         =   "属性(&R)"
      End
      Begin VB.Menu exit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu cut 
         Caption         =   "剪切(&T)"
      End
      Begin VB.Menu copy 
         Caption         =   "复制(&C)"
      End
      Begin VB.Menu paste 
         Caption         =   "粘贴(&P)"
      End
      Begin VB.Menu find 
         Caption         =   "查找(&F)…"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu view 
      Caption         =   "查看(&V)"
      Begin VB.Menu socfile 
         Caption         =   "源文件(&C)"
      End
   End
   Begin VB.Menu fav 
      Caption         =   "收藏(&F)"
   End
   Begin VB.Menu tools 
      Caption         =   "工具(&T)"
      Begin VB.Menu option 
         Caption         =   "选项(&I)"
      End
      Begin VB.Menu IE_option 
         Caption         =   "Internet 选项(&O)"
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu hlpfl 
         Caption         =   "帮助文档"
      End
      Begin VB.Menu about 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public HomeAddress As String
Dim strURL As String

    Private Declare Function LaunchInternetControlPanel Lib "inetcpl.cpl" (ByVal hwndParent As Long) As Long
    Private Declare Function LaunchConnectionDialog Lib "inetcpl.cpl" (ByVal hwndParent As Long) As Long
    Dim Title As String
    

Private Sub about_Click()
Form3.Show
End Sub


Private Sub Combo1_Change()
Timer1.Enabled = False
End Sub

Private Sub Combo1_Click()
Timer1.Enabled = False
End Sub

Private Sub Combo1_GotFocus()
Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
WBControl1.OpenURL Combo1.Text
End Sub

Private Sub copy_Click()
WBControl1.copy
End Sub

Private Sub cut_Click()
WBControl1.cut

End Sub

Private Sub find_Click()
WBControl1.find
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    WBControl1.Width = Form1.Width - 250
    WBControl1.Height = Form1.Height - 1500
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer) '使用组合框浏览
Timer1.Enabled = False
If KeyAscii = 13 Then
Combo1.AddItem Combo1.Text
If Combo1.ListCount > 10 Then '防止组合框内容过多
Combo1.RemoveItem 1
End If
WBControl1.OpenURL Combo1.Text
End If
End Sub
 
Private Sub exit_Click()
    End
End Sub
 
Private Sub Form_Load()
    Title = " - Internet Browser"
   ' HomeAddress = "www.bilibili.com" '填写主页地址
    'WBControl1.Navigate HomeAddress
    'Combo1.Text = "http://" & HomeAddress
    Timer1.Interval = 500
End Sub

Private Sub hlpfl_Click()
On Error Resume Next
Shell App.Path & "\hlp.chm"
End Sub
Private Sub IE_option_Click() '调用IE选项
    Dim rc As Long
    rc = LaunchInternetControlPanel(Me.hwnd)
    If rc = 0 Then
    MsgBox "LaunchInternetControlPanel failed!", vbExclamation
    End If
End Sub

Private Sub open_Click()
strURL = InputBox("http://", "输入地址")
WBControl1.OpenURL strURL
End Sub

Private Sub option_Click()
Form2.Show
End Sub

Private Sub paste_Click()
WBControl1.paste
End Sub

Private Sub print_Click()
WBControl1.PrintPage
End Sub

Private Sub pview_Click()
WBControl1.PrintView
End Sub

Private Sub save_Click()
WBControl1.SaveFile
End Sub

Private Sub shuxing_Click()
WBControl1.Getattribute
End Sub

Private Sub socfile_Click()
WBControl1.GetFile
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
WBControl1.Goback
Case 2
WBControl1.Goforward
Case 3
WBControl1.StopLoading
Case 4
WBControl1.Refresh
Case 5
WBControl1.GoHome
Case 6
WBControl1.OpenURL "www.baidu.com/s?word=" & Text1.Text
Case 8
WBControl1.OpenURL App.Path & "\History.htm"
Case 9
WBControl1.NewWindow "about:blank"
End Select
End Sub

Private Sub WBControl1_URLChang(Index As Integer, URL As String)
Combo1.Text = WBControl1.URL
Me.Caption = WBControl1.LocationName & Title
End Sub
Private Sub WBControl1_SelectTab()
Combo1.Text = WBControl1.URL
Me.Caption = WBControl1.LocationName & Title
End Sub
Private Sub Timer1_Timer()
Combo1.Text = WBControl1.URL
Me.Caption = WBControl1.LocationName & Title
Timer1.Enabled = True
End Sub
 'Private Sub wbcontrol1_NewWindow2(ppDisp As Object, Cancel As Boolean)
' Cancel = True
' WBControl1.Navigate strURL
' End Sub
 Private Sub WBControl1_StatusTextChange(ByVal Text As String)
 StatusBar1.Panels(1) = WBControl1.Text
 End Sub
 Private Sub WBControl1_DownloadBegin(Index As Integer, URL As Variant)
 StatusBar1.Panels(1) = "（剩下1 项）" & WBControl1.URL
 End Sub
 Private Sub WBControl1_DownloadComplete(Index As Integer)
 Dim Text As String
 Dim a As String
 StatusBar1.Panels(1) = "完毕"
 Timer1.Enabled = True
 Combo1.Text = WBControl1.URL
 Me.Caption = WBControl1.LocationName & Title
 If Dir(App.Path & "\History.htm") <> "" Then
 Open App.Path & "\History.htm" For Input As #1
 Do Until EOF(1)
 Line Input #1, a
 Text = Text & a
 Loop
 Close #1
 End If
 Open App.Path & "\History.htm" For Output As #1
 Print #1, Text & "<br><a href=" & WBControl1.URL & ">" & "标题："; WBControl1.LocationName; "，时间：" & Date & ","; Time; "</a>"
 Close #1
 End Sub
 

