VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.UserControl WBControl 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10725
   ScaleHeight     =   7410
   ScaleWidth      =   10725
   ToolboxBitmap   =   "WBControl.ctx":0000
   Begin VB.ListBox List2 
      Height          =   3840
      Left            =   8520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   -3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3840
      Left            =   6960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   -3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox WBCnt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   0
      Left            =   600
      ScaleHeight     =   4695
      ScaleWidth      =   5775
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5775
      Begin SHDocVwCtl.WebBrowser WB 
         Height          =   4695
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   1200
         Width           =   5775
         ExtentX         =   10186
         ExtentY         =   8281
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.PictureBox MTab 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   615
      ScaleWidth      =   1335
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
      Begin VB.PictureBox Tcbt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   960
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image TIcon 
         Height          =   255
         Index           =   0
         Left            =   120
         OLEDropMode     =   1  'Manual
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Tstr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MTabBox"
         Height          =   180
         Index           =   0
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.Image Icon1 
      Height          =   240
      Left            =   3480
      Picture         =   "WBControl.ctx":0314
      Top             =   -4560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image tc2 
      Height          =   210
      Left            =   2400
      Picture         =   "WBControl.ctx":089E
      Top             =   -3840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image tc1 
      Height          =   210
      Left            =   2040
      Picture         =   "WBControl.ctx":5346
      Top             =   -3840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image tc0 
      Height          =   210
      Left            =   1680
      Picture         =   "WBControl.ctx":9DFB
      Top             =   -3840
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00D3A788&
      X1              =   0
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FBE5D6&
      X1              =   0
      X2              =   3120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FBE5D6&
      X1              =   0
      X2              =   3120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FBE5D6&
      X1              =   0
      X2              =   3120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D3A788&
      X1              =   0
      X2              =   3120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image T2 
      Height          =   420
      Left            =   4440
      Picture         =   "WBControl.ctx":E763
      Top             =   -4080
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image T0 
      Height          =   420
      Left            =   -7080
      Picture         =   "WBControl.ctx":13F3C
      Top             =   5880
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image T1 
      Height          =   420
      Left            =   1560
      Picture         =   "WBControl.ctx":195CB
      Top             =   -4080
      Visible         =   0   'False
      Width           =   2700
   End
End
Attribute VB_Name = "WBControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const WM_SETTEXT = &HC
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim OldIdx As Integer
Event SelectTab()
Event URLChang(Index As Integer, URL As String)
Event ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
Event DownloadComplete(Index As Integer)
Event PropertyChange(Index As Integer, ByVal szProperty As String)
Event DownloadBegin(Index As Integer, URL As Variant)
Event NavigateComplete2(Index As Integer, ByVal pDisp As Object, URL As Variant)
Event TitleChange(Index As Integer, ByVal Text As String)
Event PrintPage()
Event BeforeNavigate2(Index As Integer, URL As Variant)
' 在用户控件的通用声明部分添加：
Private m_TabTitles() As String

Public Sub SaveFile()
WB(OldIdx).ExecWB 4, 1
End Sub

Public Sub PrintView()
WB(OldIdx).ExecWB 7, 2
End Sub

Public Sub GoHome()
WB(OldIdx).GoHome
End Sub

Public Sub RefreshPage()
WB(OldIdx).Refresh
End Sub

Public Sub StopLink()
WB(OldIdx).Stop
End Sub

Public Sub Goforward()
On Error Resume Next
WB(OldIdx).Goforward
End Sub

Public Sub Goback()
On Error Resume Next
WB(OldIdx).Goback
End Sub

Public Sub PrintPage()
WB(OldIdx).ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Public Sub SelectAll()
WB(OldIdx).ExecWB 17, 2
End Sub

Public Sub Getattribute()
WB(OldIdx).ExecWB 10, 2
End Sub

Public Sub cut()
WB(OldIdx).ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
End Sub

Public Sub copy()
WB(OldIdx).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
End Sub

Public Sub paste()
WB(OldIdx).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
End Sub


Public Sub find()
WB(OldIdx).SetFocus
SendKeys "^f"
End Sub

Public Sub GetFile()
Dim doc, objhtml    As Object
Dim i     As Integer
Dim strhtml     As String
Dim nHwnd, eHwnd As Long
If WB(OldIdx).Busy = False Then

  Set doc = WB(OldIdx).Document
  i = 0
  Set objhtml = doc.body.createtextrange()
  If Not IsNull(objhtml) Then
  strhtml = objhtml.htmltext
  
  Shell "Notepad.exe", vbNormalFocus
  nHwnd = FindWindow("Notepad", "无标题 - 记事本")
  eHwnd = FindWindowEx(nHwnd, 0&, "Edit", "")
  SendMessage eHwnd, WM_SETTEXT, 0, ByVal strhtml
  
  End If
End If
End Sub

Public Sub OpenURL(URL As Variant)
Tstr(OldIdx).Caption = URL
If Tstr(OldIdx).Width >= MTab(OldIdx).Width - 600 Then Tstr(OldIdx).Caption = Left(Tstr(OldIdx).Caption, 10)
Tstr(OldIdx).Move (MTab(OldIdx).Width - Tstr(OldIdx).Width) / 2, (MTab(OldIdx).Height - Tstr(OldIdx).Height) / 2
WB(OldIdx).Navigate URL
End Sub

Public Sub NewWindow(URL As String)
On Error Resume Next
Dim ResizeIdx As Integer
If List2.ListCount > 0 Then
    If List1.ListCount > 0 Then
        Tcbt(List2.List(0)).Picture = tc0.Picture
        MTab(List2.List(0)).Move MTab(List1.List(List1.ListCount - 1)).Left + 2700, 0
        MTab(List2.List(0)).Visible = True
        WB(List2.List(0)).Navigate URL
        WBCnt(List2.List(0)).Visible = True
        Call MTab_MouseDown(List2.List(0), 1, 0, 0, 0)
        
        WBCnt(List2.List(0)).Move 0, 480, UserControl.Width, UserControl.Height - 480
        WB(List2.List(0)).Move -20, -20, WBCnt(List2.List(0)).Width + 40, WBCnt(List2.List(0)).Height + 40

        List1.AddItem List2.List(0)
        List2.RemoveItem 0

        Exit Sub
    Else
        Tcbt(List2.List(0)).Picture = tc0.Picture
        MTab(List2.List(0)).Move MTab(0).Left + 2700, 0
        MTab(List2.List(0)).Visible = True
        WB(List2.List(0)).Navigate URL
        WBCnt(List2.List(0)).Move 0, 480, UserControl.Width, UserControl.Height - 480
        
        WBCnt(List2.List(0)).Visible = True
        WB(List2.List(0)).Move -20, -20, WBCnt(List2.List(0)).Width, WBCnt(List2.List(0)).Height
        Call MTab_MouseDown(List2.List(0), 1, 0, 0, 0)
        
        WBCnt(List2.List(0)).Move 0, 480, UserControl.Width, UserControl.Height - 480
        WB(List2.List(0)).Move -20, -20, WBCnt(List2.List(0)).Width + 40, WBCnt(List2.List(0)).Height + 40
        
        List1.AddItem List2.List(0)
        List2.RemoveItem 0
        
        Exit Sub
    End If
    
Else
     If List1.ListCount > 0 Then
             Load MTab(List1.List(List1.ListCount - 1) + 1)
             Load Tstr(List1.List(List1.ListCount - 1) + 1)
             Load TIcon(List1.List(List1.ListCount - 1) + 1)
             Load Tcbt(List1.List(List1.ListCount - 1) + 1)
             Load WBCnt(List1.List(List1.ListCount - 1) + 1)
             Load WB(List1.List(List1.ListCount - 1) + 1)
             
             MTab(List1.List(List1.ListCount - 1) + 1).Picture = T0.Picture
             MTab(List1.List(List1.ListCount - 1) + 1).Move MTab(List1.List(List1.ListCount - 1)).Left + 2700, 0
             MTab(List1.List(List1.ListCount - 1) + 1).Visible = True
        
             Set Tstr(List1.List(List1.ListCount - 1) + 1).Container = MTab(List1.List(List1.ListCount - 1) + 1)
             Tstr(List1.List(List1.ListCount - 1) + 1).Move (MTab(List1.List(List1.ListCount - 1) + 1).Width - Tstr(List1.List(List1.ListCount - 1) + 1).Width) / 2, (MTab(List1.List(List1.ListCount - 1) + 1).Height - Tstr(List1.List(List1.ListCount - 1) + 1).Height) / 2
             Tstr(List1.List(List1.ListCount - 1) + 1).Visible = True
        
             Set TIcon(List1.List(List1.ListCount - 1) + 1).Container = MTab(List1.List(List1.ListCount - 1) + 1)
             TIcon(List1.List(List1.ListCount - 1) + 1).Picture = Icon1.Picture
             TIcon(List1.List(List1.ListCount - 1) + 1).Move 60, (MTab(List1.List(List1.ListCount - 1) + 1).Height - TIcon(List1.List(List1.ListCount - 1) + 1).Height) / 2
             TIcon(List1.List(List1.ListCount - 1) + 1).Visible = True
             
             Set Tcbt(List1.List(List1.ListCount - 1) + 1).Container = MTab(List1.List(List1.ListCount - 1) + 1)
             Tcbt(List1.List(List1.ListCount - 1) + 1).Picture = tc0.Picture
             Tcbt(List1.List(List1.ListCount - 1) + 1).Move (MTab(List1.List(List1.ListCount - 1) + 1).Width - Tcbt(List1.List(List1.ListCount - 1) + 1).Width - 60), (MTab(List1.List(List1.ListCount - 1) + 1).Height - Tcbt(List1.List(List1.ListCount - 1) + 1).Height) / 2
             Tcbt(List1.List(List1.ListCount - 1) + 1).Visible = True
             
             Set WB(List1.List(List1.ListCount - 1) + 1).Container = WBCnt(List1.List(List1.ListCount - 1) + 1)
             WB(List1.List(List1.ListCount - 1) + 1).Move 0, 480, UserControl.Width, UserControl.Height
             WB(List1.List(List1.ListCount - 1) + 1).Visible = True
             
             WBCnt(List1.List(List1.ListCount - 1) + 1).Move 0, 480, UserControl.Width, UserControl.Height - 480
             WBCnt(List1.List(List1.ListCount - 1) + 1).Visible = True
             WB(List1.List(List1.ListCount - 1) + 1).Navigate URL
             WB(List1.List(List1.ListCount - 1) + 1).Move -20, -20, WBCnt(List1.List(List1.ListCount - 1) + 1).Width, WBCnt(List1.List(List1.ListCount - 1) + 1).Height
             WB(List1.List(List1.ListCount - 1) + 1).Visible = True
             Call MTab_MouseDown((List1.List(List1.ListCount - 1) + 1), 1, 0, 0, 0)
             
             WBCnt(List1.List(List1.ListCount - 1) + 1).Move 0, 480, UserControl.Width, UserControl.Height - 480
             WB(List1.List(List1.ListCount - 1) + 1).Move -20, -20, WBCnt(OldIdx).Width + 40, WBCnt(List1.List(List1.ListCount - 1) + 1).Height + 40
             
             List1.AddItem List1.List(List1.ListCount - 1) + 1
             Call UserControl_Resize
             Exit Sub
     Else
             Load MTab(1)
             Load Tstr(1)
             Load TIcon(1)
             Load Tcbt(1)
             Load WBCnt(1)
             Load WB(1)
             
             MTab(1).Picture = T0.Picture
             MTab(1).Move MTab(0).Left + 2700, 0
             MTab(1).Visible = True
        
             Set Tstr(1).Container = MTab(1)
             Tstr(1).Move (MTab(1).Width - Tstr(1).Width) / 2, (MTab(1).Height - Tstr(1).Height) / 2
             Tstr(1).Visible = True
        
             Set TIcon(1).Container = MTab(1)
             TIcon(1).Picture = Icon1.Picture
             TIcon(1).Move 60, (MTab(1).Height - TIcon(1).Height) / 2
             TIcon(1).Visible = True
             
             Set Tcbt(1).Container = MTab(1)
             Tcbt(1).Picture = tc0.Picture
             Tcbt(1).Move (MTab(1).Width - Tcbt(1).Width - 60), (MTab(1).Height - Tcbt(1).Height) / 2
             Tcbt(1).Visible = True
             
             Set WB(1).Container = WBCnt(1)
             WB(1).Move 0, 480, UserControl.Width, UserControl.Height
             WB(1).Visible = True
             
             WBCnt(1).Move 0, 480, UserControl.Width, UserControl.Height - 480
             WBCnt(1).Visible = True
             If Blank = False Then WB(1).Navigate URL
             WB(1).Move -20, -20, WBCnt(1).Width, WBCnt(1).Height
             WB(1).Visible = True
             
             Call MTab_MouseDown(1, 1, 0, 0, 0)
             
             WBCnt(1).Move 0, 480, UserControl.Width, UserControl.Height - 480
             WB(1).Move -20, -20, WBCnt(1).Width + 40, WBCnt(1).Height + 40

             
             
             List1.AddItem "1"
             
             Exit Sub
     End If
End If
End Sub

Private Sub MTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 And OldIdx <> Index Then
MTab(OldIdx).Picture = T1.Picture
Tcbt(OldIdx).Visible = False
WBCnt(OldIdx).Visible = False

MTab(Index).Picture = T0.Picture
WBCnt(Index).Visible = True
If Index <> 0 Then Tcbt(Index).Visible = True
MTab(Index).Top = 0

OldIdx = Index
Call UserControl_Resize
End If
If Button = 1 Then RaiseEvent SelectTab
End Sub

Private Sub MTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If OldIdx <> Index Then
With MTab(Index)
If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
ReleaseCapture
If MTab(Index).Picture <> T1.Picture Then MTab(Index).Picture = T1.Picture
Else
SetCapture .hwnd
If MTab(Index).Picture <> T2.Picture Then MTab(Index).Picture = T2.Picture
End If
End With
End If
End Sub


Private Sub MTab_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Effect = vbDropEffectCopy
If Data.GetFormat(vbCFText) Then
     WB(Index).Navigate Data.GetData(vbCFText)
     Call MTab_MouseDown(Index, 1, Shift, 0, 0)
ElseIf Data.GetFormat(vbCFDIB) Then
        ShowPic.Show
        ShowPic.Picture1.Picture = Data.GetData(vbCFDIB)
End If
End Sub

Private Sub Tcbt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Tcbt(Index).Picture <> tc2.Picture Then Tcbt(Index).Picture = tc2.Picture
End Sub

Private Sub Tcbt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With Tcbt(Index)
If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
ReleaseCapture
If Tcbt(Index).Picture <> tc0.Picture Then Tcbt(Index).Picture = tc0.Picture
Else
SetCapture .hwnd
    If Button = 1 Then
        If Tcbt(Index).Picture <> tc2.Picture Then Tcbt(Index).Picture = tc2.Picture
    Else
        If Tcbt(Index).Picture <> tc1.Picture Then Tcbt(Index).Picture = tc1.Picture
    End If
End If
End With
End Sub

Private Sub CloseTab(Index As Integer)
MTab(Index).Visible = False
WBCnt(Index).Visible = False
WB(Index).Navigate "About:Blank"

If List1.ListCount = 1 Then
List1.RemoveItem List1.ListCount - 1
Call MTab_MouseDown(0, 1, 0, 0, 0)
List2.AddItem Index
Exit Sub
ElseIf List1.List(List1.ListCount - 1) = Index Then
List1.RemoveItem List1.ListCount - 1
Call MTab_MouseDown(List1.List(List1.ListCount - 1), 1, 0, 0, 0)
List2.AddItem Index
GoTo SetTab
End If

Dim K As Integer
For K = 0 To List1.ListCount - 1
If List1.List(K) = Index Then
List2.AddItem Index
List1.RemoveItem K
GoTo SetTab
End If
Next K
Exit Sub
SetTab:
Dim J As Integer
For J = 0 To List1.ListCount - 1
    If J = 0 Then
        MTab(List1.List(J)).Move MTab(0).Left + 2700
    Else
        MTab(List1.List(J)).Move MTab(List1.List(J - 1)).Left + 2700
    End If
Next J
Call MTab_MouseDown(List1.List(List1.ListCount - 1), 1, 0, 0, 0)
End Sub

Private Sub Tcbt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And X >= 0 And X <= Tcbt(Index).Width And Y >= 0 And Y <= Tcbt(Index).Height Then
'请在这里添加关闭选项卡的代码
Call CloseTab(Index)
End If
End Sub

Private Sub Tcbt_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MTab_OLEDragDrop(Index, Data, Effect, Button, Shift, X + Tcbt(Index).Left, Y + Tcbt(Index).Top)
End Sub

Private Sub TIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MTab_MouseDown(Index, Button, Shift, X + TIcon(Index).Left, Y + TIcon(Index).Top)
End Sub


Private Sub TIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MTab_MouseMove(Index, Button, Shift, X + TIcon(Index).Left, Y + TIcon(Index).Top)
End Sub

Private Sub TIcon_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MTab_OLEDragDrop(Index, Data, Effect, Button, Shift, X + TIcon(Index).Left, Y + TIcon(Index).Top)
End Sub

Private Sub Tstr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MTab_MouseDown(Index, Button, Shift, X + Tstr(Index).Left, Y + Tstr(Index).Top)
End Sub

Private Sub Tstr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MTab_MouseMove(Index, Button, Shift, X + Tstr(Index).Left, Y + Tstr(Index).Top)
End Sub

Private Sub Tstr_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MTab_OLEDragDrop(Index, Data, Effect, Button, Shift, X + Tstr(Index).Left, Y + Tstr(Index).Top)
End Sub

Private Sub UserControl_GotFocus()
WB(OldIdx).SetFocus
End Sub

Private Sub UserControl_Initialize()
ReDim m_TabTitles(0)
m_TabTitles(0) = "新标签页" ' 默认标题
OldIdx = 0
MTab(0).Picture = T0.Picture
TIcon(0).Picture = Icon1.Picture
MTab(0).Move 60, 0
Tstr(0).Move (MTab(0).Width - Tstr(0).Width) / 2, (MTab(0).Height - Tstr(0).Height) / 2
TIcon(0).Move 60, (MTab(0).Height - TIcon(0).Height) / 2 + 20
WB(0).GoHome
End Sub

Private Sub UserControl_Resize()
If UserControl.Height > 600 Then
Line1.X1 = 0
Line1.X2 = UserControl.Width
Line1.Y1 = 405
Line1.Y2 = 405

Line2.X1 = 0
Line2.X2 = UserControl.Width
Line2.Y1 = 420
Line2.Y2 = 420

Line3.X1 = 0
Line3.X2 = UserControl.Width
Line3.Y1 = 435
Line3.Y2 = 435

Line4.X1 = 0
Line4.X2 = UserControl.Width
Line4.Y1 = 450
Line4.Y2 = 450

Line5.X1 = 0
Line5.X2 = UserControl.Width
Line5.Y1 = 465
Line5.Y2 = 465
WBCnt(OldIdx).Move 0, 480, UserControl.Width, UserControl.Height - 480
WB(OldIdx).Move -20, -20, WBCnt(OldIdx).Width + 40, WBCnt(OldIdx).Height + 40
End If
End Sub

Private Sub WB_BeforeNavigate2(Index As Integer, ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
WB(Index).Silent = True
RaiseEvent BeforeNavigate2(Index, URL)
End Sub

Private Sub WB_DocumentComplete(Index As Integer, ByVal pDisp As Object, URL As Variant)
On Error Resume Next
    ' 强制更新标题
    If Not WB(Index).Document Is Nothing Then
        RaiseEvent TitleChange(Index, WB(Index).Document.Title)
    End If
        If Index > UBound(m_TabTitles) Then
        ReDim Preserve m_TabTitles(Index)
    End If
    
    If WB(Index).Document Is Nothing Then
        m_TabTitles(Index) = "空白页"
    Else
        m_TabTitles(Index) = WB(Index).Document.Title
        If m_TabTitles(Index) = "" Then m_TabTitles(Index) = WB(Index).LocationURL
    End If
End Sub

Private Sub WB_DownloadBegin(Index As Integer)
RaiseEvent DownloadBegin(Index, WB(Index).LocationURL)
End Sub

Private Sub WB_DownloadComplete(Index As Integer)
RaiseEvent DownloadComplete(Index)
WB(OldIdx).Move -20, -20, WBCnt(OldIdx).Width + 45, WBCnt(OldIdx).Height + 20
If WB(OldIdx).Busy = False Then
Dim a As Object
Set a = WB(OldIdx).Document
'可见宽度= a.body.clientWidth
End If
End Sub

Private Sub WB_NavigateComplete2(Index As Integer, ByVal pDisp As Object, URL As Variant)
RaiseEvent NavigateComplete2(Index, pDisp, URL)
End Sub

Private Sub WB_NewWindow2(Index As Integer, ppDisp As Object, Cancel As Boolean)
Call NewWindow("About:Blank")
Set ppDisp = WB(List1.List(List1.ListCount - 1)).Application
WB(List1.List(List1.ListCount - 1)).ZOrder
End Sub

Private Sub WB_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
RaiseEvent ProgressChange(Index, Progress, ProgressMax)
End Sub

Private Sub WB_PropertyChange(Index As Integer, ByVal szProperty As String)
RaiseEvent PropertyChange(Index, szProperty)
End Sub

Private Sub WB_TitleChange(Index As Integer, ByVal Text As String)
If Text <> "" Then
If UCase(WB(Index).LocationURL) = "ABOUT:BLANK" Then
Tstr(Index).Caption = "空白页"
Else
Tstr(Index).Caption = Text
End If
If Tstr(Index).Width >= MTab(Index).Width - 600 Then Tstr(Index).Caption = Left(Tstr(Index), 10) & "..."
Tstr(Index).Move (MTab(Index).Width - Tstr(Index).Width) / 2, (MTab(Index).Height - Tstr(Index).Height) / 2
End If
If WB(Index).Busy = False Then
RaiseEvent URLChang(Index, WB(Index).LocationURL)
End If
RaiseEvent TitleChange(Index, Text)
End Sub

Public Property Get URL() As Variant
URL = WB(OldIdx).LocationURL
End Property

Public Property Let URL(ByVal New_URL As Variant)
WB(OldIdx).Navigate URL
PropertyChanged "URL"
End Property

Public Property Get Title() As String
On Error Resume Next
Debug.Print "OldIdx=" & OldIdx & ", Document Is Nothing: " & (WB(OldIdx).Document Is Nothing)
If WB(OldIdx).Document Is Noting Then
    Title = "空白页"
Else
        Title = WB(OldIdx).Document.Title
        If Title = "" Then Title = WB(OldIdx).LocationURL ' 回退到 URL
    End If
    If OldIdx >= 0 And OldIdx <= UBound(m_TabTitles) Then
        Title = m_TabTitles(OldIdx)
    Else
        Title = "无效索引"
    End If
'Title = WB(OldIdx).LocationName
End Property

Private Sub WB_WindowClosing(Index As Integer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
Cancel = True
Call CloseTab(Index)
End Sub
