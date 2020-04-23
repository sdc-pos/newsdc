VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMENU 
   BackColor       =   &H00808080&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "メニュー"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMENU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12225
   StartUpPosition =   2  '画面の中央
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   840
      MaskColor       =   &H8000000F&
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   0
      Top             =   360
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   1
      Top             =   960
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   2
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   2
      Top             =   1560
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   3
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   3
      Top             =   2160
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   4
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   4
      Top             =   3240
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   5
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   5
      Top             =   3840
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   6
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   6
      Top             =   4440
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   7
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   7
      Top             =   5040
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   8
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   8
      Top             =   6120
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   9
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   9
      Top             =   6720
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   10
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   10
      Top             =   7320
      Width           =   10572
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   11
      Left            =   840
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   11
      Top             =   7920
      Width           =   10572
   End
   Begin VB.PictureBox Picture1 
      Height          =   252
      Left            =   11760
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label L_Pass 
      BackColor       =   &H80000009&
      Caption         =   "Pass"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label L_Pass 
      BackColor       =   &H80000009&
      Caption         =   "Pass"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Lab_Mi 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '実線
      Caption         =   "未確定色"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   7.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   204
      Left            =   10680
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   672
   End
   Begin VB.Label Lab_Sumi 
      AutoSize        =   -1  'True
      BorderStyle     =   1  '実線
      Caption         =   "通常色"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   7.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   204
      Left            =   10200
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   516
   End
   Begin VB.Label Lab_Default 
      BackColor       =   &H00808080&
      BorderStyle     =   1  '実線
      Height          =   252
      Left            =   11400
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   372
   End
End
Attribute VB_Name = "frmMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Form_Caption$ = "サンプルメニュー"
Private Const LAST_UPDATE_DAY$ = "[SDCMENU] 2017.01.10 09:30"

Dim Menu_Page As Integer               'ﾒﾆｭｰﾍﾟｰｼﾞ№
Dim Focus_Pos As Integer

Private Type task_tbl_tag
    Task_No     As String
    Task_Title  As String
End Type
Dim task_tbl(11) As task_tbl_tag    '表示中ﾒﾆｭｰの起動可能ﾀｽｸ
Dim Menu_tbl(11) As String          'ﾒﾆｭｰONLY

Dim Command_Color_LOS As Long       '未選択時のボタンの色
Dim Command_Color_GOT As Long       '選択時のボタンの色




Private Function Ko_Menu_Set_Proc() As Integer
Dim i As Integer
Dim c As String * 128
Dim Wk As String
Dim Max_Disp_Len As Integer

    Ko_Menu_Set_Proc = True


    'パスワード確認
    If GetIni("MENU" + Format(Menu_Page, "0"), "PASSWD", "TITLE", c) Then
        Me.L_Pass(0).Caption = ""
    Else
        Me.L_Pass(0).Caption = Trim(c)
    End If
    If Me.L_Pass(0).Caption <> "" And Me.L_Pass(0).Caption <> "/" Then
        Me.L_Pass(1).Caption = ""
        frmPASS.Show vbModal
        If Me.L_Pass(1).Caption = "" Then   'ｷｬﾝｾﾙ？
            Ko_Menu_Set_Proc = False
            Exit Function
        End If
    End If


    'メニュータイトル表示
    If GetIni("MENU" + Format(Menu_Page, "00"), "M_TITLE", "TITLE", c) Then
        Me.Caption = Form_Caption & LAST_UPDATE_DAY
    Else
        Me.Caption = Trim(c) & LAST_UPDATE_DAY
    End If

    'メニュー背景色 設定
    If GetIni("MENU" + Format(Menu_Page, "0"), "COLOR", "TITLE", c) Then
        Me.BackColor = Lab_Default.BackColor
    Else
        If IsNumeric(Trim(c)) Then
            Me.BackColor = CLng(Trim(c))
        Else
            Me.BackColor = Lab_Default.BackColor
        End If
    End If
    Lab_Sumi.BackColor = Me.BackColor

    'ボタンタイトル表示
    Max_Disp_Len = 0
    For i = 0 To 11
        If GetIni("MENU" + Format(Menu_Page, "00"), "TITLE" + Format(i, "00"), "TITLE", c) Then
            Beep
            MsgBox "システム異常発生！！ 起動できません。[ﾒﾆｭｰﾌｧｲﾙ 取得ｴﾗｰ]", vbCritical
            Exit Function
        End If
        If Right(RTrim(c), 1) = ":" Then
            Command1(i).Enabled = False                                   'ﾌｧﾝｸｼｮﾝ未使用
        Else
            Command1(i).Enabled = True                                    'ﾌｧﾝｸｼｮﾝ未使用
        End If

        Command1(i).Caption = RTrim(c)                                    'ﾌｧﾝｸｼｮﾝ未使用
        If LenB(StrConv(RTrim(c), vbFromUnicode)) > Max_Disp_Len Then
            Max_Disp_Len = LenB(StrConv(RTrim(c), vbFromUnicode))
        End If

        If GetIni("MENU" + Format(Menu_Page, "00"), "EXE" + Format(i, "00"), "TITLE", c) Then
            Beep
            MsgBox "システム異常発生！！ 起動できません。[ﾒﾆｭｰﾌｧｲﾙ 取得ｴﾗｰ]", vbCritical
            Exit Function
        End If
        task_tbl(i).Task_No = RTrim(c)
                                    '実行可能なプログラムのタイトルを取り込む
        If IsNumeric(task_tbl(i).Task_No) Then
            If GetIni("TITLE", RTrim(task_tbl(i).Task_No), "TITLE", c) Then
                Beep
                MsgBox "システム異常発生！！ 起動できません。[ﾒﾆｭｰﾌｧｲﾙ 取得ｴﾗｰ]", vbCritical
                Exit Function
            End If
            task_tbl(i).Task_Title = RTrim(c)
        End If

    Next i

    '表示長を揃える
    For i = 0 To 11
        Wk = Command1(i).Caption
        If LenB(StrConv(Wk, vbFromUnicode)) < Max_Disp_Len Then
            Command1(i).Caption = Wk & _
                        Space(Max_Disp_Len - LenB(StrConv(Wk, vbFromUnicode)))
        End If
    Next i


    Ko_Menu_Set_Proc = False

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    Me.MousePointer = vbHourglass

    Call Ctrl_Lock(Me)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(Me)

    Me.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)
Dim TASK_ID As Long
Dim Yn As Integer
Dim sts As Integer
Dim i As Integer
On Error Resume Next


Dim PrcSet As SWbemObjectSet
Dim Prc As SWbemObject
Dim Locator As SWbemLocator
Dim Service As SWbemServices
Dim MesStr As String


Dim wkProg_ID   As String
Dim FSW         As Integer




    '指定したKEYは無効
    If RTrim(task_tbl(Index).Task_No) = "NON" Then
        Beep
        GoTo Command1_Click_Exit
    End If

    '終了指定なら
    If RTrim(task_tbl(Index).Task_No) = "END" Then
        Unload Me
        GoTo Command1_Click_Exit
    End If

    'システム終了指定なら(ﾒﾆｭｰ単独で使用する場合）
    If RTrim(task_tbl(Index).Task_No) = "SYSEND" Then
        Yn = MsgBox("システムを終了します。" & Chr(13) & Chr(10) & _
                    Chr(13) & Chr(10) & _
                    "宜しいですか？", vbDefaultButton2 + vbYesNo + vbQuestion)
            If Yn = vbNo Then GoTo Command1_Click_Exit
        Unload Me
        GoTo Command1_Click_Exit
    End If

    '子メニューへの切替えなら
    If Left(task_tbl(Index).Task_No, 4) = "MENU" Then
        Menu_Page = CInt(Right(task_tbl(Index).Task_No, 2))
        Focus_Pos = Index
        If Ko_Menu_Set_Proc Then
            Unload Me
            GoTo Command1_Click_Exit
        End If
        Command1(11).SetFocus
        For i = 0 To 11
            If Command1(i).Enabled Then
                Command1(i).SetFocus
                Exit For
            End If
        Next i
        GoTo Command1_Click_Exit
    End If

    'ﾌﾟﾛｸﾞﾗﾑ起動指定なら
    

''>>>>>>>>>>>>>>>>
'Set Locator = New WbemScripting.SWbemLocator
'Set Service = Locator.ConnectServer
'
'Set PrcSet = Service.ExecQuery("Select * From Win32_Process")
'
'
'For i = Len(RTrim(task_tbl(Index).Task_No)) To 1 Step -1
'
'
'    If Mid(task_tbl(Index).Task_No, i, 1) = "\" Then
'        Exit For
'    End If
'
'
'Next i
'
'
'wkProg_ID = Mid(task_tbl(Index).Task_No, i + 1, Len(task_tbl(Index).Task_No) - i)
'
'
'FSW = 0
'For Each Prc In PrcSet
'
'
'Debug.Print Prc.Subclasses_
'
'
'
'    If Trim(Str(Prc.Description)) = Trim(wkProg_ID) Then
'        FSW = 1
'        Exit For
'    End If
'
'Next
'
'Set PrcSet = Nothing
'Set Prc = Nothing
'Set Service = Nothing
'Set Locator = Nothing
'
'If FSW = 1 Then
'    GoTo Command1_Click_Exit
'End If
'
'
''>>>>>>>>>>>>>>>>
    
    
    Call Input_Lock
    Shell RTrim(task_tbl(Index).Task_No), vbNormalFocus
    Call Input_UnLock

Command1_Click_Exit:


End Sub

Private Sub Command1_GotFocus(Index As Integer)

    Command1(Index).BackColor = Command_Color_GOT

End Sub

Private Sub Command1_LostFocus(Index As Integer)
    Command1(Index).BackColor = Command_Color_LOS

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sts As Integer
Dim c As String

    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            If KeyCode = vbKeyF12 And _
               Shift = vbShiftMask + vbCtrlMask Then
                CommonDialog1.COLOR = Me.BackColor
                CommonDialog1.ShowColor
                Me.BackColor = CommonDialog1.COLOR
                c = CStr(CommonDialog1.COLOR)
                sts = WriteIni("MENU" + Format(Menu_Page, "0"), "COLOR", "TITLE", c)
            Else
                If Command1(KeyCode - vbKeyF1).Enabled = True Then
                    Command1(KeyCode - vbKeyF1).SetFocus
                    Command1(KeyCode - vbKeyF1).Value = True
                End If
            End If
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c As String * 128

'二重起動防止（起動時間が掛かる為）
    If App.PrevInstance Then
        End
    End If

'ｸﾞﾛｰﾊﾞﾙ項目　初期化
'    If Global_Init Then
'        Unload Me
'        Exit Sub
'    End If

'画面初期設定
'    me.Caption = Form_Caption
    Menu_Page = 1

    If Ko_Menu_Set_Proc Then
        Unload Me
        Exit Sub
    End If

    Command_Color_LOS = vbButtonFace
    Command_Color_GOT = vbHighlightText

'フォーム表示
    Show
    DoEvents

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If UnloadMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmMENU = Nothing       'ﾌｫｰﾑ資源解放
    End
End Sub
