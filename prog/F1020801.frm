VERSION 5.00
Begin VB.Form F1020801 
   BackColor       =   &H00FFFFFF&
   Caption         =   "入庫リストデータ出力 "
   ClientHeight    =   6960
   ClientLeft      =   2028
   ClientTop       =   2568
   ClientWidth     =   11292
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11292
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Caption         =   "対象倉庫選択"
      Height          =   1095
      Index           =   1
      Left            =   2760
      TabIndex        =   16
      Top             =   2760
      Width           =   2775
      Begin VB.CheckBox Check1 
         Caption         =   "81"
         Height          =   495
         Index           =   4
         Left            =   2040
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "80"
         Height          =   495
         Index           =   3
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "90"
         Height          =   495
         Index           =   2
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "対象倉庫選択"
      Height          =   1095
      Index           =   0
      Left            =   2760
      TabIndex        =   13
      Top             =   1320
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "80"
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "90"
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "データ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   6480
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1020801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Y_NYU_DATA  As String                   '入荷データフルパス

Private Const Last_Update_Day$ = "([Ｆ102080] 2011.07.14 12:00)"

Private Sub Command_Click(Index As Integer)
Dim yn  As Integer


    Select Case Index
        Case 7
            
            If Frame1(0).Visible Then
                If Check1(0).Value Or Check1(1).Value Then
                Else
                    MsgBox "対象倉庫を選択してください。"
                    Exit Sub
                End If
            Else
                If Check1(2).Value Or Check1(3).Value Or Check1(4).Value Then
                Else
                    MsgBox "対象倉庫を選択してください。"
                    Exit Sub
                End If
            End If
            
            yn = MsgBox("入庫リストデータ出力しますか？", vbYesNo, "確認入力")
            If yn = vbYes Then
            
                If Output_Proc() Then
                    Unload Me
                End If
            
            End If
            
        Case 11
            Unload Me
    End Select

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)


                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1020801.Caption = "入庫リストデータ出力（" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '入庫リストデータファイル名取り込み
    If GetIni("FILE", "Y_NYU_DATA", "SYS", c) Then
        Beep
        MsgBox "入庫リストデータファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    Y_NYU_DATA = Trim(c)
                                '入荷データファイルOPEN
    If Y_NYU_Open(0) Then
        Unload Me
    End If
                                '在庫データファイルOPEN
    If ZAIKO_Open(0) Then
        Unload Me
    End If
                                '月平均データファイルOPEN
    If AVE_SYUKA_Open(0) Then
        Unload Me
    End If
    
    
    If JGYOBU_T(0).CODE = SOJIKI Then
        Frame1(0).Visible = True
        Frame1(1).Visible = False
    Else
        Frame1(0).Visible = False
        Frame1(1).Visible = True
    End If

    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '入荷データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷データファイル")
        End If
    End If
                                            '在庫データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データファイル")
        End If
    End If
                                            '月平均出荷ＣＬＯＳＥ
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "月平均出荷")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020801 = Nothing

    End
End Sub


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1020801.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020801)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020801)


    F1020801.MousePointer = vbDefault

End Sub


Private Function Output_Proc() As Integer
'----------------------------------------------------------------------------
'                   ＣＳＶデータ出力処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim Ret             As Integer

Dim c               As String * 128

Dim i               As Integer

Dim FileNo          As Integer
Dim fileName        As String

Dim Den_Date        As String * 8
Dim Rec_Cnt         As Integer

Dim Skip_Flg        As Boolean

Dim Work_Soko       As String * 2
Dim Soko_No         As String * 2

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim Fast_Flg        As Boolean

    Output_Proc = True
'実行中はイベント取得不可
    Call Input_Lock         '画面項目ロック

    FileNo = FreeFile
    fileName = Y_NYU_DATA


    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & Right(Format(Now, "YYYY/MM/DD"), 2) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo

    Rec_Cnt = 0

    Fast_Flg = True
    com = BtOpGetFirst


    Do
        DoEvents
        
        sts = BTRV(com, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)

        Select Case sts
            Case BtNoErr
                If StrConv(Y_NYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
                '対象倉庫の判定
                Skip_Flg = True
                
                
                If StrConv(Y_NYUREC.NYU_LIST_OUT, vbUnicode) = "9" Then
                Else
                    Select Case Last_JGYOBU
                        Case SOJIKI                     '滋賀
                                                
                            
                            Select Case Trim(StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode))
                            
                                Case "91H"
                                    Work_Soko = "90"
                                Case Else
                                    Work_Soko = "80"
                            End Select
                        
                        
                            For i = 0 To 1
                                If Check1(i).Value Then
                                    If Trim(Check1(i).Caption) = Work_Soko Then
                                        Skip_Flg = False
                                        Exit For
                                    End If
                                End If
                            Next i
                        
                        
                        Case DENKA, SUIHAN, SENTAKU     '小野
                    
                            Select Case Trim(StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode))
                                Case "G22"
                                    Work_Soko = "80"
                                Case "G11"
                                    Work_Soko = "81"
                                Case Else
                                    Work_Soko = "90"
                            End Select
                    
                    
                            For i = 2 To 4
                                If Check1(i).Value Then
                                    If Trim(Check1(i).Caption) = Work_Soko Then
                                        Skip_Flg = False
                                        Exit For
                                    End If
                                End If
                            Next i
                    
                    End Select
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "入荷予定データ")
                Exit Function
        End Select

        If Not Skip_Flg Then
            
            
            
            Rec_Cnt = Rec_Cnt + 1

            If Fast_Flg Then
                Den_Date = StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode)
                Write #FileNo, , , "入庫リスト", , "作成日：", Left(Den_Date, 4) & "/" & Mid(Den_Date, 5, 2) & "/" & Right(Den_Date, 2) & "分"
                Write #FileNo, "標準棚番", "品番（外部）", "品番（内部）", "伝票№", "入庫数", "入庫数－前借数", "予算単位", "荷姿", "入庫先", "未商品", "商品化済み", "月平均"
                Fast_Flg = False
            End If

            '標準棚番
            If GetIni("SOKO_NO", Left(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 2), "SYS", c) Then
                Soko_No = Left(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 2)
            Else
                Soko_No = Trim(c)
            End If
                    
            Write #FileNo, Soko_No & "-" & Mid(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 3, 2) & "-" & Mid(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 5, 2) & "-" & Mid(StrConv(Y_NYUREC.HTANABAN, vbUnicode), 7, 2),
            Write #FileNo, StrConv(Y_NYUREC.HIN_NO, vbUnicode),
            Write #FileNo, StrConv(Y_NYUREC.HIN_NAI, vbUnicode),
            Write #FileNo, StrConv(Y_NYUREC.DEN_NO, vbUnicode),
            Write #FileNo, Format(CLng(StrConv(Y_NYUREC.SURYO, vbUnicode)), "#0"),
            Write #FileNo, Format(CLng(StrConv(Y_NYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode)), "#0"),
            Write #FileNo, StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode) & " " & StrConv(Y_NYUREC.YOSAN_TO, vbUnicode), , ,
                    
        
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    StrConv(Y_NYUREC.JGYOBU, vbUnicode), _
                                    StrConv(Y_NYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_NYUREC.HIN_NO, vbUnicode)) Then
                Exit Function
            End If
        
            Write #FileNo, Format(MI_QTY, "#0"),
            Write #FileNo, Format(SUMI_QTY, "#0"),
        
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            
            Select Case sts
                Case BtNoErr
                    Write #FileNo, Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#0")
                Case BtErrKeyNotFound
                    Write #FileNo,
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "月平均出荷数")
                    Exit Function
            End Select
        
        End If


        com = BtOpGetNext
    Loop



    Write #FileNo, "伝票日付：", Left(Den_Date, 4) & "/" & Mid(Den_Date, 5, 2) & "/" & Right(Den_Date, 2), "データ件数：", Format(Rec_Cnt, "#0")


    Close #FileNo
    
    Call Input_UnLock         '画面項目ロック解除
    Beep
    MsgBox "「" & fileName & "」は正常に出力されました。"
    
    
    
    
    Call Input_UnLock         '画面項目ロック解除
    Output_Proc = False

    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "が使用中です。"
        Call Input_UnLock         '画面項目ロック解除
        Output_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Output_Proc = True
    End If

End Function

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1020801.Caption = "入庫リストデータ出力（" + RTrim(JGYOBU_T(Index).NAME) + ")" & Last_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub
