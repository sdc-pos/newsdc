VERSION 5.00
Begin VB.Form F1010451 
   BackColor       =   &H00FFFFFF&
   Caption         =   "初期設定在庫登録（To 標準棚）（小野用）"
   ClientHeight    =   6312
   ClientLeft      =   1920
   ClientTop       =   2436
   ClientWidth     =   11220
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6312
   ScaleWidth      =   11220
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
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
      Caption         =   "実  行"
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
   Begin VB.Label CntLab 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Z,ZZZ,ZZ9 件"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5520
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label CntLab 
      Alignment       =   1  '右揃え
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Z,ZZZ,ZZ9 件"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   5520
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label CMsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "無効件数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5880
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label CMsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "電化調理機器"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2640
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Label CMsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "炊飯機器"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "在庫初期設定が終了しました！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   4
      Left            =   2520
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   6720
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "在庫初期設定中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   3
      Left            =   3840
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "在庫設定データ読込み中！"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   2
      Left            =   2880
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "初期設定在庫の登録を行います"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   0
      Left            =   2280
      TabIndex        =   13
      Top             =   2400
      Visible         =   0   'False
      Width           =   6720
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "『実行』を選択して下さい"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   1
      Left            =   2760
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   5760
   End
End
Attribute VB_Name = "F1010451"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'#################################################################################################
'［テキストファイル処理の注意！！］
'
'　　このプログラムでは、テキストファイルをユーザー定義型の構造体で読書きしている。
'　　この形式によるＩ－Ｏは、「GET」および「PUT」ｽﾃｰﾄﾒﾝﾄにより行う事になるが、これらのｽﾃｰﾄﾒﾝﾄ
'　　を使用する為には、「RANDOM」または「BINARY」モードで「OPEN」しなければならない。
'　　但し、以下のロジックでは、テキストファイルの存在をチェックする目的で「INPUT」モードによる
'　　「OPEN」も行っており、存在チェック後、すぐに読込みを行う様な場合には、一旦「CLOSE」してから
'　　「BINARYモードでOPEN」している事に注意が必要。
'
'　　※．「INPUT」モードOPENでは、「INPUT#」「OUTPUT#」のみ使用できるが、これらは構造体のメンバ
'　　　　単位のＩ－Ｏになる。
'#################################################################################################
Private Type YUKO_SOKO_TBL4             '有効ﾎｽﾄ倉庫取り込みテーブル（炊飯機器事業部）
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type
Dim SOKO_T4(0 To 9) As YUKO_SOKO_TBL4

Private Type YUKO_SOKO_TBLD             '有効ﾎｽﾄ倉庫取り込みテーブル（電化調理機器事業部）
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type
Dim SOKO_TD(0 To 9) As YUKO_SOKO_TBLD

Dim HS_NaiG As String                   '国内外（決定内容）････　ﾎｽﾄﾃﾞｰﾀ内容により設定
Dim WS_NO As String * 2                 'ﾜｰｸｽﾃｰｼｮﾝ番号

Dim PRT_CAN As Boolean                  '印刷途中キャンセル要求
Dim NormalFont As New StdFont           '印刷フォント
Dim B_Jgyobu As String

Dim Lcnt As Integer
Const LMAX% = 47

Dim Command_Max As Integer              '画面項目別最大ｲﾝﾃﾞｯｸｽ

Dim NG_QTY4 As Long                     '無効データ件数（炊飯機器）
Dim NG_QTYD As Long                     '　　　　　　　（電化調理機器）
                                    '画面初期表示（処理済「便」表示など）
Private Sub Scr_Init()
Dim i As Integer
Dim sts As Integer
Dim Work As String

    MsgLab(ZERO).Visible = True
    MsgLab(1).Visible = True
    MsgLab(2).Visible = False
    MsgLab(3).Visible = False
    Command(ZERO).SetFocus

End Sub
                                            '在庫初期設定 処理
Private Sub Zaiko_IniSet_Main()
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim Work As String
Dim Cp_From As String
Dim Cp_To As String

    Call Input_Lock           '画面項目ロック

    MsgLab(ZERO).Visible = False
    MsgLab(1).Visible = False
    MsgLab(2).Visible = True        '取込み中ﾒｯｾｰｼﾞ表示
    DoEvents

    If Data_Load = False Then       '便別取込みワークにﾃﾞｰﾀﾛｰﾄﾞ
        MsgLab(2).Visible = False       '取込み中ﾒｯｾｰｼﾞ ｸﾘｱ
        MsgLab(3).Visible = True        '在庫登録中ﾒｯｾｰｼﾞ表示
        DoEvents

        If WK_ZAI_Open(1, Work) Then    '取込みﾜｰｸ OPEN（更新ﾓｰﾄﾞ)
            Unload Me
        End If
    
        If ER_ZAI_Open(ZERO, Work) = False Then    '取込みﾜｰｸ有無ﾁｪｯｸ（有り→削除）
            Close #ER_ZAI_No
            Kill Work
        End If

        If ER_ZAI_Open(1, Work) Then            '取込みﾜｰｸ OPEN（更新ﾓｰﾄﾞ)
            Close #ER_ZAI_No
            Unload Me
        End If
                            '印刷フォント設定
        Printer.Font.NAME = F1010451.Font.NAME
        Printer.Font.Size = F1010451.Font.Size

        Lcnt = 99
        B_Jgyobu = Space(1)

        Do
            
            If WK_ZAI_Get Then          '取込みﾜｰｸ 読込み
                Exit Do
            End If
            If StrConv(WK_ZAIREC.JGYOBU, vbUnicode) < " " Then      ' EOF ?
                Exit Do
            End If

            If Data_Chk = False Then
                Call Zaiko_Set          '初期設定在庫　登録
            End If
        
            DoEvents
        Loop

        Close #WK_ZAI_No                '取込みﾜｰｸ CLOSE
        Close #ER_ZAI_No
        MsgLab(3).Visible = False       '在庫登録中ﾒｯｾｰｼﾞ ｸﾘｱ
    End If

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If

'画面項目ロック解除　＆　終了メッセージ表示
    MsgLab(ZERO).Visible = False
    MsgLab(1).Visible = False
    MsgLab(2).Visible = False
    MsgLab(3).Visible = False
    MsgLab(4).Visible = True

    CMsgLab(ZERO).Visible = True
    CMsgLab(1).Visible = True
    CMsgLab(2).Visible = True

    CntLab(ZERO).Visible = True
    CntLab(ZERO).Caption = Format(NG_QTY4, "0") & " 件"
    CntLab(1).Visible = True
    CntLab(1).Caption = Format(NG_QTYD, "0") & " 件"
    
    Call Input_UnLock

End Sub
                                            'データロード(各業部別ﾎｽﾄﾃﾞｰﾀ→取込みﾜｰｸ）
Private Function Data_Load() As Integer
Dim sts As Integer
Dim Work As String
Dim Zai4_F As Integer
Dim ZaiD_F As Integer

    Data_Load = False
    Zai4_F = True
    ZaiD_F = True
    NG_QTY4 = ZERO
    NG_QTYD = ZERO

'取込みワーク　ＯＰＥＮ
    If WK_ZAI_Open(ZERO, Work) = False Then    '取込みﾜｰｸ有無ﾁｪｯｸ（有り→削除）
        Close #WK_ZAI_No
        Kill Work
    End If

    If WK_ZAI_Open(1, Work) Then            '取込みﾜｰｸ OPEN（更新ﾓｰﾄﾞ)
        Unload Me
    End If

'炊飯機器事業部  在庫設定データ取込み
    If HS_ZAI_Open4(ZERO, Work) = False Then   'ﾌｧｲﾙ無しなら処理しない
        Close #HS_ZAI_No

        If HS_ZAI_Open4(1, Work) Then       '炊飯機器事業部ﾎｽﾄﾃﾞｰﾀ OPEN
            Close #WK_ZAI_No
            Unload Me
        End If

        Call Data_Load_Sub                  '炊飯機器事業部  在庫設定データ→取込みﾜｰｸ

        Close #HS_ZAI_No                    'ﾎｽﾄﾃﾞｰﾀ CLOSE
        Zai4_F = False
    End If

'電化調理機器事業部  在庫設定データ取込み
    If HS_ZAI_OpenD(ZERO, Work) = False Then   'ﾌｧｲﾙ無しなら処理しない
        Close #HS_ZAI_No

        If HS_ZAI_OpenD(1, Work) Then       '電化調理機器事業部ﾎｽﾄﾃﾞｰﾀ OPEN
            Close #WK_ZAI_No
            Unload Me
        End If

        Call Data_Load_Sub                  '電化調理機器事業部  在庫設定データ→取込みﾜｰｸ

        Close #HS_ZAI_No                    'ﾎｽﾄﾃﾞｰﾀ CLOSE
        Zai4_F = False
    End If

    Close #WK_ZAI_No                    '取込みﾜｰｸ CLOSE

    If Zai4_F = True And ZaiD_F = True Then
        Data_Load = True
    End If
End Function
                                            '事業部別データロード(事業部別ﾎｽﾄ在庫ﾃﾞｰﾀ→取込みﾜｰｸ）
Private Sub Data_Load_Sub()
Dim i As Integer
Dim Put_Sel As Integer
Dim sts As Integer
Dim Work As String

    Do
        If HS_ZAI_Get Then          'ﾎｽﾄﾃﾞｰﾀ 読込み
            Exit Do
        End If
        If StrConv(HS_ZAIREC.JGYOBU, vbUnicode) < " " Then
            Exit Do
        End If
        
                                '事業部区分
        Call UniCode_Conv(WK_ZAIREC.JGYOBU, StrConv(HS_ZAIREC.JGYOBU, vbUnicode))
                                '倉庫区分（ﾎｽﾄ）
        Call UniCode_Conv(WK_ZAIREC.HOST_SOKO, StrConv(HS_ZAIREC.HOST_SOKO, vbUnicode))
                                '品番（外部）
        Call UniCode_Conv(WK_ZAIREC.HIN_GAI, StrConv(HS_ZAIREC.HIN_GAI, vbUnicode))
                                '品番（内部）
        Call UniCode_Conv(WK_ZAIREC.HIN_NAI, StrConv(HS_ZAIREC.HIN_NAI, vbUnicode))
                                '品名
        Call UniCode_Conv(WK_ZAIREC.HIN_NAME, StrConv(HS_ZAIREC.HIN_NAME, vbUnicode))
                                '棚番（ﾎｽﾄ）
        Call UniCode_Conv(WK_ZAIREC.HOST_TANA, StrConv(HS_ZAIREC.HOST_TANA, vbUnicode))
                                '数量サイン
        Call UniCode_Conv(WK_ZAIREC.QTY_SIGN, StrConv(HS_ZAIREC.QTY_SIGN, vbUnicode))
                                '前日在庫数
        Call UniCode_Conv(WK_ZAIREC.ZEN_Z_QTY, StrConv(HS_ZAIREC.ZEN_Z_QTY, vbUnicode))
                                'FILLER
        Call UniCode_Conv(WK_ZAIREC.FILLER, StrConv(HS_ZAIREC.FILLER, vbUnicode))
                                'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
        Call UniCode_Conv(WK_ZAIREC.REC_END, StrConv(HS_ZAIREC.REC_END, vbUnicode))
                                'CR.LF
        Call UniCode_Conv(WK_ZAIREC.CR_LF, StrConv(HS_ZAIREC.CR_LF, vbUnicode))

        sts = WK_ZAI_Put                    '取込みﾜｰｸ書込み（対象倉庫）
        If sts Then
            Close #HS_ZAI_No
            Close #WK_ZAI_No
            Unload Me
        End If
    Loop

End Sub
                                            '取込みデータ 項目内容チェック
Private Function Data_Chk() As Integer
Dim sts As Integer
Dim i As Integer
Dim Command As Integer
Dim Work As String

    Data_Chk = False

'事業部区分　範囲外？
        For i = ZERO To UBound(JGYOBU_T) - 1
            If JGYOBU_T(i).Code = " " Then
                Data_Chk = True
                Exit For
            End If
            If JGYOBU_T(i).Code = StrConv(WK_ZAIREC.JGYOBU, vbUnicode) Then
                Exit For
            End If
        Next i

'ホスト倉庫　範囲外？
        If StrConv(WK_ZAIREC.JGYOBU, vbUnicode) = "4" Then
            For i = ZERO To UBound(SOKO_T4) - 1
                If SOKO_T4(i).HS_SOKO = "  " Then
                    Data_Chk = True
                    Exit For
                End If
                If RTrim(StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_T4(i).HS_SOKO) Then
                    Exit For
                End If
            Next i
        Else
            For i = ZERO To UBound(SOKO_TD) - 1
                If SOKO_TD(i).HS_SOKO = "  " Then
                    Data_Chk = True
                    Exit For
                End If
                If RTrim(StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_TD(i).HS_SOKO) Then
                    Exit For
                End If
            Next i
        End If
        
'品番（外部）：＝空白？　＆　品番（内部）：＝空白？
'    If StrConv(WK_ZAIREC.HIN_GAI, vbUnicode) = Space(13) Or _
'       StrConv(WK_ZAIREC.HIN_NAI, vbUnicode) = Space(13) Then
'            Data_Chk = True
'    End If

'数量サイン　：≠空白？
    If StrConv(WK_ZAIREC.QTY_SIGN, vbUnicode) <> " " Then
        Data_Chk = True
    End If

    If Data_Chk = True Then
'        Call P_Err_Print             'エラーリスト印刷
        sts = ER_ZAI_Put                    '取込みﾜｰｸ書込み（対象外倉庫）
        If sts Then
            Close #WK_ZAI_No
            Close #ER_ZAI_No
            Unload Me
        End If
        If StrConv(WK_ZAIREC.JGYOBU, vbUnicode) = "4" Then
            NG_QTY4 = NG_QTY4 + 1
        Else
            NG_QTYD = NG_QTYD + 1
        End If
    End If

End Function
                                            '初期設定在庫　登録
Private Sub Zaiko_Set()
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim Wqty As Long
Dim W_SOKO As String
Dim W_RETU As String
Dim W_REN As String
Dim W_DAN As String
Dim Work As String
Dim i As Integer

'国内外区分の設定
    If StrConv(WK_ZAIREC.JGYOBU, vbUnicode) = "4" Then
        For i = 0 To UBound(SOKO_T4)
            If SOKO_T4(i).HS_SOKO = "  " Then
                Exit For
            End If
            If RTrim(StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_T4(i).HS_SOKO) Then
                HS_NaiG = SOKO_T4(i).NAIGAI
                Exit For
            End If
        Next i
    Else
        For i = ZERO To UBound(SOKO_TD)
            If SOKO_TD(i).HS_SOKO = "  " Then
                Exit For
            End If
            If RTrim(StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode)) = RTrim(SOKO_TD(i).HS_SOKO) Then
                HS_NaiG = SOKO_TD(i).NAIGAI
                Exit For
            End If
        Next i
    End If

'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If

'品目マスタ更新
    Call Upd_Item(W_SOKO, W_RETU, W_REN, W_DAN)


    If CLng(StrConv(WK_ZAIREC.ZEN_Z_QTY, vbUnicode)) = 0 Then   '在庫数＝０→在庫登録無し
        GoTo End_Tran
    End If

    If Nyuko_Update_Proc(StrConv(WK_ZAIREC.JGYOBU, vbUnicode), _
                        HS_NaiG, _
                        StrConv(WK_ZAIREC.HIN_GAI, vbUnicode), _
                        Format(Date, "yyyymmdd"), _
                        W_SOKO & W_RETU & W_REN & W_DAN, _
                        YOIN_TU_NYUKA, _
                        CLng(StrConv(WK_ZAIREC.ZEN_Z_QTY, vbUnicode)), _
                        WS_NO) Then
        GoTo Abort_Tran
    End If



End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    Exit Sub

Abort_Tran:
                                        'トランザクション異常終了
    sts = BTRV(BtOpAbortTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    Unload Me
End Sub
                                            '品目マスタ更新
Private Function Upd_Item(WSOKO As String, _
                            WRETU As String, _
                            WREN As String, _
                            WDAN As String) As Integer
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim Work As String

    Upd_Item = True
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(WK_ZAIREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, HS_NaiG)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(WK_ZAIREC.HIN_GAI, vbUnicode))
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr
                Command = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                Command = BtOpInsert
                Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(WK_ZAIREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(ITEMREC.NAIGAI, HS_NaiG)
                Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(WK_ZAIREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(ITEMREC.HIN_NAME, Space(25))
                Call UniCode_Conv(ITEMREC.ST_SET_DT, Space(8))
                Call UniCode_Conv(ITEMREC.ST_SOKO, Space(2))
                Call UniCode_Conv(ITEMREC.ST_RETU, Space(2))
                Call UniCode_Conv(ITEMREC.ST_REN, Space(2))
                Call UniCode_Conv(ITEMREC.ST_DAN, Space(2))
                Call UniCode_Conv(ITEMREC.BEF_SOKO, Space(2))
                Call UniCode_Conv(ITEMREC.BEF_RETU, Space(2))
                Call UniCode_Conv(ITEMREC.BEF_REN, Space(2))
                Call UniCode_Conv(ITEMREC.BEF_DAN, Space(2))
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, Space(8))
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, Space(8))
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, Space(2))
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, Space(8))
                Call UniCode_Conv(ITEMREC.SIZAI_CD, Space(5))
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, Space(8))
                
                Call UniCode_Conv(ITEMREC.LOCK_F, "0")          '排他フラグ
                Call UniCode_Conv(ITEMREC.WEL_ID, "")           '使用中子機ＩＤ
                Call UniCode_Conv(ITEMREC.PRG_ID, "")           '使用中プログラム
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "0000000")
                
                Call UniCode_Conv(ITEMREC.FILLER, Space(32))
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
    Loop

                                '入荷先倉庫の決定
    If StrConv(ITEMREC.ST_SET_DT, vbUnicode) = Space(8) Then
        WSOKO = KASO_NYUKA_Soko
        WRETU = "01"
        WREN = "01"
        WDAN = "01"
    Else
        WSOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        WRETU = StrConv(ITEMREC.ST_RETU, vbUnicode)
        WREN = StrConv(ITEMREC.ST_REN, vbUnicode)
        WDAN = StrConv(ITEMREC.ST_DAN, vbUnicode)
    End If

                                '品名（≠空白の時のみセット）
    If StrConv(WK_ZAIREC.HIN_NAME, vbUnicode) <> Space(25) Then
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(WK_ZAIREC.HIN_NAME, vbUnicode))
    End If
                                '品番（内部）
    If StrConv(WK_ZAIREC.HIN_NAI, vbUnicode) <> Space(13) Then
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(WK_ZAIREC.HIN_NAI, vbUnicode))
    End If

                                '備考：ﾎｽﾄ倉庫区分
                                '　　［倉庫区分の読替え　設定条件]
                                '　　　①受信データの倉庫区分≠空白
                                '　　　②　　〃　　　倉庫区分＞品目ﾏｽﾀの倉庫区分
    If StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode) <> Space(2) And _
       StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode) > StrConv(ITEMREC.BIKOU_SOKO, vbUnicode) Then
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode))
    End If

                                '備考：ﾎｽﾄ棚番（≠空白の時）
    If StrConv(WK_ZAIREC.HOST_TANA, vbUnicode) <> Space(8) Then
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(WK_ZAIREC.HOST_TANA, vbUnicode))
    End If

                                '最終入荷日：入荷数≠０の時設定
    If CLng(StrConv(WK_ZAIREC.ZEN_Z_QTY, vbUnicode)) <> ZERO Then
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, Left(Date, 4) & Mid(Date, 6, 2) & Right(Date, 2))
    End If

                                'ｻﾝﾌﾟﾙ数　救済ﾛｼﾞｯｸ（ﾃﾞﾌｫﾙﾄ＝１の為、内容＝０なら「１」を設定）
    If CLng(StrConv(ITEMREC.SAMPLE_QTY, vbUnicode)) = 0 Then
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")
    End If

    Do
        sts = BTRV(Command, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr, BtErrEOF, BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, Command, "品目マスタ")
                Exit Function
        End Select
    Loop

    Upd_Item = False

End Function
                                            'エラーリストヘッダー印刷
Private Sub P_Err_Head()

    If Lcnt <> 99 Then
        Printer.NewPage
    End If
                                        'ヘッダ印刷
    Printer.Print
    Printer.Print Tab(31);
    Printer.Print "＊＊＊　在庫設定エラーリスト　＊＊＊";
    Printer.Print Tab(91);
    Printer.Print Date & "     P." & Format$(Printer.Page, "000")
    Printer.Print
                                        '明細ヘッダ印刷
    Printer.Print Tab(2);
    Printer.Print "事業部";
    Printer.Print Tab(25);
    Printer.Print "倉庫";
    Printer.Print Tab(31);
    Printer.Print "品番（外部）";
    Printer.Print Tab(46);
    Printer.Print "品番（内部）";
    Printer.Print Tab(61);
    Printer.Print "品　名";
    Printer.Print Tab(88);
    Printer.Print "棚番";
    Printer.Print Tab(98);
    Printer.Print "±";
    Printer.Print Tab(104);
    Printer.Print "在庫数";
    Printer.Print
        
    Printer.Print
    Lcnt = 6

End Sub
                                            'エラーリスト明細印刷
Private Sub P_Err_Print()

Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim i As Integer
Dim sts As Integer

                                        'ヘッダーコントロール
    If Lcnt > LMAX Then
        Call P_Err_Head
        B_Jgyobu = Space(1)
    End If
                                '事業部区分
    If StrConv(WK_ZAIREC.JGYOBU, vbUnicode) <> B_Jgyobu Then
        B_Jgyobu = StrConv(WK_ZAIREC.JGYOBU, vbUnicode)
        Printer.Print Tab(2);
        Printer.Print StrConv(WK_ZAIREC.JGYOBU, vbUnicode);

        For i = ZERO To UBound(JGYOBU_T) - 1
            If JGYOBU_T(i).Code = " " Then
                Exit For
            End If
            If JGYOBU_T(i).Code = StrConv(WK_ZAIREC.JGYOBU, vbUnicode) Then
                Printer.Print Tab(4);
                Printer.Print RTrim(JGYOBU_T(i).NAME);
                Exit For
            End If
        Next i
    End If
    
    Printer.Print Tab(26);      '倉庫区分（ﾎｽﾄ）
    Printer.Print StrConv(WK_ZAIREC.HOST_SOKO, vbUnicode);
    
    Printer.Print Tab(31);      '品番（外部）
    Printer.Print StrConv(WK_ZAIREC.HIN_GAI, vbUnicode);
    
    Printer.Print Tab(46);      '品番（内部）
    Printer.Print StrConv(WK_ZAIREC.HIN_NAI, vbUnicode);
    
    Printer.Print Tab(61);      '品名
    Printer.Print StrConv(WK_ZAIREC.HIN_NAME, vbUnicode);
    
    Printer.Print Tab(88);      '棚番（ﾎｽﾄ）
    Printer.Print StrConv(WK_ZAIREC.HOST_TANA, vbUnicode);
    
    Printer.Print Tab(99);      '数量サイン
    Printer.Print StrConv(WK_ZAIREC.QTY_SIGN, vbUnicode);

                                '前日在庫数
    sts = Numeric_Check(EDIT_ONLY, 8, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, StrConv(WK_ZAIREC.ZEN_Z_QTY, vbUnicode), Work)
    If sts = False Then
        Printer.Print Tab(101);
        Printer.Print Work;
    Else
        Printer.Print Tab(103);
        Printer.Print StrConv(WK_ZAIREC.ZEN_Z_QTY, vbUnicode);
    End If
    Printer.Print

    Printer.Print
    Lcnt = Lcnt + 2

End Sub
Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer
Dim Work As String

    Select Case Index
        Case ZERO
            Beep
            yn = MsgBox("在庫初期設定（ＴＯ標準棚）　実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                Call Zaiko_IniSet_Main            '在庫初期設定処理
                Command(ZERO).Enabled = False
                Command(11).SetFocus
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
End Sub
Private Sub Command_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case KeyCode
        Case vbKeyF1
            Command(ZERO).Value = True
        Case vbKeyF12
            Command(11).Value = True
        Case Else
            Beep
            If Command(ZERO).Enabled = True Then
                Command(ZERO).SetFocus
            Else
                Command(11).SetFocus
            End If
    End Select
End Sub
Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = ZERO
    End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim c As String * 128
Dim sts As Integer

Dim sBuffer As String * 255
Dim com     As String

    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
    
    Command_Max = 11            '画面項目別最大ｲﾝﾃﾞｯｸｽ

    Show
    
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                'システム予約済要因取り込み
    If SYSTEM_YOIN_Set() Then
        Beep
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                '有効ﾎｽﾄ倉庫取り込み（炊飯機器）
    For i = 0 To UBound(SOKO_T4) - 1
        SOKO_T4(i).HS_SOKO = "  "
        SOKO_T4(i).NAIGAI = " "
    Next i
    i = ZERO
    Do
        If GetIni("INIZAI_OK_SOKO", "SOKO4" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call Log_Out(LOG_F, "[SYS.INI] [INIZAI_OK_SOKO] [SOKO] READ ERROR")
            End
        End If
        If RTrim(c) = "**" Then
            Exit Do
        End If
        SOKO_T4(i).HS_SOKO = RTrim(c)
        If GetIni("INIZAI_OK_SOKO", "NAIG4" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call Log_Out(LOG_F, "[SYS.INI] [INIZAI_OK_SOKO] [NAIG] READ ERROR")
            End
        End If
        SOKO_T4(i).NAIGAI = RTrim(c)
        i = i + 1
    Loop

                                '有効ﾎｽﾄ倉庫取り込み（電化調理機器）
    For i = ZERO To UBound(SOKO_TD) - 1
        SOKO_TD(i).HS_SOKO = "  "
        SOKO_TD(i).NAIGAI = " "
    Next i
    i = ZERO
    Do
        If GetIni("INIZAI_OK_SOKO", "SOKOD" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call Log_Out(LOG_F, "[SYS.INI] [INIZAI_OK_SOKO] [SOKO] READ ERROR")
            End
        End If
        If RTrim(c) = "**" Then
            Exit Do
        End If
        SOKO_TD(i).HS_SOKO = RTrim(c)
        If GetIni("INIZAI_OK_SOKO", "NAIGD" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call Log_Out(LOG_F, "[SYS.INI] [INIZAI_OK_SOKO] [NAIG] READ ERROR")
            End
        End If
        SOKO_TD(i).NAIGAI = RTrim(c)
        i = i + 1
    Loop

    If Kaso_Soko_No_Set() Then
        Beep
        MsgBox "仮想倉庫の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタ（更新用ワーク）ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '画面初期設定
    Call Scr_Init

End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '品目マスタ（更新用ワーク）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010451 = Nothing

    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    F1010451.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010451)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010451)


    F1010451.MousePointer = vbDefault

End Sub

