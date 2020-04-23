VERSION 5.00
Begin VB.Form FHINCNV1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "品目マスタコンバート 2015.07.02 08:00"
   ClientHeight    =   6315
   ClientLeft      =   1920
   ClientTop       =   2430
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
   ScaleHeight     =   6315
   ScaleWidth      =   11220
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   11.25
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
         Size            =   9.75
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
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "品目マスタコンバート処理が終了しました。"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   16
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "品目マスタコンバート処理実行中です。"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   15
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "件"
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   14
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "品目作成件数"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   12
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "FHINCNV1"
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
Private Type YUKO_SOKO_TBL7             '有効ﾎｽﾄ倉庫取り込みテーブル（掃除機事業部）
    HS_SOKO As String * 2
    NAIGAI As String * 1
End Type
Dim SOKO_T7(0 To 9) As YUKO_SOKO_TBL7


Dim HS_NaiG As String                   '国内外（決定内容）････　ﾎｽﾄﾃﾞｰﾀ内容により設定

Dim Command_Max As Integer              '画面項目別最大ｲﾝﾃﾞｯｸｽ

Dim OK_QTY As Long                      '処理データ件数（掃除機）
                                    '画面初期表示（処理済「便」表示など）
Private Sub Scr_Init()

    Label2(0).Visible = False
    Label2(1).Visible = False
    
    Command(0).SetFocus

End Sub
                                        '品目マスタ登録 処理
Private Sub SYS_HIN_Main()
Dim sts As Integer
Dim Work As String

    If SYS_HIN_Open(1, Work) Then       '取込みﾜｰｸ OPEN（更新ﾓｰﾄﾞ)
        Label2(1).Visible = True        '終了表示
        Label1(1).Caption = "0"
        Exit Sub
    End If
    
    Call Scr_Lock                       '画面項目ロック
    Label2(0).Visible = True            '取込み中ﾒｯｾｰｼﾞ表示


    Do
        DoEvents
        
        If SYS_HIN_Get Then             '取込みﾜｰｸ 読込み
            Exit Do
        End If
        If StrConv(SYS_HINREC.No, vbUnicode) < " " Then      ' EOF ?
            Exit Do
        End If

        Call Upd_Item                   '品目マスタ 登録
        OK_QTY = OK_QTY + 1
        Label1(1).Caption = Format(OK_QTY, "#####0")
    Loop

    Close #SYS_HIN_No                   '取込みﾜｰｸ CLOSE
    Label2(0).Visible = False           '実行中ﾒｯｾｰｼﾞ ｸﾘｱ
    Label2(1).Visible = True            '終了表示

    Label1(1).Caption = Format(OK_QTY, "#####0")

    Call Scr_UnLock
End Sub
                                            '品目マスタ更新
Private Sub Upd_Item()
Dim sts As Integer
Dim ans As Integer
Dim com As Integer
Dim Work As String

    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(SYS_HINREC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(SYS_HINREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(SYS_HINREC.HIN_GAI, vbUnicode))
    
Call LOG_OUT(LOG_F, StrConv(SYS_HINREC.HIN_GAI, vbUnicode) & " " & StrConv(SYS_HINREC.HIN_NAME, vbUnicode))
    
    Do
        sts = BTRV(BtOpGetEqual + 200, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE
                Sleep (500)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Unload Me
        End Select
    Loop
                
                
    If com = BtOpInsert Then
        Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(SYS_HINREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(SYS_HINREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(SYS_HINREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(SYS_HINREC.HIN_NAME, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(SYS_HINREC.HIN_NAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.BIKOU, "")
        Call UniCode_Conv(ITEMREC.IRI_QTY, "00000000")
        Call UniCode_Conv(ITEMREC.FILLER, "")
    End If
'    If StrConv(ITEMREC.HIN_NAI, vbUnicode) <> StrConv(SYS_HINREC.HIN_NAI, vbUnicode) Then  2015.07.01
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(SYS_HINREC.HIN_NAME, vbUnicode))
'    End If                                                                                 2015.07.01
    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(SYS_HINREC.HIN_NAI, vbUnicode))
    Call UniCode_Conv(ITEMREC.ST_SET_DT, StrConv(SYS_HINREC.ST_SET_DT, vbUnicode))
    Call UniCode_Conv(ITEMREC.ST_SOKO, StrConv(SYS_HINREC.ST_SOKO, vbUnicode))
    Call UniCode_Conv(ITEMREC.ST_RETU, StrConv(SYS_HINREC.ST_RETU, vbUnicode))
    Call UniCode_Conv(ITEMREC.ST_REN, StrConv(SYS_HINREC.ST_REN, vbUnicode))
    Call UniCode_Conv(ITEMREC.ST_DAN, StrConv(SYS_HINREC.ST_DAN, vbUnicode))
    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(SYS_HINREC.BIKOU_SOKO, vbUnicode))
    Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(SYS_HINREC.BIKOU_TANA, vbUnicode))
        

    sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrEOF And sts <> BtErrKeyNotFound Then
            Call File_Error(sts, com, "品目マスタ")
            Unload Me
        End If
    End If
End Sub
                                    '画面項目ロック（イベント取得不可）
Private Sub Scr_Lock()

Dim i As Integer

    FHINCNV1.MousePointer = vbHourglass

    For i = 0 To Command_Max
        Command(i).Enabled = False
    Next i

End Sub
                                    '画面項目ロック解除（イベント取得可）
Private Sub Scr_UnLock()

Dim i As Integer

    For i = 0 To Command_Max
        Command(i).Enabled = True
    Next i

    FHINCNV1.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)
Dim yn As Integer

    Select Case Index
        Case 0
            yn = MsgBox("品目マスタコンバート処理　実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                Call SYS_HIN_Main            '品目マスタ設定処理
                Command(0).Enabled = False
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
            Command(0).Value = True
        Case vbKeyF12
            Command(11).Value = True
        Case Else
            Beep
            If Command(0).Enabled = True Then
                Command(0).SetFocus
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
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim j As Integer
Dim c As String * 128
Dim sts As Integer

    Command_Max = 11            '画面項目別最大ｲﾝﾃﾞｯｸｽ

    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        Unload Me
    End If

                                '有効ﾎｽﾄ倉庫取り込み（掃除機）
    For i = 0 To UBound(SOKO_T7) - 1
        SOKO_T7(i).HS_SOKO = "  "
        SOKO_T7(i).NAIGAI = " "
    Next i
    i = 0
    Do
        If GetIni("INIZAI_OK_SOKO", "SOKO7" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call LOG_OUT(LOG_F, "[SYS.INI] [INIZAI_OK_SOKO] [SOKO] READ ERROR")
            End
        End If
        If RTrim(c) = "**" Then
            Exit Do
        End If
        SOKO_T7(i).HS_SOKO = RTrim(c)
        If GetIni("INIZAI_OK_SOKO", "NAIG7" & RTrim(Format$(i + 1, "#0")), "SYS", c) Then
            Beep
            MsgBox "システム異常が発生しました。処理を中止して下さい。"
            Call LOG_OUT(LOG_F, "[SYS.INI] [INIZAI_OK_SOKO] [NAIG] READ ERROR")
            End
        End If
        SOKO_T7(i).NAIGAI = RTrim(c)
        i = i + 1
    Loop

                                '品目マスタＯＰＥＮ
    If ITEM_Open(0) Then
        Beep
        MsgBox "システム異常が発生しました。処理を中止して下さい。"
        End
    End If
                                '画面初期設定
    Call Scr_Init

End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
            Beep
            MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
        End If
    End If
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "品目マスタ")
        Beep
        MsgBox "システム異常が発生しました。処理を終了して下さい。", vbOKOnly
    End If

    End
End Sub

