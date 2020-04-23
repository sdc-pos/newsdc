VERSION 5.00
Object = "{D4A17F03-6EDB-11D2-A6E0-0040262B3978}#2.0#0"; "CtrsWsk.ocx"
Begin VB.Form F1100101 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "スキャナ制御「停止中」"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6255
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "一時停止"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   10695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   10695
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "業務終了"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "業務開始"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin CTRSWSKLib.CtrsWsk CtrsWsk1 
      Left            =   240
      Top             =   360
      _Version        =   131072
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "F1100101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LocalPort               As Long                 'データ受信ポート番号
Private RemotePort              As Long                 'データ送信ポート番号

Private Normal_End              As Boolean              '終了状態

Private Next_Step               As Integer              '次実行指示

Private Menu_Type               As Integer              '1:個別 2:共通


Private Type NAIGAI_tag
    CODE                        As String * 1
    NAME                        As String
End Type

Private NAIGAI()                As NAIGAI_tag

Private Const M_Gyo% = 5                                '最大画面行数
Private Const M_Keta% = 20                              '最大画面桁数

Private Type Recv_Text_Tag                              '受信テキスト
    ID                          As Integer              'IDNO
    LCD(0 To M_Gyo - 1)         As String * 20          '受信内容１～５行目
    Time                        As String * 8           '送信時刻
    RETRY                       As Integer              '送信リトライ2004.04.10
End Type

Private Recv_text               As Recv_Text_Tag

Private Const Start_Para$ = "START"     '子機電源ＯＮ
Private Const Can_Para$ = "CANCEL"      'キャンセル要求
Private Const Fin_Para$ = "FINISHI"     '作業終了要求
Private Const Ent_Para$ = "ENT"         'ENTのみ

Private Const Qty_OK_Para$ = "OK"       '数量ＯＫ
Private Const Loc_OK_Para$ = "T"        '棚番ＯＫ
Private Const DEN_OK_Para$ = "ALL"      '伝票ＯＫ
Private Const LAST_Para$ = "LAST"       '最終作業


Private Type Box_Type_tag               'ＢＯＸ指定
    Box_Type                    As String * 1           'BOX属性
    LCD(0 To M_Keta - 1)        As Byte                 '表示内容
    INIT                        As String * 10          '初期内容
    Start_Pos                   As String * 2           '開始カーソル位置
    Max_Size                    As String * 2           '入力桁数（最大）
    MENU                        As String * 9           'メニュー内容
End Type

Private Type Send_Text_tag              '送信用バッファー
    sts                         As String * 1           'ステータス 1:OK 2:NG
    Display_Flg                 As String * 1           '表示画面フラグ 1:通常入力画面 2:メニュー画面 3:参照画面
    End_Menu                    As String * 1           '最終メニューフラグ 1:次画面あり　2:最終画面
    Menu_Suu                    As String * 2           'メニュー個数
    fileName                    As String * 12          'ファイル名(*.*)
    Buzzer                      As String * 1           'ブザー指定
    Box_Type(0 To M_Gyo - 1)    As Box_Type_tag
    CRLF                        As String * 2
End Type

Private Send_Text           As Send_Text_tag
Private Const Sts_OK$ = "1"             'ステータス　OK
Private Const Sts_NG$ = "2"             'ステータス　NG

Private Const Display_DEF$ = "1"        '表示画面フラグ　通常入力画面
Private Const Display_MENU$ = "2"       '表示画面フラグ　メニュー画面
Private Const Display_REF$ = "3"        '表示画面フラグ　参照画面

Private Const Menu_Head$ = "1"          '最終メニューフラグ 先頭(前あり／後なし)
Private Const Menu_Mid$ = "2"           '最終メニューフラグ 中間(前あり／後あり)
Private Const Menu_End$ = "3"           '最終メニューフラグ 最終(前あり／後なし)
Private Const Menu_Only$ = "4"          '最終メニューフラグ 単独(前なし／後なし)

Private Const Buzzer_DEF$ = "1"         '標準の音
Private Const Buzzer_CONTI$ = "4"       '連続音

Private Const TYPE_REF$ = "X"           '表示ＢＯＸ
Private Const TYPE_BCANK$ = 1           'バーコード＆英数字
Private Const TYPE_BCNUM$ = 2           'バーコード＆数字
Private Const TYPE_BCONLY$ = 3          'バーコードのみ
Private Const TYPE_MENU$ = "M"          'メニュー表示


Private Type Sagyo_Code_tag
    
    CODE_TYPE       As String * 1       '主バーコードタイプ
    YOIN_CODE       As String * 1       '要因
    PARAM           As String * 2       'パラメータ

End Type

Private Type ID_KANRI_TBL_tag               '子機情報の管理
    RETRY                       As Integer              '子機送信リトライ
    ID                          As Integer              'IDNO
    Step                        As Integer              '進捗フラグ
    JGYOBU                      As String * 1           '事業部
    NAIGAI                      As String * 1           '国内外
    Hinban                      As String * 13          '品番（２レスポンス以上の作業時のみ使用）
    Tanaban                     As String * 8           '棚番（２レスポンス以上の作業時のみ使用）
    GOODS_ON_F                  As String * 1           '商品化用倉庫
    
    '---------------------------------------------------'送信数量
    Send_SUMI_QTY               As Long                 '商品化済み数量（移動時）
    Send_MI_QTY                 As Long                 '未商品数量（移動時）
    Send_Syuka_QTY              As Long                 '出荷数量（出荷時）
    '---------------------------------------------------'出荷処理用▽
    MTS_CODE                    As String * 8           '得意先コード
    SS_CODE                     As String * 8           '直送先コード
    CYU_KBN                     As String * 1           '注文区分
    Y_SYU_CNT                   As Integer              '対象伝票枚数
    ID_NO                       As String * 12          '伝票ID
    DEN_NO                      As String * 6           '伝票番号
    YUKO_SUMI_QTY               As Long                 '使用可能な商品化済み在庫
    YUKO_MI_QTY                 As Long                 '使用可能な未商品在庫
    SYUKA_QTY                   As Long                 '出荷数量（全数）
    
    '---------------------------------------------------'出荷処理用△
'    MENU_GRP                    As String * 2           '使用メニュー
    
    
    MENU_LV1                    As String * 2           'メニューレベル１   2006.01.30 3-->2
    MENU_LV2                    As String * 2           'メニューレベル２
'    MENU_LV3                    As String * 3           'メニューレベル３  2006.01.30
    
    SAGYO_LOG                   As String * 1           '作業ﾛｸﾞ出力 0:なし　1:あり 2006.01.30
    
    
    PageNo_LV1                  As Integer              'ページ№（メニュー）
    PageNo_LV2                  As Integer              'ページ№（メニュー）
'    PageNo_LV3                  As Integer              'ページ№（メニュー）  2006.01.30
    Sagyo_Code                  As Sagyo_Code_tag       '作業コード
    YOIN_DNAME                  As String * 5           '表示用名称
    TANTO_CODE                  As String * 5           '担当者コード
    Recv_text(0 To M_Gyo - 1)   As String * 20          '最終受信内容１～５行目
    Send_Text                   As Send_Text_tag        '最終送信内容(正常値)
    Last_Send_Text              As Send_Text_tag        '最終送信内容(全て)
    Time                        As String * 8           '送信時刻
    Last_Send                   As Integer              '0:通常テキスト 1:エラー情報


    S_JGYOBU                    As String * 1           '資材を踏まえての事業部
    S_NAIGAI                    As String * 1           '資材を踏まえての国内外



End Type

Private ID_KANRI_TBL()      As ID_KANRI_TBL_tag

Private ING_No              As Integer  '処理中の添字

Private Const Step_Start% = 0           '子機電源ＯＮ
Private Const Step_TANTO_REQ% = 1       '担当者要求
Private Const Step_TANTO_RES% = 2       '担当者回答

Private Const Step_JGYOBU_REQ% = 3      '事業部要求
Private Const Step_JGYOBU_RES% = 4      '事業部回答

Private Const Step_NAIGAI_REQ% = 5      '国内外要求
Private Const Step_NAIGAI_RES% = 6      '国内外回答


Private Const Step_MENU1_REQ% = 10      'メニュー１要求
Private Const Step_MENU1_RES% = 11      'メニュー１回答
Private Const Step_MENU2_REQ% = 12      'メニュー２要求
Private Const Step_MENU2_RES% = 13      'メニュー２回答
'2006.01.30 Private Const Step_MENU3_REQ% = 14      'メニュー３要求
'2006.01.30 Private Const Step_MENU3_RES% = 15      'メニュー３回答

Private Const Step_Sagyo1_REQ% = 20     '作業１要求
Private Const Step_Sagyo1_RES% = 21     '作業１回答
Private Const Step_Sagyo2_REQ% = 22     '作業２要求
Private Const Step_Sagyo2_RES% = 23     '作業２回答
Private Const Step_Sagyo3_REQ% = 24     '作業３要求
Private Const Step_Sagyo3_RES% = 25     '作業３回答
Private Const Step_Sagyo4_REQ% = 26     '作業４要求
Private Const Step_Sagyo4_RES% = 27     '作業４回答



Private Const BEF_Page$ = "$B"          '前頁
Private Const NEXT_Page$ = "$N"         '次頁


Private Type Menu_Tbl_tag               'メニュー送信用テーブル
    MENU_NO     As String * 2
    PARAM       As String * 16
    Disp        As String
    Log_Out     As String * 1

    
End Type



Private Type Wel_Para_Tag               'WELCAT送信用テーブル
    Box_Type    As String * 1
    LCD         As String * 10
    Keta        As Integer
End Type

Private Type WEL_Para_Tbl_tag
    Action      As String * 2
    Wel_Para(0 To M_Gyo - 1) As _
                Wel_Para_Tag
End Type
                                        '※作業宣言の数により増減必須！！
Private WEL_Para_Tbl(0 To 14, 0 To 9) As WEL_Para_Tbl_tag

Private Const LCD_Tanaban$ = "棚番"
Private Const LCD_Hinban$ = "品番"
Private Const LCD_Suryo$ = "数量"
Private Const LCD_Syuka$ = "出荷残"
Private Const LCD_SUMI_Suryo$ = "商品"
Private Const LCD_MI_Suryo$ = "未商品"
Private Const LCD_ID_No$ = "伝票ID"
Private Const LCD_SYUKO_HYO_No$ = "出庫表№"
Private Const LCD_MTS$ = "向け先"


Private Const LCD_To_Tanaban$ = "移動先棚番"



Private FILE_RETRY  As Integer          'ファイル使用中時のリトライ回数

Private Const Wel_TANAOROSI$ = "B0"         '「WEL 棚卸し」の要因
Private Const Wel_TANAHYOJI$ = "B1"         '「WEL 棚番表示」の要因
Private Const Wel_HIN_SHOGO$ = "B2"         '「WEL 品番別照合」の要因
Private Const Wel_AVE_SYUKA$ = "B3"         '「WEL 月平均出荷数」の要因
Private Const Wel_HOST_ZAIKO$ = "B4"        '「WEL ホスト在庫照会」の要因
Private Const Wel_ST_TANABAN$ = "B5"        '「WEL 標準棚番設定」の要因
Private Const Wel_RIREKI$ = "B6"            '「WEL 当日出庫履歴」の要因
Private Const Wel_SUII$ = "B7"              '「WEL 出荷推移」の要因
Private Const Wel_TANA_HIN_SHOGO$ = "B8"    '「WEL 棚番・品番別照合」の要因

Private Const Wel_TANAHYOJI_KASO$ = "B9"    '「WEL 棚番表示(仮想優先)」の要因


Private Const Wel_GOODS_ONOFF_ONO$ = "D0"   '「WEL 商品/未商品切り替え　小野」の要因
Private Const Wel_GOODS_ONOFF_SIGA$ = "D1"  '「WEL 商品/未商品切り替え　滋賀」の要因


Private Const Wel_RETURNED_GOODS$ = "E0"    '「WEL 良品返品」の要因
Private Const Wel_LOCATION_MOVE$ = "E1"     '「WEL 棚移動」の要因



Private B1_SendFile As String               '「WEL 棚番表示」の送信ファイル名
Private B6_SendFile As String               '「WEL 当日出庫履歴」の送信ファイル名
Private B7_SendFile As String               '「WEL 出荷推移」の送信ファイル名

Private B9_SendFile As String               '「WEL 棚番表示(仮想優先)」の送信ファイル名



Private Type SendFileRec_Tag                '送信ファイルレコード定義
    Title           As String * 1           'タイトル
    LCD(0 To 19)    As Byte                 '表示メッセージ
    CRLF            As String * 2           'CR/LF
End Type


Private Const Wel_Kbn_Title$ = "0"      'タイトル行
Private Const Wel_Kbn_Normal$ = "1"     '通常表示行


Private Type Tanahyoji_tag              '棚表示用集計テーブル
    Tanaban         As String * 8
    SUMI_QTY        As Long
    MI_QTY          As Long
End Type

Private Inspection_Flg      As Integer  '検品チェックフラグ(0:出庫未完NG 1:出庫未完でもOK)
Private B2_MEMO     As String           '品番別在庫照合（要因＝B2）のメモ
Private B8_MEMO     As String           '品番別棚別在庫照合（要因＝B8）のメモ

Private ALL_MENU_GRP    As String * 2





Private Sub Command1_Click(Index As Integer)
'-------------------------------------------------------
'
'   『業務開始指示』
'       １． ポートの獲得
'       ２． ポートの開放
'-------------------------------------------------------
    
Dim ans As Integer
    
    On Error GoTo Error
    
    Select Case Index
        
        Case 0                              '業務開始
            
            CtrsWsk1.Bind LocalPort, RemotePort
            F1100101.Caption = "スキャナ制御「実行中」"
    
            Command1(0).Enabled = False
            Command1(1).Enabled = True
            Command1(2).Enabled = True
    
        Case 1                              '業務終了
            
            
            ans = MsgBox("本日の業務終了しますか？", vbYesNo + vbDefaultButton2, "業務終了")
            If ans = vbNo Then
                Exit Sub
            End If
            
            CtrsWsk1.Unbind
            
            Normal_End = False              '正常終了
            
            Next_Step = 1                   '次処理起動する
            Unload Me
    
        Case 2
            CtrsWsk1.Unbind
            
            Normal_End = False              '正常終了
            Next_Step = 0                   '次処理起動しない
            Unload Me
    
    End Select
    
    Exit Sub

Error:
    MsgBox "Winsock Error= " & Err.Description    'ステータス行にエラーを表示します。
    
    Call Log_Out(LOG_F, "Winsock Error= " & Err.Description)
    
    Normal_End = True                       '異常終了
    Unload Me
    


End Sub

Private Sub CtrsWsk1_OnReceive(ByVal ID_NO As Integer, ByVal RecvText As String, ByVal Resp_Mode As Boolean)
'-------------------------------------------------------
'
'   『レコード受信時処理』
'
'-------------------------------------------------------

Dim nErrCode    As Integer
Dim strErrMsg   As String
Dim intLine     As Integer
Dim i           As Integer
Dim j           As Integer
Dim Chk_Time    As String * 8
Dim Sendbuf     As String

Dim Errbuf      As String

Dim sts         As Integer

Dim Start_Flg   As Integer


    Text1(0).Text = Format(ID_NO, "000") & ", Recv=" & RecvText
    
'    Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Recv=" & RecvText)
        
    RecvText = Left(RecvText, Len(RecvText) - 2)
    
                                    'ＩＤ№で受信済みテキスト検索
    ING_No = -1
    
    Start_Flg = False
    
    For i = 0 To UBound(ID_KANRI_TBL)
        If ID_NO = ID_KANRI_TBL(i).ID Then
            ING_No = i
            Chk_Time = ID_KANRI_TBL(i).Time
            Exit For
        End If
    Next i
    
    
    
    
    
    If i > UBound(ID_KANRI_TBL) Then
                                                'ＩＤ№新規登録
        For i = 0 To UBound(ID_KANRI_TBL)
            If ID_KANRI_TBL(i).ID = 0 Then
                
                Start_Flg = True
                
                ID_KANRI_TBL(i).ID = ID_NO      'ID_No  保存
                
'                ID_KANRI_TBL(i).MENU_GRP = ""
                ID_KANRI_TBL(i).MENU_LV1 = ""
                ID_KANRI_TBL(i).MENU_LV2 = ""
''                ID_KANRI_TBL(i).MENU_LV3 = ""
                
                If UBound(JGYOBU_T) = 0 Then    '１事業部固定
                Else
                    ID_KANRI_TBL(i).JGYOBU = ""
                End If
                
                If UBound(NAIGAI) = 0 Then   '国内外固定
                Else
                    ID_KANRI_TBL(i).NAIGAI = ""
                End If
                
                ING_No = i
                Chk_Time = ""
                Exit For
            End If
        
        Next i
    End If
    
    
'Call Log_Out(LOG_F, Format(ID_NO, "000") & ",Yoin= " & ID_KANRI_TBL(i).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(i).Sagyo_Code.YOIN_CODE)
    
    
    If ING_No = -1 Then
        MsgBox "ＩＮＩファイルの子機数の設定を変更して下さい。", vbCritical
        Normal_End = True
        Unload Me
    End If
    
                                            '前回受信値を再受信した？
    If Left(Right(RecvText, 9), 8) = ID_KANRI_TBL(i).Time And _
         Right(RecvText, 1) = "1" Then
            Call Send_Err_Proc(Sendbuf)
    
            Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf & "[再送信]")
    
    Else
                                            '受信内容を保存
        ID_KANRI_TBL(ING_No).Recv_text(0) = Left(RecvText, 20)       '受信内容１行目
        ID_KANRI_TBL(ING_No).Recv_text(1) = Mid(RecvText, 21, 20)    '受信内容２行目
        ID_KANRI_TBL(ING_No).Recv_text(2) = Mid(RecvText, 41, 20)    '受信内容３行目
        ID_KANRI_TBL(ING_No).Recv_text(3) = Mid(RecvText, 61, 20)    '受信内容４行目
        ID_KANRI_TBL(ING_No).Recv_text(4) = Mid(RecvText, 81, 20)    '受信内容４行目
        ID_KANRI_TBL(ING_No).Time = Right(RecvText, 8)               '送信時刻
        
        
                                            
        If Start_Flg Then
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> Start_Para Then
                Call Err_Send_Proc("再起動してください。", "", "", "", "")
                Sendbuf = Text_Create_Proc()
            End If
        End If
                                            
                                            
                                            '[START][CANCEL][FINISHI]受信は初期化する
        Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
            Case Start_Para     '開始(子機電源ON)
                
                
                            '出荷予定／在庫の予約解除
                sts = Data_Clear_Proc(0, Sendbuf)
                Select Case sts
                    Case SYS_ERR
                        Normal_End = True
                End Select
                
                
                ID_KANRI_TBL(ING_No).Step = Step_Start
        
                Call Start_Proc(Sendbuf)
            
            
            Case Ent_Para       'ENT
                If Not Start_Flg Then
                    
                    If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_MTS Then
                        '検品時の確認
                        
                        
                        
    '                    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    '                    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    '                    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    '                    Select Case sts
    '                        Case BtNoErr
    '                        '   -------------------------------- エラーメッセージ作成
    '                        Case Else
    '                       '重要な要因なので未登録はシステム停止とする
    '                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
    '                        Sendbuf = Text_Create_Proc()
    '                        Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
    '                        Normal_End = True
    '                    End Select
    '
                        
                        
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                        
                        If Sagyo_Main_Proc(Sendbuf) Then
                            Normal_End = True
    '                        Unload Me
                        End If
                    Else
                        
                        '参照画面の確認時のみ
                        
                        
                        Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                        Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                        Select Case sts
                            Case BtNoErr
                            '   -------------------------------- エラーメッセージ作成
                            Case Else
                           '重要な要因なので未登録はシステム停止とする
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                            Normal_End = True
                        End Select
                       
                        
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Normal_End = True
                        End If
                        Sendbuf = Text_Create_Proc()
                    End If
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            Case Can_Para       'CANCEL
                If Not Start_Flg Then
                
                    If ID_KANRI_TBL(ING_No).Last_Send = 1 Then
                                
                                
                        '検品時はデータの開放を行う　2004.06.14 ↓
                        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_MTS Then
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    If Sagyo_Send_Proc() Then
                                        Sendbuf = Text_Create_Proc()
                                        Normal_End = True
                                    End If
                                    Sendbuf = Text_Create_Proc()
                                
                                Case SYS_ERR
                                    Normal_End = True
                            End Select
                        
                        
                        End If
                                
                        '検品時はデータの開放を行う　2004.06.14 ↑
                                
                                '前回がエラー送信
                        Call Re_Send_Proc(Sendbuf)
                
                    Else
                                '出荷予定／在庫の予約解除
                        sts = Data_Clear_Proc(0, Sendbuf)
                        Select Case sts
                            Case SYS_CANCEL
                                If Sagyo_Send_Proc() Then
                                    Sendbuf = Text_Create_Proc()
                                    Normal_End = True
                                End If
                                Sendbuf = Text_Create_Proc()
                            
                            Case SYS_ERR
                                Normal_End = True
                        End Select
                
                
                
                        Call Cancel_Proc(Sendbuf)
                
                    End If
            
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            
            Case Fin_Para       'FINISH
                    
                If Not Start_Flg Then
                
                    '出荷予定／在庫の予約解除
                    sts = Data_Clear_Proc(0, Sendbuf)
                    Select Case sts
                        Case SYS_CANCEL
                            If Sagyo_Send_Proc() Then
                                Normal_End = True
                            End If
                            Sendbuf = Text_Create_Proc()
                    
                        Case SYS_ERR
                            Normal_End = True
                    End Select
                
                
'                    If Step_MENU1_REQ < ID_KANRI_TBL(ING_No).Step Then
                    If Step_TANTO_REQ <> ID_KANRI_TBL(ING_No).Step Then      '2005.01.07 if ～　else ～　end if
                    
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30                        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                
                        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.03                        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
                
                
                
                        ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                
                        If Menu_Send_Proc(Sendbuf) Then
                            Normal_End = True
    '                Unload Me
                        End If
                
                    Else                                                    '2005.01.07
'                        ID_KANRI_TBL(ING_No).Step = Step_Start
                                                                            '2005.01.07
                        Call Start_Proc(Sendbuf)                            '2005.01.07
                                                                            '2005.01.07
                    End If                                                  '2005.01.07
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            Case Else
                If Not Start_Flg Then
                                            '進捗チェック
                    Select Case ID_KANRI_TBL(ING_No).Step
            
            
                        Case Step_TANTO_REQ         '担当者要求に対するレス
                        
                            
                            
                            ID_KANRI_TBL(ING_No).Step = Step_TANTO_RES
        
                            If Normal_Proc(Sendbuf) Then
                                Normal_End = True
    '                        Unload Me
                            End If
        
                        Case Step_JGYOBU_REQ        '事業部要求に対するレス
                
                            Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
    '                        Case BEF_Page       '前頁
    '                        Case NEXT_Page      '次頁
                                Case Else            '事業部受信
                
                                    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_RES
              
                                    ID_KANRI_TBL(ING_No).JGYOBU = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
        
                                    If Normal_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
                                    
                        Case Step_NAIGAI_REQ
                    
                            Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                Case Else           'メニューパラメータ受信
                
                                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_RES
                                        
                                    ID_KANRI_TBL(ING_No).NAIGAI = Trim(ID_KANRI_TBL(i).Recv_text(0))
                                
        
                                    If Normal_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
    
'2006.01.30                        Case Step_MENU1_REQ, Step_MENU2_REQ, Step_MENU3_REQ
                        Case Step_MENU1_REQ, Step_MENU2_REQ
                
                             Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                Case BEF_Page       '前頁
                            
                                    
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV1 = ID_KANRI_TBL(ING_No).PageNo_LV1 - 1
                                        Case Step_MENU2_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = ID_KANRI_TBL(ING_No).PageNo_LV2 - 1
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = ID_KANRI_TBL(ING_No).PageNo_LV3 - 1
                                    End Select
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                    
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                            
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            
                                Case NEXT_Page      '次頁
                                    
                                    
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV1 = ID_KANRI_TBL(ING_No).PageNo_LV1 + 1
                                        Case Step_MENU2_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = ID_KANRI_TBL(ING_No).PageNo_LV2 + 1
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = ID_KANRI_TBL(ING_No).PageNo_LV3 + 1
                                    End Select
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                            
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                
                                Case Else           'メニューパラメータ受信
                
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
                                        Case Step_MENU2_REQ
'                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
                                    End Select
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_RES
                                            ID_KANRI_TBL(ING_No).MENU_LV1 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                        Case Step_MENU2_RES
                                            ID_KANRI_TBL(ING_No).MENU_LV2 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
 
                                            ID_KANRI_TBL(ING_No).MTS_CODE = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 3, 8)
                                            ID_KANRI_TBL(ING_No).SS_CODE = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 11, 8)
                            
                            
                            
                                            

'2006.01.30                                        Case Step_MENU3_RES
'2006.01.30                                            ID_KANRI_TBL(ING_No).MENU_LV3 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                    End Select
                
                                
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
                        
                        Case Step_Sagyo1_REQ, Step_Sagyo2_REQ, Step_Sagyo3_REQ, Step_Sagyo4_REQ
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                        
                            If Sagyo_Main_Proc(Sendbuf) Then
                                Normal_End = True
    '                        Unload Me
                            End If
                        Case Else
                    End Select
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
    
        
        End Select
    
    End If
    
    

    If Resp_Mode Then
        On Error GoTo ShowError

        CtrsWsk1.SendResp Sendbuf

'        Text1(1).Text = Format(ID_NO, "000") & ", Send=" & SendBuf
        
'        Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf)

        On Error GoTo 0
    
        ID_KANRI_TBL(ING_No).Last_Send_Text.sts = Send_Text.sts                     'ステータス
        ID_KANRI_TBL(ING_No).Last_Send_Text.Display_Flg = Send_Text.Display_Flg     '表示画面フラグ
        ID_KANRI_TBL(ING_No).Last_Send_Text.End_Menu = Send_Text.End_Menu           '最終メニューフラグ
        ID_KANRI_TBL(ING_No).Last_Send_Text.Menu_Suu = Send_Text.Menu_Suu           'メニュー個数
        ID_KANRI_TBL(ING_No).Last_Send_Text.fileName = Send_Text.fileName           'ファイル名
        ID_KANRI_TBL(ING_No).Last_Send_Text.Buzzer = Send_Text.Buzzer               'ブザー指定
        
        For j = 0 To M_Gyo - 1
                                                                                    'BOX属性
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Box_Type = Send_Text.Box_Type(j).Box_Type
                                                                                    '表示内容
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).LCD, StrConv(Send_Text.Box_Type(j).LCD, vbUnicode))
                                                                                    '初期内容
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).INIT = Send_Text.Box_Type(j).INIT
                                                                                    '開始カーソル位置
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Start_Pos = Send_Text.Box_Type(j).Start_Pos
                                                                                    '入力桁数（最大）
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Max_Size = Send_Text.Box_Type(j).Max_Size
                                                                                    'メニュー内容
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).MENU = Send_Text.Box_Type(j).MENU
                    
        
        Next j
    
    
        If Normal_End Then
            
'            MsgBox "システム異常が発生しました！！処理をしてください。"
 
            
'            Unload Me
        End If
    End If

    Exit Sub

ShowError:
    nErrCode = Err.Number
    strErrMsg = Err.Description         'エラーメッセージ
    
    intLine = CtrsWsk1.ErrLineNo        '接続番号を取得します。
    If intLine > 0 Then
        strErrMsg = strErrMsg & Chr(&HD) & Chr(&HA) & "接続番号 = " & intLine
    End If

    Text1(2).Text = strErrMsg           'ステータス行にエラーを表示します。
    
'    Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf)


End Sub

Private Sub Form_Load()
    
Dim c           As String * 128
Dim Out_Data    As String

Dim Box_Type    As String * 1
Dim LCD         As String * 10
Dim Keta        As String * 2

Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
    
Dim sts         As Integer
    
    Normal_End = False
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
'---------------------------------------------- 'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)

'---------------------------------------------- 'データ受信ポート番号取り込み
    If GetIni(App.EXEName, "LocalPort", "SYS", c) Then
        Beep
        MsgBox "データ受信ポート番号の獲得に失敗しました。処理を中止します。"
        End
    End If
    LocalPort = CLng(RTrim(c))

'---------------------------------------------- 'データ送信ポート番号取り込み
    If GetIni(App.EXEName, "RemotePort", "SYS", c) Then
        Beep
        MsgBox "データ送信ポート番号の獲得に失敗しました。処理を中止します。"
        End
    End If
    RemotePort = CLng(RTrim(c))
'---------------------------------------------- '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止します。"
        End
    End If
'---------------------------------------------- '国内外情報取り込み
    i = 0
    
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI_CODE" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI(i - 1)
        NAIGAI(i - 1).CODE = Trim(c)
        If GetIni(App.EXEName, "NAIGAI_NAME" & Format(i, "0"), "SYS", c) Then
            MsgBox "国内外の獲得に失敗しました。処理を中止します。"
            End
        End If
        NAIGAI(i - 1).NAME = Trim(c)
    
    Loop
    
    If i = 1 Then
        Beep
        MsgBox "国内外情報の獲得に失敗しました。処理を中止します。"
        End
    End If
'---------------------------------------------- 前借り入荷情報獲得
    If GetIni("YOIN", "YOIN_MAEGARI", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAEGARI] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_MAEGARI = Trim(c)
'---------------------------------------------- '国内外振替情報獲得
    If GetIni("YOIN", "YOIN_FURIKAE", "SYS", c) Then
    Else
        YOIN_FURIKAE = RTrim(c)
        '国内外振替設定時、以下の項目必須
        If GetIni("YOIN", "YOIN_FURIKAE_OUT", "SYS", c) Then
            Beep
            MsgBox "国内外振替情報[YOIN_FURIKAE_OUT]の獲得に失敗しました。処理を中止します。"
            End
        End If
    
        YOIN_FURIKAE_OUT = RTrim(c)
    
        If GetIni("YOIN", "YOIN_FURIKAE_IN", "SYS", c) Then
            Beep
            MsgBox "国内外振替情報[YOIN_FURIKAE_IN]の獲得に失敗しました。処理を中止します。"
            End
        End If
    
        YOIN_FURIKAE_IN = RTrim(c)
    
    End If
'---------------------------------------------- 棚照合情報獲得
    If GetIni("YOIN", "YOIN_WEL_TANASHOGO", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANASHOGO] READ ERROR")
        MsgBox "システム予約済要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_TANASHOGO = Trim(c)

'---------------------------------------------- 棚品照合情報獲得
    If GetIni("YOIN", "YOIN_WEL_TANAHINSHOGO", "SYS", c) Then
        YOIN_TANAHINSHOGO = Wel_TANA_HIN_SHOGO
    Else
        YOIN_TANAHINSHOGO = Trim(c)
    End If

'---------------------------------------------- '子機台数取り込み
    If GetIni(App.EXEName, "KO_SU", "SYS", c) Then
        Beep
        MsgBox "子機台数の獲得に失敗しました。処理を中止します。"
        End
    End If
    ReDim ID_KANRI_TBL(0 To CInt(RTrim(c)) - 1)

    For i = 0 To UBound(ID_KANRI_TBL)
        ID_KANRI_TBL(i).ID = 0          'IDNoクリアー
        ID_KANRI_TBL(i).Step = 0        '進捗クリアー
    
    Next i
'---------------------------------------------- '送信用パラメータ取り込み
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
            WEL_Para_Tbl(i, j).Action = ""
        Next j
    Next i
    
    i = 0
    Do
        i = i + 1
        
        If GetIni("ACTION", "ACTION_CD" & Format(i, "00"), "SYS", c) Then
            Beep
            MsgBox "WELCAT送信用パラーメータ([ACTION] [ACTION_CD])の獲得に失敗しました。処理を中止します。"
            End
        End If
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
    
        j = 0
    
        Do
            j = j + 1
            If GetIni("ACTION", "ACTION_WEL_PARA" & Format(i, "00") & Format(j, "00"), "SYS", c) Then
                Beep
                MsgBox "WELCAT送信用パラーメータ([ACTION] [ACTION_WEL_PARA])の獲得に失敗しました。処理を中止します。"
                End
            End If
            If Trim(c) = "NON" Then
                Exit Do
            End If
        
            Call Data_Select(Trim(c), 1, 14, Out_Data)
            
            WEL_Para_Tbl(i - 1, j - 1).Action = Trim(Out_Data)
        
            Call Data_Select(Trim(c), 2, 14, Out_Data)
            
            WEL_Para_Tbl(i - 1, j - 1).Wel_Para(0).Box_Type = Trim(Out_Data)
            WEL_Para_Tbl(i - 1, j - 1).Wel_Para(0).LCD = ""
        
        
            k = 2
            Do
                
                k = k + 1
                
                If k > 14 Then
                    Exit Do
                End If
                
                Call Data_Select(Trim(c), k, 14, Out_Data)
                Box_Type = Trim(Out_Data)
                
                k = k + 1
                Call Data_Select(Trim(c), k, 14, Out_Data)
                LCD = Trim(Out_Data)
            
                k = k + 1
                Call Data_Select(Trim(c), k, 14, Out_Data)
                Keta = Trim(Out_Data)
            
                Select Case k
                    Case 5
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).Keta = CInt(Keta)
                    Case 8
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).Keta = CInt(Keta)
                    Case 11
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).Keta = CInt(Keta)
                    Case 14
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).Keta = CInt(Keta)
                
                End Select
            
            Loop
            
        
        Loop
    Loop
'---------------------------------------------- '対WELCAT　送受信ログファイル取り込み
    
    If GetIni(App.EXEName, "LOG_F", "SYS", c) Then
        CtrsWsk1.LogFile = ""
    Else
        CtrsWsk1.LogFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　データ送信用フォルダ取り込み
    If GetIni(App.EXEName, "SendFolder", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用フォルダ([F110010] [SendFolder])の獲得に失敗しました。処理を中止します。"
        End
    Else
        CtrsWsk1.SendFolder = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　棚番表示用データファイル名取り込み
    If GetIni(App.EXEName, "B1", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B1])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B1_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　出庫履歴用データファイル名取り込み
    If GetIni(App.EXEName, "B6", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B6])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B6_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　出荷推移用データファイル名取り込み
    If GetIni(App.EXEName, "B7", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B7])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B7_SendFile = Trim(c)
    End If
'---------------------------------------------- '対WELCAT　棚番表示(仮想優先)用データファイル名取り込み
    If GetIni(App.EXEName, "B9", "SYS", c) Then
        Beep
        MsgBox "WELCAT送信用ファイル([F110010] [B9])の獲得に失敗しました。処理を中止します。"
        End
    Else
        B9_SendFile = Trim(c)
    End If

'---------------------------------------------- '共通メニュー情報取り込み
    If GetIni(App.EXEName, "ALL_MENU_GRP", "SYS", c) Then
        Beep
        MsgBox "共通メニュー情報の獲得に失敗しました。処理を中止します。"
        End
    End If


    ALL_MENU_GRP = Trim(c)

'---------------------------------------------- '検品チェック
    If GetIni(App.EXEName, "Inspection", "SYS", c) Then
        Inspection_Flg = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            Inspection_Flg = 1
        Else
            Inspection_Flg = CInt(Trim(c))
        End If
    End If
'---------------------------------------------- '在庫照合メモ項目
    If GetIni(App.EXEName, "B2_MEMO", "SYS", c) Then
        B2_MEMO = ""
    Else
        B2_MEMO = Trim(c)
    End If
'--
    If GetIni(App.EXEName, "B8_MEMO", "SYS", c) Then
        B8_MEMO = ""
    Else
        B8_MEMO = Trim(c)
    End If
'---------------------------------------------- 'ファイルリトライ回数取り込み
    If GetIni("SYSTEM", "RETRY", "SYS", c) Then
        FILE_RETRY = 1
    Else
        If Not IsNumeric(Trim(c)) Then
            FILE_RETRY = 1
        Else
            FILE_RETRY = CInt(Trim(c))
        End If
    End If
'---------------------------------------------- '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '棚マスタＯＰＥＮ
    If TANA_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '品目マスタ(ワーク)ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
'---------------------------------------------- '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- 'メニュー管理マスタＯＰＥＮ
    If P_MENU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '担当者別メニューＯＰＥＮ
    If P_TMENU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫データ（移動処理用）ＯＰＥＮ
    If wZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫データ（商品／未商品切り替え用）ＯＰＥＮ
    If tmpZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '前借データＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '出荷予定データＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫移動歴データＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '在庫集計データＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '資材前借ﾃﾞｰﾀＯＰＥＮ
    If P_NYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '作業実績ﾛｸﾞＯＰＥＮ
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- 'メニュー機能チェック（個別 or 共通）
'    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
'    sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
'    Select Case sts
'        Case BtNoErr
'            Menu_Type = 1           '共通メニューで運用
'        Case BtErrKeyNotFound
            Menu_Type = 2           '担当者別メニューで運用
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "担当者別メニュー")
'            Unload Me
'    End Select
    
    


    Show

    If Data_Clear_Proc(1, "") Then
        MsgBox "データ初期設定が出来ませんでした。"
        Unload Me
    End If


    If tmpZaiko_Clear_Proc() Then
        MsgBox "データ初期設定が出来ませんでした。"
        Unload Me
    End If
End Sub

Private Sub Start_Proc(Sendbuf As String)
'-------------------------------------------------------
'
'   『子機開始処理』
'
'-------------------------------------------------------
Dim i   As Integer
                                                '送信テキスト作成＆管理テーブルに転送
    ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ                      '担当者要求
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                             '表示画面フラグ 通常入力
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = ""                                         '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = ""
    
    Send_Text.Menu_Suu = "05"                                       'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                         '送信データファイル名
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    '---------------------------------------------------------------
    Send_Text.Box_Type(0).Box_Type = TYPE_REF                       'ボックス属性　表示ＢＯＸ
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, "担当者入力")      '表示メッセージ
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, "担当者入力")
                                                                    '初期表示内容（数値）
    Send_Text.Box_Type(0).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
    
    
    Send_Text.Box_Type(0).Start_Pos = ""                            'カーソル位置
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
    
    Send_Text.Box_Type(0).Max_Size = "00"                           '最大桁数
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
    
    Send_Text.Box_Type(0).MENU = ""                                 'メニュー番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(1).Box_Type = TYPE_BCANK                     'ボックス属性　表示ＢＯＸ
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_BCANK
    
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "")                '表示メッセージ
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "")
                                                                    '初期表示内容（数値）
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
    
    Send_Text.Box_Type(1).Start_Pos = "01"                          'カーソル位置
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
    
    Send_Text.Box_Type(1).Max_Size = "05"                           '最大桁数
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "05"
    
    Send_Text.Box_Type(1).MENU = ""                                 'メニュー番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(2).Box_Type = TYPE_REF                       'ボックス属性　表示ＢＯＸ
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "")                '表示メッセージ
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "")
                                                                    '初期表示内容（数値）
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
    
    Send_Text.Box_Type(2).Start_Pos = ""                            'カーソル位置
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
    
    Send_Text.Box_Type(2).Max_Size = "00"                           '最大桁数
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
    
    Send_Text.Box_Type(2).MENU = ""                                 'メニュー番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(3).Box_Type = TYPE_REF                       'ボックス属性　表示ＢＯＸ
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")                '表示メッセージ
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                    '初期表示内容（数値）
    Send_Text.Box_Type(3).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
    
    Send_Text.Box_Type(3).Start_Pos = ""                            'カーソル位置
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
    
    Send_Text.Box_Type(3).Max_Size = "00"                           '最大桁数
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
    
    Send_Text.Box_Type(3).MENU = ""                                 'メニュー番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(4).Box_Type = TYPE_REF                       'ボックス属性　表示ＢＯＸ
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")                '表示メッセージ
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                    '初期表示内容（数値）
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
    
    Send_Text.Box_Type(4).Start_Pos = ""                            'カーソル位置
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
    
    Send_Text.Box_Type(4).Max_Size = "00"                           '最大桁数
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
    
    Send_Text.Box_Type(4).MENU = ""                                 'メニュー番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    '---------------------------------------------------------------
    Send_Text.CRLF = vbCrLf
    '------------------------------------------ 送信バッファーへ転送
    Sendbuf = Text_Create_Proc()

    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信

End Sub


Private Function Normal_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『通常テキスト受信』
'
'-------------------------------------------------------
    Normal_Proc = True
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Start                 '子機電源ON（ここには来ない）
        
        Case Step_TANTO_REQ             '担当者要求（ここには来ない）
        
        Case Step_TANTO_RES             '担当者回答
    
            If Tanto_Check_Proc(Sendbuf) Then
                Exit Function
            End If
    
        Case Step_JGYOBU_REQ            '事業部要求（ここには来ない）
                
        Case Step_JGYOBU_RES            '事業部回答
            
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
                
        Case Step_NAIGAI_REQ            '国内外要求（ここには来ない）
                
        Case Step_NAIGAI_RES            '国内外回答
            
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
                                        'メニュー要求（ここには来ない）
'2006.01.30        Case Step_MENU1_REQ, Step_MENU2_REQ, Step_MENU3_REQ
        Case Step_MENU1_REQ, Step_MENU2_REQ
                                        'メニュー回答
'2006.01.30        Case Step_MENU1_RES, Step_MENU2_RES, Step_MENU3_RES
        Case Step_MENU1_RES, Step_MENU2_RES
    
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
    
    End Select
    
    Normal_Proc = False

End Function

Private Sub Form_Unload(Cancel As Integer)

Dim sts As Integer

'---------------------------------------------- '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
'---------------------------------------------- '棚マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚マスタ")
        End If
    End If
'---------------------------------------------- '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
'---------------------------------------------- '品目マスタ（ワーク）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
'---------------------------------------------- '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
'---------------------------------------------- '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
'---------------------------------------------- '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
'---------------------------------------------- 'メニュー管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "メニュー管理マスタ")
        End If
    End If
'---------------------------------------------- '発番マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "発番マスタ")
        End If
    End If
'---------------------------------------------- '担当者別メニューＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者別メニュー")
        End If
    End If
'---------------------------------------------- '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
'---------------------------------------------- '在庫データ（移動処理用）ＣＬＯＳＥ

    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
'---------------------------------------------- '前借データＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "前借データ")
        End If
    End If
'---------------------------------------------- '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
'---------------------------------------------- '在庫移動歴データＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴データ")
        End If
    End If
'---------------------------------------------- '在庫集計データＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫集計データ")
        End If
    End If
'---------------------------------------------- '月平均出荷数ＣＬＯＳＥ
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "月平均出荷数")
        End If
    End If
'---------------------------------------------- '資材前借ﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材前借ﾃﾞｰﾀ")
        End If
    End If
'---------------------------------------------- '作業実績ﾛｸﾞＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "資材前借ﾃﾞｰﾀ")
        End If
    End If
'---------------------------------------------- 'ファイル環境リセット
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If



    If Next_Step = 1 Then
        sts = Shell("d:\newsdc\exe\F1100501.bat", vbNormalFocus)
        If sts = 0 Then
            MsgBox "[F110050]終了処理の起動に失敗しました。 "
            Call Log_Out(LOG_F, "[F110050]終了処理の起動に失敗しました。")
        End If
    End If


    Set F1100101 = Nothing


    


    End
End Sub
Private Function Tanto_Check_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『担当者コードのチェック』
'
'-------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Tanto_Check_Proc = True

    For i = 0 To M_Gyo
        
        If ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = TYPE_REF Then
        Else
                                '担当者マスタ読み込み
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    ID_KANRI_TBL(ING_No).TANTO_CODE = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                    Exit For
                Case BtErrKeyNotFound
                    
                    '   -------------------------------- エラーメッセージ作成
                    Call Err_Send_Proc("担当者未登録", "", "", "", "")
                    
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                    
                    Tanto_Check_Proc = False
                    Exit Function
                Case Else
                    Sendbuf = Text_Create_Proc()
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ", 0)
                    Exit Function
            End Select
        
        End If
    
    Next i

    If i > M_Gyo Then                       '実際はありえない（担当者が未入力）
        ID_KANRI_TBL(ING_No).Step = Step_Start
        '   -------------------------------- エラーメッセージ作成
        Call Err_Send_Proc("担当者未登録", "", "", "", "")
                    
        Sendbuf = Text_Create_Proc()
        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                    
        Tanto_Check_Proc = False
        Exit Function
    End If

'----------------------------------------------- '専用メニュー獲得
    If Menu_Type = 1 Then
                        '共通メニュー
    Else
        Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
        sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
        Select Case sts
            Case BtNoErr
'                ID_KANRI_TBL(ING_No).MENU_GRP = StrConv(P_TMENUREC.TANTO_CODE, vbUnicode)
            Case BtErrKeyNotFound
                    
'                ID_KANRI_TBL(ING_No).MENU_GRP = ""
                '   -------------------------------- エラーメッセージ作成
                Call Err_Send_Proc("担当者メニュー", "未登録", "", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
            
                Tanto_Check_Proc = False
                Exit Function
            Case Else
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "担当者メニュー", 0)
                Exit Function
        End Select
            
    
    
    End If
'----------------------------------------------- 'メニュー情報＆作業情報の初期化
    
    If UBound(JGYOBU_T) = 0 Then
                                                '１事業部固定
        ID_KANRI_TBL(ING_No).JGYOBU = JGYOBU_T(0).CODE
    Else
        ID_KANRI_TBL(ING_No).JGYOBU = ""
    End If
    
    If UBound(NAIGAI) = 0 Then
                                                '国内外固定
        ID_KANRI_TBL(ING_No).NAIGAI = NAIGAI(0).CODE
    Else
        ID_KANRI_TBL(ING_No).NAIGAI = ""
    End If
    
    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'    ID_KANRI_TBL(ING_No).MENU_LV3 = ""

    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'    ID_KANRI_TBL(ING_No).PageNo_LV3 = 0

    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ""
    ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = ""
    ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = ""

'---------------------------------------------- 'メニュー送信
    If Menu_Send_Proc(Sendbuf) Then
        Exit Function
    End If

    Tanto_Check_Proc = False

End Function
Private Sub Err_Send_Proc(Errmsg0 As String, _
                            Errmsg1 As String, _
                            Errmsg2 As String, _
                            Errmsg3 As String, _
                            Errmsg4 As String)
'-------------------------------------------------------
'
'   『エラーメッセージ電文の作成』
'
'-------------------------------------------------------
    
    Send_Text.sts = Sts_NG                  'ステータス
    Send_Text.Display_Flg = Display_DEF     '表示画面フラグ
    Send_Text.End_Menu = ""                 '最終メニューフラグ
    Send_Text.Menu_Suu = ""                 'メニュー個数
    Send_Text.fileName = ""                 'ファイル名
    Send_Text.Buzzer = Buzzer_CONTI         'ブザー音
'-------------------------------------------------------
    Send_Text.Box_Type(0).Box_Type = TYPE_REF               '行１　 BOX属性
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, Errmsg0)   '       表示内容
    Send_Text.Box_Type(0).INIT = ""                         '       数値初期値
    Send_Text.Box_Type(0).Start_Pos = ""                    '       開始位置
    Send_Text.Box_Type(0).Max_Size = ""                     '       入力桁数
    Send_Text.Box_Type(0).MENU = ""                         '       メニュー内容
'-------------------------------------------------------
    Send_Text.Box_Type(1).Box_Type = TYPE_REF               '行２   BOX属性
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Errmsg1)   '       表示内容
    Send_Text.Box_Type(1).INIT = ""                         '       数値初期値
    Send_Text.Box_Type(1).Start_Pos = ""                    '       開始位置
    Send_Text.Box_Type(1).Max_Size = ""                     '       入力桁数
    Send_Text.Box_Type(1).MENU = ""                         '       メニュー内容
'-------------------------------------------------------
    Send_Text.Box_Type(2).Box_Type = TYPE_REF               '行３　 BOX属性
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Errmsg2)   '       表示内容
    Send_Text.Box_Type(2).INIT = ""                         '       数値初期値
    Send_Text.Box_Type(2).Start_Pos = ""                    '       開始位置
    Send_Text.Box_Type(2).Max_Size = ""                     '       入力桁数
    Send_Text.Box_Type(2).MENU = ""                         '       メニュー内容
'-------------------------------------------------------
    Send_Text.Box_Type(3).Box_Type = TYPE_REF               '行４　 BOX属性
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Errmsg3)   '       表示内容
    Send_Text.Box_Type(3).INIT = ""                         '       数値初期値
    Send_Text.Box_Type(3).Start_Pos = ""                    '       開始位置
    Send_Text.Box_Type(3).Max_Size = ""                     '       入力桁数
    Send_Text.Box_Type(3).MENU = ""                         '       メニュー内容
'-------------------------------------------------------
    Send_Text.Box_Type(4).Box_Type = TYPE_REF               '行５　 BOX属性
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Errmsg4)   '       表示内容
    Send_Text.Box_Type(4).INIT = ""                         '       数値初期値
    Send_Text.Box_Type(4).Start_Pos = ""                    '       開始位置
    Send_Text.Box_Type(4).Max_Size = ""                     '       入力桁数
    Send_Text.Box_Type(4).MENU = ""                         '       メニュー内容
'-------------------------------------------------------
    Send_Text.CRLF = vbCrLf
        
    ID_KANRI_TBL(ING_No).Last_Send = 1                      'エラー送信

End Sub
Private Function Text_Create_Proc() As String
'-------------------------------------------------------
'
'   『送信テキスト作成』
'
'-------------------------------------------------------
Dim i   As Integer
    
    Text_Create_Proc = Send_Text.sts & _
                Send_Text.Display_Flg & _
                Send_Text.End_Menu & _
                Send_Text.Menu_Suu & _
                Send_Text.fileName & _
                Send_Text.Buzzer

    For i = 0 To 4
        Text_Create_Proc = Text_Create_Proc & Send_Text.Box_Type(i).Box_Type & _
                            StrConv(Send_Text.Box_Type(i).LCD, vbUnicode) & _
                            Send_Text.Box_Type(i).INIT & _
                            Send_Text.Box_Type(i).Start_Pos & _
                            Send_Text.Box_Type(i).Max_Size & _
                            Send_Text.Box_Type(i).MENU
    Next i
    
    Text_Create_Proc = Text_Create_Proc & Send_Text.CRLF

End Function
'2006.01.30Private Function Menu_Send_Proc(Optional Sendbuf As String) As Integer
'2006.01.30'-------------------------------------------------------
'2006.01.30'
'2006.01.30'   『メニューテキスト作成』
'2006.01.30'
'2006.01.30'-------------------------------------------------------
'2006.01.30Dim sts         As Integer
'2006.01.30Dim com         As Integer
'2006.01.30
'2006.01.30Dim i           As Integer
'2006.01.30Dim j           As Integer
'2006.01.30
'2006.01.30Dim Menu_Tbl()  As Menu_Tbl_tag
'2006.01.30Dim Menu_Cnt    As Integer
'2006.01.30Dim Max_Page    As Integer
'2006.01.30
'2006.01.30
'2006.01.30Dim PageNo      As Integer
'2006.01.30
'2006.01.30Dim Gyo_Suu     As Integer
'2006.01.30Dim Start_Gyo   As Integer
'2006.01.30Dim End_Gyo     As Integer
'2006.01.30
'2006.01.30
'2006.01.30Dim WK_LV1      As String * 3
'2006.01.30Dim WK_LV2      As String * 3
'2006.01.30Dim WK_LV3      As String * 3
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Send_Proc = True
'2006.01.30'----------------------------------------------- '事業部選択あり
'2006.01.30    If ID_KANRI_TBL(ING_No).JGYOBU = " " Then
'2006.01.30        Call JGYOBU_MENU_SET
'2006.01.30
'2006.01.30        Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30'----------------------------------------------- '国内外選択あり
'2006.01.30    If ID_KANRI_TBL(ING_No).NAIGAI = " " Then
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        Call NAIGAI_MENU_SET
'2006.01.30        Sendbuf = Text_Create_Proc
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30'----------------------------------------------- 'メニュー管理の読み込み
'2006.01.30    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_GRP)) = 0 Then
'2006.01.30                                    'ここで未確定なのは共通メニュー時だけ
'2006.01.30        Call UniCode_Conv(K1_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30        Call UniCode_Conv(K1_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_GRP_NO, ALL_MENU_GRP)
'2006.01.30
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, "")
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV2, "")
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV3, "")
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        sts = BTRV(BtOpGetGreaterEqual, MENU_POS, MENUREC, Len(MENUREC), K1_MENU, Len(K1_MENU), 1)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30            Case BtErrEOF
'2006.01.30
'2006.01.30            '   -------------------------------- エラーメッセージ作成
'2006.01.30                Call Err_Send_Proc("メニュー未登録", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                If UBound(NAIGAI) = 0 Then
'2006.01.30                    ID_KANRI_TBL(ING_No).Step = Step_Start
'2006.01.30                Else
'2006.01.30                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'2006.01.30                End If
'2006.01.30                Menu_Send_Proc = False
'2006.01.30                Exit Function
'2006.01.30            Case Else
'2006.01.30                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, BtOpGetEqual, "メニュー管理マスタ", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
'2006.01.30
'2006.01.30
'2006.01.30    End If
'2006.01.30    '   -------------------------------- メニュー管理マスタ読込み
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, ID_KANRI_TBL(ING_No).MENU_GRP)
'2006.01.30    Call UniCode_Conv(K0_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30    Call UniCode_Conv(K0_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV1, ID_KANRI_TBL(ING_No).MENU_LV1)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV2, ID_KANRI_TBL(ING_No).MENU_LV2)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV3, ID_KANRI_TBL(ING_No).MENU_LV3)
'2006.01.30
'2006.01.30    Erase Menu_Tbl
'2006.01.30
'2006.01.30    com = BtOpGetGreater
'2006.01.30
'2006.01.30    Menu_Cnt = -1
'2006.01.30    Do
'2006.01.30        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30                If ID_KANRI_TBL(ING_No).MENU_GRP <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Or _
'2006.01.30                    ID_KANRI_TBL(ING_No).JGYOBU <> StrConv(MENUREC.JGYOBU, vbUnicode) Or _
'2006.01.30                    ID_KANRI_TBL(ING_No).NAIGAI <> StrConv(MENUREC.NAIGAI, vbUnicode) Then
'2006.01.30                    Exit Do
'2006.01.30                End If
'2006.01.30
'2006.01.30                WK_LV1 = ID_KANRI_TBL(ING_No).MENU_LV1
'2006.01.30                WK_LV2 = ID_KANRI_TBL(ING_No).MENU_LV2
'2006.01.30                WK_LV3 = ID_KANRI_TBL(ING_No).MENU_LV3
'2006.01.30
'2006.01.30
'2006.01.30                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV1)) = 0 Then
'2006.01.30                    WK_LV1 = StrConv(MENUREC.MENU_LV1, vbUnicode)
'2006.01.30                    WK_LV2 = ""
'2006.01.30                    WK_LV3 = ""
'2006.01.30
'2006.01.30
'2006.01.30'                    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30                    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30                    ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30                Else
'2006.01.30                    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) = 0 Then
'2006.01.30                        WK_LV2 = StrConv(MENUREC.MENU_LV2, vbUnicode)
'2006.01.30                        WK_LV3 = ""
'2006.01.30
'2006.01.30'                        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30                        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30                    Else
'2006.01.30                        If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) = 0 Then
'2006.01.30                            WK_LV3 = StrConv(MENUREC.MENU_LV3, vbUnicode)
'2006.01.30
'2006.01.30'                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30                        End If
'2006.01.30                    End If
'2006.01.30                End If
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30            Case BtErrEOF
'2006.01.30                Exit Do
'2006.01.30            Case Else
'2006.01.30
'2006.01.30                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, com, "メニュー管理マスタ", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30
'2006.01.30    '   -------------------------------- メニュー情報を保存
'2006.01.30        If WK_LV1 <> StrConv(MENUREC.MENU_LV1, vbUnicode) Or _
'2006.01.30            WK_LV2 <> StrConv(MENUREC.MENU_LV2, vbUnicode) Or _
'2006.01.30            WK_LV3 <> StrConv(MENUREC.MENU_LV3, vbUnicode) Then
'2006.01.30        Else
'2006.01.30
'2006.01.30
'2006.01.30            Menu_Cnt = Menu_Cnt + 1
'2006.01.30
'2006.01.30            ReDim Preserve Menu_Tbl(Menu_Cnt)
'2006.01.30
'2006.01.30            If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV1)) = 0 Then
'2006.01.30                Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV1, vbUnicode)
'2006.01.30            Else
'2006.01.30                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) = 0 Then
'2006.01.30                    Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV2, vbUnicode)
'2006.01.30                Else
'2006.01.30                    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) = 0 Then
'2006.01.30                        Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV3, vbUnicode)
'2006.01.30                    End If
'2006.01.30                End If
'2006.01.30            End If
'2006.01.30
'2006.01.30
'2006.01.30            Menu_Tbl(Menu_Cnt).Disp = StrConv(MENUREC.DISPLAY_ITEM, vbUnicode)
'2006.01.30        End If
'2006.01.30
'2006.01.30        com = BtOpGetNext
'2006.01.30
'2006.01.30    Loop
'2006.01.30
'2006.01.30    If Menu_Cnt = -1 Then
'2006.01.30            '   -------------------------------- エラーメッセージ作成
'2006.01.30        Call Err_Send_Proc("メニュー未登録", "", "", "", "")
'2006.01.30        Sendbuf = Text_Create_Proc()
'2006.01.30        If UBound(NAIGAI) = 0 Then
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_Start
'2006.01.30        Else
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'2006.01.30        End If
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30
'2006.01.30'----------------------------------------------- 'メニュー送信テキスト作成
'2006.01.30'''    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1   'メニュー送信
'2006.01.30    '---------------------------------------------------------------
'2006.01.30    Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30'    Start_Gyo = ID_KANRI_TBL(ING_No).PageNo * M_Gyo
'2006.01.30'    End_Gyo = (ID_KANRI_TBL(ING_No).PageNo * M_Gyo) + (M_Gyo - 1)
'2006.01.30
'2006.01.30
'2006.01.30    Select Case ID_KANRI_TBL(ING_No).Step
'2006.01.30        Case Step_MENU1_REQ, Step_MENU1_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV1
'2006.01.30        Case Step_MENU2_REQ, Step_MENU2_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV2
'2006.01.30        Case Step_MENU3_REQ, Step_MENU3_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV3
'2006.01.30    End Select
'2006.01.30
'2006.01.30
'2006.01.30    Start_Gyo = PageNo * M_Gyo
'2006.01.30    End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)
'2006.01.30
'2006.01.30
'2006.01.30    Send_Text.sts = Sts_OK                                      'ステータス　OK
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
'2006.01.30
'2006.01.30    Send_Text.Display_Flg = Display_MENU                        '表示画面フラグ メニュー画面
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
'2006.01.30                                                                '最終メニューフラグ
'2006.01.30    If Max_Page = 1 Then
'2006.01.30        Send_Text.End_Menu = Menu_Only          '１画面のみ
'2006.01.30        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
'2006.01.30    Else
'2006.01.30        If (Max_Page - 1) = PageNo Then
'2006.01.30            Send_Text.End_Menu = Menu_End       '最終ページ
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_End
'2006.01.30        Else
'2006.01.30            If PageNo = 0 Then
'2006.01.30                Send_Text.End_Menu = Menu_Head  '先頭ページ
'2006.01.30                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
'2006.01.30            Else
'2006.01.30                Send_Text.End_Menu = Menu_Mid   '途中ページ
'2006.01.30                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
'2006.01.30            End If
'2006.01.30        End If
'2006.01.30    End If
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Send_Text.fileName = ""                                         '送信データファイル名
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
'2006.01.30
'2006.01.30    Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
'2006.01.30    '---------------------------------------------------------------
'2006.01.30    Gyo_Suu = 0
'2006.01.30    j = -1
'2006.01.30    For i = Start_Gyo To End_Gyo
'2006.01.30        j = j + 1
'2006.01.30        If i > UBound(Menu_Tbl) Then
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Box_Type = ""                 'BOX属性
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
'2006.01.30
'2006.01.30            Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '表示内容
'2006.01.30            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).INIT = ""                     '数値初期値
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).MENU = ""                     'メニュ―番号
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
'2006.01.30
'2006.01.30
'2006.01.30        Else
'2006.01.30            Gyo_Suu = Gyo_Suu + 1
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX属性
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
'2006.01.30                                                                '表示内容
'2006.01.30            Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
'2006.01.30            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).INIT = ""                     '数値初期値
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).MENU = Menu_Tbl(i).CODE       'メニュ―番号
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).CODE
'2006.01.30
'2006.01.30        End If
'2006.01.30
'2006.01.30    Next i
'2006.01.30
'2006.01.30    Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      'メニュー項目数
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
'2006.01.30
'2006.01.30
'2006.01.30    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
'2006.01.30
'2006.01.30    Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Send_Proc = False
'2006.01.30
'2006.01.30End Function

Private Function Menu_Send_Proc(Optional Sendbuf As String) As Integer

'-------------------------------------------------------
'
'   『メニューテキスト作成』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Menu_Tbl()  As Menu_Tbl_tag
Dim Menu_Cnt    As Integer
Dim Max_Page    As Integer


Dim PageNo      As Integer

Dim Gyo_Suu     As Integer
Dim Start_Gyo   As Integer
Dim End_Gyo     As Integer


Dim WK_LV1      As String * 3
Dim WK_LV2      As String * 3


    Menu_Send_Proc = True
'----------------------------------------------- '事業部選択あり
    If Trim(ID_KANRI_TBL(ING_No).JGYOBU) = "" Then
        Call JGYOBU_MENU_SET

        Sendbuf = Text_Create_Proc()


        Menu_Send_Proc = False
        Exit Function
    End If
'----------------------------------------------- '国内外選択あり
    If Trim(ID_KANRI_TBL(ING_No).NAIGAI) = " " Then

        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        ID_KANRI_TBL(ING_No).MENU_LV2 = ""

        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0


        Call NAIGAI_MENU_SET
        Sendbuf = Text_Create_Proc
        Menu_Send_Proc = False
        Exit Function
    End If
    '   -------------------------------- レベル１　トップメニューの管理
    If Trim(ID_KANRI_TBL(ING_No).MENU_LV1) = "" Then
        'ﾒﾆｭｰｸﾞﾙｰﾌﾟ
        Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
        Erase Menu_Tbl

        sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
        Select Case sts
            
            Case BtNoErr
            Case BtErrKeyNotFound
            
                        
                Call Err_Send_Proc("メニュー未登録", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                If UBound(NAIGAI) = 0 Then
                    ID_KANRI_TBL(ING_No).Step = Step_Start
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                End If
                Menu_Send_Proc = False
                Exit Function
            
            
            Case Else

                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, com, "担当者別ﾒﾆｭｰ", 0)
                Exit Function
        End Select


        Menu_Cnt = -1
        For i = 0 To 29
            If Trim(StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)) = "" Then
                Exit For
            End If
        
            If StrConv(P_TMENUREC.MENU_T(i).JGYOBU, vbUnicode) = ID_KANRI_TBL(ING_No).JGYOBU Or _
                StrConv(P_TMENUREC.MENU_T(i).NAIGAI, vbUnicode) = ID_KANRI_TBL(ING_No).NAIGAI Then
        
                Menu_Cnt = Menu_Cnt + 1
                ReDim Preserve Menu_Tbl(Menu_Cnt)
            
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)
                
                Call UniCode_Conv(K0_P_MENU.JGYOBU, StrConv(P_TMENUREC.MENU_T(i).JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_MENU.NAIGAI, StrConv(P_TMENUREC.MENU_T(i).NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_MENU.MENU_NO, StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                Select Case sts
                    
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                                
                        Call Err_Send_Proc("メニュー異常", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        If UBound(NAIGAI) = 0 Then
                            ID_KANRI_TBL(ING_No).Step = Step_Start
                        Else
                            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                        End If
                        Menu_Send_Proc = False
                        Exit Function
                    
                    
                    Case Else
        
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, com, "メニュー管理マスタ", 0)
                        Exit Function
                End Select
                
                
                
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)
                Menu_Tbl(Menu_Cnt).Disp = StrConv(P_MENUREC.MENU_DSP, vbUnicode)
        
        
                        
            End If
        
        Next i


        If Menu_Cnt = -1 Then
        '   -------------------------------- エラーメッセージ作成
            Call Err_Send_Proc("メニュー未登録", "", "", "", "")
            Sendbuf = Text_Create_Proc()
            If UBound(NAIGAI) = 0 Then
                ID_KANRI_TBL(ING_No).Step = Step_Start
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
            End If
            Menu_Send_Proc = False
            Exit Function
        End If


        Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
        PageNo = ID_KANRI_TBL(ING_No).PageNo_LV1

        Start_Gyo = PageNo * M_Gyo
        End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)


        Send_Text.sts = Sts_OK                                      'ステータス　OK
        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK

        Send_Text.Display_Flg = Display_MENU                        '表示画面フラグ メニュー画面
        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
        If Max_Page = 1 Then
            Send_Text.End_Menu = Menu_Only          '１画面のみ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
        Else
            If (Max_Page - 1) = PageNo Then
                Send_Text.End_Menu = Menu_End       '最終ページ
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_End
            Else
                If PageNo = 0 Then
                    Send_Text.End_Menu = Menu_Head  '先頭ページ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
                Else
                    Send_Text.End_Menu = Menu_Mid   '途中ページ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
                End If
            End If
        End If
        Send_Text.fileName = ""                                         '送信データファイル名
        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
        Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
'---------------------------------------------------------------
        Gyo_Suu = 0
        j = -1
        For i = Start_Gyo To End_Gyo
            j = j + 1
            If i > UBound(Menu_Tbl) Then
                Send_Text.Box_Type(j).Box_Type = ""                 'BOX属性
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
                Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '表示内容
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
                Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                Send_Text.Box_Type(j).MENU = ""                     'メニュ―番号
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
            Else
                Gyo_Suu = Gyo_Suu + 1
                Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX属性
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
                Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO
                
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO
            
            
            End If
        Next i
        
        Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      'メニュー項目数
        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
        ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
        Sendbuf = Text_Create_Proc()
        
    Else
        '   -------------------------------- レベル２　作業メニューの管理
        If Trim(ID_KANRI_TBL(ING_No).MENU_LV2) = "" Then
            
            
            
            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                        ID_KANRI_TBL(ING_No).MENU_LV1, _
                                        "ST", , , , , , , , , FILE_RETRY) Then
                            
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            
            End If
            
            
            
            
            
            
            
            
            
            
            'ﾒﾆｭｰｸﾞﾙｰﾌﾟ
            Call UniCode_Conv(K0_P_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
            Call UniCode_Conv(K0_P_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
            Call UniCode_Conv(K0_P_MENU.MENU_NO, ID_KANRI_TBL(ING_No).MENU_LV1)
            
            
            Erase Menu_Tbl
    
            sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
            Select Case sts
                
                Case BtNoErr
                Case BtErrKeyNotFound
                
                            
                    Call Err_Send_Proc("メニュー未登録", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    If UBound(NAIGAI) = 0 Then
                        ID_KANRI_TBL(ING_No).Step = Step_Start
                    Else
                        ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                    End If
                    Menu_Send_Proc = False
                    Exit Function
                
                
                Case Else
    
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, com, "担当者別ﾒﾆｭｰ", 0)
                    Exit Function
            End Select
    
    
            Menu_Cnt = -1
            For i = 0 To 19
                If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = "" Then
                    Exit For
                End If
            
            
                Menu_Cnt = Menu_Cnt + 1
                ReDim Preserve Menu_Tbl(Menu_Cnt)
                
                    
                    
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)
                Menu_Tbl(Menu_Cnt).PARAM = StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode)
                Menu_Tbl(Menu_Cnt).Disp = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
            
                Menu_Tbl(Menu_Cnt).Log_Out = StrConv(P_MENUREC.SAGYO(i).Log_Out, vbUnicode)
            
            
            
            
            Next i
    
    
            If Menu_Cnt = -1 Then
            '   -------------------------------- エラーメッセージ作成
                Call Err_Send_Proc("メニュー未登録", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                If UBound(NAIGAI) = 0 Then
                    ID_KANRI_TBL(ING_No).Step = Step_Start
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                End If
                Menu_Send_Proc = False
                Exit Function
            End If
    
    
            Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV2
    
            Start_Gyo = PageNo * M_Gyo
            End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)
    
    
            Send_Text.sts = Sts_OK                                      'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_MENU                        '表示画面フラグ メニュー画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
            If Max_Page = 1 Then
                Send_Text.End_Menu = Menu_Only          '１画面のみ
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
            Else
                If (Max_Page - 1) = PageNo Then
                    Send_Text.End_Menu = Menu_End       '最終ページ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_End
                Else
                    If PageNo = 0 Then
                        Send_Text.End_Menu = Menu_Head  '先頭ページ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
                    Else
                        Send_Text.End_Menu = Menu_Mid   '途中ページ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
                    End If
                End If
            End If
            Send_Text.fileName = ""                                         '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
            Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    '---------------------------------------------------------------
            Gyo_Suu = 0
            j = -1
            For i = Start_Gyo To End_Gyo
                j = j + 1
                If i > UBound(Menu_Tbl) Then
                    Send_Text.Box_Type(j).Box_Type = ""                 'BOX属性
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
                    Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '表示内容
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
                    Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                    Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                    Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                    Send_Text.Box_Type(j).MENU = ""                     'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
                Else
                    Gyo_Suu = Gyo_Suu + 1
                    Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX属性
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
                    Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                    Send_Text.Box_Type(j).INIT = ""                     '数値初期値
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                    Send_Text.Box_Type(j).Start_Pos = ""                '初期カーソル位置
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                    Send_Text.Box_Type(j).Max_Size = "00"               '入力桁数
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                                                                        'メニュ―番号 & ﾊﾟﾗﾒｰﾀ
                    Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM
                    
''''2006.05.31                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM
                
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Trim(CStr(Dec_To_Bcd(Menu_Tbl(i).PARAM)))
                
                
                
                End If
            Next i
            
            Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      'メニュー項目数
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
            ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
            Sendbuf = Text_Create_Proc()
            
        End If
    End If
    
    Menu_Send_Proc = False




End Function


Private Sub JGYOBU_MENU_SET()
'-------------------------------------------------------
'
'   『事業部選択用メニュー作成』
'
'-------------------------------------------------------
Dim i   As Integer

    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ                     '事業部要求
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_MENU                            '表示画面フラグ メニュー画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
    
    Send_Text.End_Menu = Menu_Only                                  '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = Format(UBound(JGYOBU_T) + 1, "00")         'メニュー項目数
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(UBound(JGYOBU_T) + 1, "00")
    
    Send_Text.fileName = ""                                         '送信データファイル名
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    
    '---------------------------------------------------------------
    For i = 0 To M_Gyo - 1
        
        If i > UBound(JGYOBU_T) Then
        
            Send_Text.Box_Type(i).Box_Type = ""                 'BOX属性
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
        
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")    '表示内容
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
            
            Send_Text.Box_Type(i).INIT = ""                     '数値初期値
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                                                
            Send_Text.Box_Type(i).Start_Pos = ""                '初期カーソル位置
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '入力桁数
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
        
        Else
            
            Send_Text.Box_Type(i).Box_Type = TYPE_MENU          'BOX属性
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = TYPE_MENU
                                                                '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, JGYOBU_T(i).NAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, JGYOBU_T(i).NAME)
                                                                                
            Send_Text.Box_Type(i).INIT = ""                     '数値初期値
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
            
            Send_Text.Box_Type(i).Start_Pos = ""                '初期カーソル位置
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '入力桁数
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = JGYOBU_T(i).CODE       'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = JGYOBU_T(i).CODE
                                                                                
        
        End If
    
    Next i
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信

End Sub

Private Sub NAIGAI_MENU_SET()
'-------------------------------------------------------
'
'   『内外選択用メニュー作成』
'
'-------------------------------------------------------
Dim i   As Integer

    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ                     '内外要求
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_MENU                             '表示画面フラグ メニュー画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
    
    Send_Text.End_Menu = Menu_Only                                  '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = Format(UBound(NAIGAI) + 1, "00")         'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(UBound(NAIGAI) + 1, "00")
    
    Send_Text.fileName = ""                                         '送信データファイル名
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    
    '---------------------------------------------------------------
    For i = 0 To M_Gyo - 1
        
        If i > UBound(NAIGAI) Then
        
            Send_Text.Box_Type(i).Box_Type = ""                 'BOX属性
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
        
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")    '表示内容
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
            
            Send_Text.Box_Type(i).INIT = ""                     '数値初期値
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                                                
            Send_Text.Box_Type(i).Start_Pos = ""                '初期カーソル位置
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '入力桁数
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
        
        Else
            
            Send_Text.Box_Type(i).Box_Type = TYPE_MENU          'BOX属性
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = TYPE_MENU
                                                                '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, NAIGAI(i).NAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, NAIGAI(i).NAME)
                                                                                
            Send_Text.Box_Type(i).INIT = ""                     '数値初期値
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                                                
            Send_Text.Box_Type(i).Start_Pos = ""                '初期カーソル位置
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '入力桁数
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = NAIGAI(i).CODE       'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = NAIGAI(i).CODE
                                                                                
        
        End If
    
    Next i
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信

End Sub

Private Sub Re_Send_Proc(Sendbuf As String)
'-------------------------------------------------------
'
'   『エラー時の再送』
'
'-------------------------------------------------------
Dim i   As Integer
    
    
    
'    Select Case ID_KANRI_TBL(ING_No).Step
'        Case 2, 4, 6
'            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'    End Select
'-------------------------------------------------------
    Send_Text.sts = ID_KANRI_TBL(ING_No).Send_Text.sts                  'ステータス　OK
    
    Send_Text.Display_Flg = ID_KANRI_TBL(ING_No).Send_Text.Display_Flg  '表示画面フラグ メニュー画面
    
    Send_Text.End_Menu = ID_KANRI_TBL(ING_No).Send_Text.End_Menu        '最終メニューフラグ
    
    Send_Text.Menu_Suu = ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu        'メニュー項目数（05固定）
    
    Send_Text.fileName = ID_KANRI_TBL(ING_No).Send_Text.fileName        '送信データファイル名
    
    Send_Text.Buzzer = ID_KANRI_TBL(ING_No).Send_Text.Buzzer            'ブザー音　標準

    For i = 0 To M_Gyo - 1
                                                                        'BOX属性
        Send_Text.Box_Type(i).Box_Type = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type
                                                                        '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                                                                        
                                                                    '初期表示内容（数値）
        Send_Text.Box_Type(i).INIT = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT
                                                                        '初期カーソル位置
        Send_Text.Box_Type(i).Start_Pos = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos
                                                                        '入力桁数
        Send_Text.Box_Type(i).Max_Size = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size
                                                                        'メニュ―番号
        Send_Text.Box_Type(i).MENU = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU

    Next
    
    Send_Text.CRLF = vbCrLf
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
    
    Sendbuf = Text_Create_Proc()


End Sub
Private Sub Send_Err_Proc(Sendbuf As String)
'-------------------------------------------------------
'
'   『送信エラー時の再送』
'
'-------------------------------------------------------
Dim i   As Integer
    
    
    
'    Select Case ID_KANRI_TBL(ING_No).Step
'        Case 2, 4, 6
'            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'    End Select
'-------------------------------------------------------
    Send_Text.sts = ID_KANRI_TBL(ING_No).Last_Send_Text.sts                  'ステータス　OK
    
    Send_Text.Display_Flg = ID_KANRI_TBL(ING_No).Last_Send_Text.Display_Flg  '表示画面フラグ メニュー画面
    
    Send_Text.End_Menu = ID_KANRI_TBL(ING_No).Last_Send_Text.End_Menu        '最終メニューフラグ
    
    Send_Text.Menu_Suu = ID_KANRI_TBL(ING_No).Last_Send_Text.Menu_Suu        'メニュー項目数（05固定）
    
    Send_Text.fileName = ID_KANRI_TBL(ING_No).Last_Send_Text.fileName        '送信データファイル名
    
    Send_Text.Buzzer = ID_KANRI_TBL(ING_No).Last_Send_Text.Buzzer            'ブザー音　標準

    For i = 0 To M_Gyo - 1
                                                                        'BOX属性
        Send_Text.Box_Type(i).Box_Type = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).Box_Type
                                                                        '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, StrConv(ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).LCD, vbUnicode))
                                                                        
                                                                    '初期表示内容（数値）
        Send_Text.Box_Type(i).INIT = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).INIT
                                                                        '初期カーソル位置
        Send_Text.Box_Type(i).Start_Pos = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).Start_Pos
                                                                        '入力桁数
        Send_Text.Box_Type(i).Max_Size = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).Max_Size
                                                                        'メニュ―番号
        Send_Text.Box_Type(i).MENU = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).MENU

    Next
    
    Send_Text.CRLF = vbCrLf
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信
    
    Sendbuf = Text_Create_Proc()


End Sub

Private Function Menu_Recv_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『２階層以上のメニュー送信』
'
'-------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    
Dim MTS     As String * 8
Dim SS      As String * 8
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_MENU1_RES
    '   -------------------------------- 次ﾚﾍﾞﾙへ
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> BEF_Page And Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> NEXT_Page Then

                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
            End If
            
            If Menu_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        Case Else
    
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = BEF_Page Or Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = NEXT_Page Then
                If Menu_Send_Proc() Then
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                End If
            Else
    
    '   -------------------------------- メニュー管理マスタ読込み
                Call UniCode_Conv(K0_P_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_P_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_P_MENU.MENU_NO, ID_KANRI_TBL(ING_No).MENU_LV1)
        
                sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                Select Case sts
                    Case BtNoErr
    
                        For i = 0 To 19
                        
                            If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = Trim(ID_KANRI_TBL(ING_No).MENU_LV2) And _
                                Left(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 7) = (ID_KANRI_TBL(ING_No).MTS_CODE & _
                                                                                        ID_KANRI_TBL(ING_No).SS_CODE) Then
                                
                                Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                
                                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                                Select Case sts
                                
                                    Case BtNoErr
                                        'ｽｷｬﾅ表示名称
                                        ID_KANRI_TBL(ING_No).YOIN_DNAME = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
                                
                                        
                                        ID_KANRI_TBL(ING_No).SAGYO_LOG = StrConv(P_MENUREC.SAGYO(i).Log_Out, vbUnicode)
                                                                    
                                        If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then   '向け先なら（出荷）
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
'2006.01.30                                            ID_KANRI_TBL(ING_No).MTS_CODE = Left(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
'2006.01.30                                            ID_KANRI_TBL(ING_No).SS_CODE = Right(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
                                        End If
                                                                                            '検品（向け先指定）なら
                                        If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_MTS Then
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        End If
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(YOINREC.CODE_TYPE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(YOINREC.Soko_No, vbUnicode)
                                
                                    Case BtErrKeyNotFound
    
                                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                                        '   -------------------------------- エラーメッセージ作成
                                        Call Err_Send_Proc("要因マスタ", "未登録", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                                        Menu_Recv_Proc = False
                                        Exit Function
                                    Case Else
                                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
                                        Exit Function
                                End Select
                                
                                Exit For
                            End If
                        
                        Next i
    
                        If i > 19 Then
                
    
                            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                            '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc("メニュー管理マスタ", "設定ミス", Trim(ID_KANRI_TBL(ING_No).MENU_LV1), "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                            Menu_Recv_Proc = False
                            Exit Function
                        
                        End If
                
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
        
        
                    Case BtErrKeyNotFound
        
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc("メニュー管理マスタ", "未登録", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                        Menu_Recv_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
                        Exit Function
        
                End Select
            End If
        End Select
        Sendbuf = Text_Create_Proc()
    
    Menu_Recv_Proc = False

End Function
'2006.01.30Private Function Menu_Recv_Proc(Sendbuf As String) As Integer
'2006.01.30'-------------------------------------------------------
'2006.01.30'
'2006.01.30'   『２階層以上のメニュー送信』
'2006.01.30'
'2006.01.30'-------------------------------------------------------
'2006.01.30Dim sts As Integer
'2006.01.30
'2006.01.30    Menu_Recv_Proc = True
'2006.01.30                                        'メニュ管理マスタの読み込み
'2006.01.30    Call UniCode_Conv(K1_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30    Call UniCode_Conv(K1_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_GRP_NO, ID_KANRI_TBL(ING_No).MENU_GRP)
'2006.01.30
'2006.01.30    If ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ Then
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, "000")
'2006.01.30    Else
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, ID_KANRI_TBL(ING_No).MENU_LV1)
'2006.01.30    End If
'2006.01.30
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_LV2, ID_KANRI_TBL(ING_No).MENU_LV2)
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_LV3, ID_KANRI_TBL(ING_No).MENU_LV3)
'2006.01.30    sts = BTRV(BtOpGetEqual, MENU_POS, MENUREC, Len(MENUREC), K1_MENU, Len(K1_MENU), 1)
'2006.01.30    Select Case sts
'2006.01.30        Case BtNoErr
'2006.01.30        Case BtErrKeyNotFound
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30            '   -------------------------------- エラーメッセージ作成
'2006.01.30            Call Err_Send_Proc("メニュー管理", "未登録", "", "", "")
'2006.01.30
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'2006.01.30
'2006.01.30            Menu_Recv_Proc = False
'2006.01.30            Exit Function
'2006.01.30        Case Else
'2006.01.30            Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
'2006.01.30            Exit Function
'2006.01.30    End Select
'2006.01.30
'2006.01.30    If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> BEF_Page And Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> NEXT_Page And _
'2006.01.30     StrConv(MENUREC.MENU_KBN, vbUnicode) = "1" Then
'2006.01.30
'2006.01.30
'2006.01.30                                            '要因の読込み
'2006.01.30        Call UniCode_Conv(K0_YOIN.CODE_TYPE, StrConv(MENUREC.CODE_TYPE, vbUnicode))
'2006.01.30        Call UniCode_Conv(K0_YOIN.YOIN_CODE, StrConv(MENUREC.YOIN_CODE, vbUnicode))
'2006.01.30        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30                ID_KANRI_TBL(ING_No).YOIN_DNAME = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
'2006.01.30
'2006.01.30                If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then   '向け先なら（出荷）
'2006.01.30                    ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(MENUREC.YOIN_CODE, vbUnicode)
'2006.01.30                    ID_KANRI_TBL(ING_No).MTS_CODE = Left(StrConv(MENUREC.PARAM, vbUnicode), 8)
'2006.01.30                    ID_KANRI_TBL(ING_No).SS_CODE = Right(StrConv(MENUREC.PARAM, vbUnicode), 8)
'2006.01.30                End If
'2006.01.30                                                                    '検品（向け先指定）なら
'2006.01.30                If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_MTS Then
'2006.01.30                    ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(MENUREC.YOIN_CODE, vbUnicode)
'2006.01.30                End If
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(YOINREC.CODE_TYPE, vbUnicode)
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(YOINREC.YOIN_CODE, vbUnicode)
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(MENUREC.PARAM, vbUnicode)
'2006.01.30
'2006.01.30            Case BtErrKeyNotFound
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30                '   -------------------------------- エラーメッセージ作成
'2006.01.30                Call Err_Send_Proc("要因マスタ", "未登録", "", "", "")
'2006.01.30
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'2006.01.30
'2006.01.30                Menu_Recv_Proc = False
'2006.01.30                Exit Function
'2006.01.30            Case Else
'2006.01.30                Call Err_Send_Proc("システム異常発生", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, BtOpGetEqual, "メニュー管理", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30        If Sagyo_Send_Proc() Then
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Exit Function
'2006.01.30        End If
'2006.01.30    Else
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
'2006.01.30
'2006.01.30        If Menu_Send_Proc() Then
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Exit Function
'2006.01.30        End If
'2006.01.30    End If
'2006.01.30
'2006.01.30    Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Recv_Proc = False
'2006.01.30
'2006.01.30End Function




Private Function Sagyo_Send_Proc() As Integer
'-------------------------------------------------------
'
'   『作業の送信』
'
'-------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer

Dim Found_Flg   As Boolean

Dim sts         As Integer

    Sagyo_Send_Proc = True
    
                        '要因の読み込み
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    
    Select Case sts
        Case BtNoErr
        '   -------------------------------- エラーメッセージ作成
        Case Else
        '「要因未登録は考えられないエラーシステム停止とする」
            Call Err_Send_Proc("システム異常発生", "", "", "", "")
'                Sendbuf = Text_Create_Proc()
            Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
            Exit Function
    End Select
    
    
    
    '   -------------------------------- 送信パラメータの検索
    Found_Flg = False
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '最初は２桁で検索
            If StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode) = _
                WEL_Para_Tbl(i, j).Action Then
                Found_Flg = True
                Exit For
            End If
        Next j
            
        If Found_Flg Then
            Exit For
        End If
    
    Next i


    If Not Found_Flg Then
        
        For i = 0 To UBound(WEL_Para_Tbl, 1)
            For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '最後は１桁で検索
               If StrConv(YOINREC.CODE_TYPE, vbUnicode) = Left(WEL_Para_Tbl(i, j).Action, 1) Then
                    Found_Flg = True
                    Exit For
                End If
        
            Next j
            
            If Found_Flg Then
                Exit For
            End If
        
        
        Next i
            
    End If

    If Not Found_Flg Then
        '信じられないエラー
        Call Log_Out(LOG_F, "要因マスタ＜＞WLECATパラメータ（INIファイル）")
        Exit Function
    End If

    '   -------------------------------- 作業作成
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ                     '通常作業開始
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                             '表示画面フラグ 通常入力画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                                  '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                                       'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                         '送信データファイル名
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    
                                                                    'パラメータの1行目には要因名称をセット
    WEL_Para_Tbl(i, j).Wel_Para(0).LCD = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
    '---------------------------------------------------------------
    For k = 0 To M_Gyo - 1
        
                                                            'BOX属性
        Send_Text.Box_Type(k).Box_Type = WEL_Para_Tbl(i, j).Wel_Para(k).Box_Type
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Box_Type = WEL_Para_Tbl(i, j).Wel_Para(k).Box_Type
                                                            
                                                            '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(k).LCD, WEL_Para_Tbl(i, j).Wel_Para(k).LCD)
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).LCD, WEL_Para_Tbl(i, j).Wel_Para(k).LCD)
                                                            '数値初期表示
        Send_Text.Box_Type(k).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).INIT = ""
                                                            
                                                            '初期カーソル位置
        If Send_Text.Box_Type(k).Box_Type = "2" Then
            Send_Text.Box_Type(k).Start_Pos = Format(M_Keta - WEL_Para_Tbl(i, j).Wel_Para(k).Keta + 1, "00")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Start_Pos = Format(M_Keta - WEL_Para_Tbl(i, j).Wel_Para(k).Keta, "00")
        Else
            Send_Text.Box_Type(k).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Start_Pos = "01"
        End If
                                                            '入力桁数
        Send_Text.Box_Type(k).Max_Size = Format(WEL_Para_Tbl(i, j).Wel_Para(k).Keta, "00")
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Max_Size = Format(WEL_Para_Tbl(i, j).Wel_Para(k).Keta, "00")
                                                                                
        Send_Text.Box_Type(k).MENU = ""                     'メニュ―番号
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).MENU = ""
                                                                                
    
    Next k
    
'    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(MENUREC.CODE_TYPE, vbUnicode)
'    ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(MENUREC.YOIN_CODE, vbUnicode)
    ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(YOINREC.Soko_No, vbUnicode)
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  'ノーマルデータ送信

    Sagyo_Send_Proc = False


End Function

Private Function Sagyo_Main_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『作業受信時のメイン処理』
'
'-------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Found_Flg   As Boolean
    
    
    Sagyo_Main_Proc = True
    
    
    Found_Flg = False
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '最初は２桁で検索
            If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = _
                WEL_Para_Tbl(i, j).Action Then
                Found_Flg = True
                Exit For
            End If
        
        Next j
            
        If Found_Flg Then
            Exit For
        End If
    
    Next i
    
    If Not Found_Flg Then
        
        For i = 0 To UBound(WEL_Para_Tbl, 1)
            For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '最後は１桁で検索
               If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = Left(WEL_Para_Tbl(i, j).Action, 1) Then
                    Found_Flg = True
                    Exit For
                End If
        
            Next j
            
            If Found_Flg Then
                Exit For
            End If
        
        
        Next i
            
    End If


    If Not Found_Flg Then
                        'ありえない異常（該当作業パラメータなし）
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = ""
                
        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                
        '   -------------------------------- エラーメッセージ作成
        Call Err_Send_Proc("作業パラメータ（INI）", "未登録", "", "", "")
                    
        Sendbuf = Text_Create_Proc()
        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
            
        Sagyo_Main_Proc = False
        Exit Function
    
    End If

    Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE
        Case ACT_ZAITEI_IN          '在訂＋
        
        
            If Zaitei_In_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_ZAITEI_OUT         '在訂－
            
            If Zaitei_Out_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_NYUKA              '入荷
    
        Case ACT_SYUKA_KEI          '出荷(出荷予定有り)→向け先宣言
        
            If MTS_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_SYUKA_HYO          '出荷(出庫表)
        
            If SYUKO_HYO_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        Case ACT_SYUKA_GAI          '出荷(出荷予定無し)
        
            If Out_Plan_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_IDO_IN             '移動入庫
        
            If Ido_In_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_IDO_OUT            '移動出庫
        
            If Ido_Out_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_DENPYO_ID          '伝票ＩＤ
        
            If DEN_ID_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        Case ACT_KENPIN             '検品
        
            If Inspe_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_WEL_ETC            'WEL専用（照会）
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
            
                Case Wel_TANAOROSI      '「WEL 棚卸し」の要因
                
                    If Tanaorosi_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_TANAHYOJI      '「WEL 棚番表示」の要因
                
                    If Tanahyoji_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                
                Case Wel_HIN_SHOGO      '「WEL 品番別照合」の要因
                    
                    If Hin_Shogo_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_AVE_SYUKA      '「WEL 月平均出荷数」の要因
                
                    If Ave_Syuka_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_HOST_ZAIKO     '「WEL ホスト在庫照会」の要因
                    
                    If Host_Zaiko_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                            
                Case Wel_ST_TANABAN     '「WEL 標準棚番設定」の要因
                    
                    If St_Tanaban_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                            
            
                Case Wel_RIREKI         '「WEL 当日出庫履歴」の要因
                
                    If Rireki_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_SUII           '「WEL 出荷推移」の要因

                    If Suii_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_TANA_HIN_SHOGO '「WEL 棚番・品番別照合」の要因

                    If Tana_Hin_Shogo_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_TANAHYOJI_KASO '「WEL 棚番表示(仮想優先)」の要因
                
                    If Tanahyoji_Kaso_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            End Select
    
        Case ACT_KENPIN_MTS             '検品（ＭＴＳ読み込みあり）
        
            If Inspe_Proc_MTS(Sendbuf, i, j) Then
                Exit Function
            End If
    
    
        Case ACT_GOODS_ONFF             '商品／未商品切り替え
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
    
                Case Wel_GOODS_ONOFF_ONO        '「WEL 商品/未商品切り替え　小野」の要因
                
                    If GOODS_ONOFF_Ono_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_GOODS_ONOFF_SIGA       '「WEL 商品/未商品切り替え　滋賀」の要因
            
                    If GOODS_ONOFF_Siga_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            End Select
    
    
        Case ACT_SPECIAL_PROCESS    '特殊処理
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
    
                Case Wel_RETURNED_GOODS         '「良品返品」の要因
                
                    If RETURNED_GOODS_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                
                Case Wel_LOCATION_MOVE         '「棚移動」の要因
                
                    If Location_Move_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
            
            End Select
        
    
    End Select

    Sagyo_Main_Proc = False

End Function

Private Function Zaitei_In_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『在訂（＋）指定時のチェック＆更新処理』
'
'-------------------------------------------------------
Dim i           As Integer
Dim Hinban      As String * 13
Dim Tanaban     As String * 8
Dim sts         As Integer

Dim QTY         As Long
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1

Dim MENU_NO     As String * 2

    Zaitei_In_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Tanaban        '棚番
                Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                
                If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                Else
                '------------------ 倉庫マスタ読込み
                    Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                            Zaitei_In_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                            Exit Function
                    End Select
                    '------------------ 混載チェック
                    If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                        If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                            StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Zaitei_In_Proc = False
                            Exit Function
                        End If
                    End If
                    '------------------ 棚マスタ読込み
                    Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                    Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                    Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                    Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Zaitei_In_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                            Exit Function
                    End Select
            
                    '------------------ 禁止棚のチェック
                    If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                        Zaitei_In_Proc = False
                        Exit Function
                    End If
            
            
                End If
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ 品目マスタ読込み
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Zaitei_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                        End If
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Zaitei_In_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                End Select
            
            Case LCD_Suryo          '数量（ここは無い）
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_In_Proc = False
                    Exit Function
                End If
                
                QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                If QTY = 0 Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_In_Proc = False
                    Exit Function
                End If
            
            
            Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_In_Proc = False
                    Exit Function
                End If
                
                
                If Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD) = LCD_SUMI_Suryo Then
                    SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                Else
                    MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                End If
        
                If i = M_Gyo - 1 Then
                    If SUMI_QTY = 0 And MI_QTY = 0 Then
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "商品／未商品＝０", "数量入力ミス", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                        Zaitei_In_Proc = False
                        Exit Function
                    End If
                End If
        
        End Select
    Next i
    '----------------------------------- データ更新処理開始 -----------
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Exit Function
    End If
                                        
                                        
    If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
        MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                
    Else
                
        MENU_NO = ""
    End If
                                        
                                        
    If RET_JGYOBU = SHIZAI Then
        Call UniCode_Conv(K0_ITEM.JGYOBU, RET_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, RET_NAIGAI)
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case Else
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
        End Select
    
    
    
    
        sts = Nyuko_Update_Proc(RET_JGYOBU, _
                                RET_NAIGAI, _
                                Hinban, _
                                Format(Now, "YYYYMMDD"), _
                                Tanaban, _
                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                SUMI_QTY, _
                                MI_QTY, _
                                Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                FILE_RETRY, , _
                                StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode), _
                                StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode), , _
                                MENU_NO)
 
        Select Case sts
            Case False
            Case True           '入庫時は発生しない
            Case SYS_CANCEL
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                Zaitei_In_Proc = False
                GoTo Abort_Tran
            Case SYS_ERR
                Sendbuf = Text_Create_Proc()
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Zaitei_In_Proc = SYS_ERR    'システム異常発生
                
                GoTo Abort_Tran
        End Select
    
    
    
    
    
    
    Else
                                        
                                            
                                            '入庫更新
        sts = Nyuko_Update_Proc(RET_JGYOBU, _
                                RET_NAIGAI, _
                                Hinban, _
                                Format(Now, "YYYYMMDD"), _
                                Tanaban, _
                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                SUMI_QTY, _
                                MI_QTY, _
                                Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                FILE_RETRY, , , , , MENU_NO)
        Select Case sts
            Case False
            Case True           '入庫時は発生しない
            Case SYS_CANCEL
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                Zaitei_In_Proc = False
                GoTo Abort_Tran
            Case SYS_ERR
                Sendbuf = Text_Create_Proc()
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Zaitei_In_Proc = SYS_ERR    'システム異常発生
                
                GoTo Abort_Tran
        End Select
    End If

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        GoTo Abort_Tran
    End If
                                        '次の作業要求
    
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
        '   -------------------------------- エラーメッセージ作成
        Case Else
        '重要な要因なので未登録はシステム停止とする
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
        Exit Function
    End Select
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    If Sagyo_Send_Proc() Then
        Sendbuf = Text_Create_Proc()
        Exit Function
    End If
            
    Sendbuf = Text_Create_Proc()
    
    
    
    Zaitei_In_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function

Private Function Zaitei_Out_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『在訂（＋）指定時のチェック＆更新処理』
'
'-------------------------------------------------------
Dim i               As Integer
Dim Hinban          As String * 13
Dim Tanaban         As String * 8
Dim sts             As Integer

Dim QTY             As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    Zaitei_Out_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Tanaban        '棚番
                Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                
                If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                Else
                '------------------ 倉庫マスタ読込み
                    Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                            Zaitei_Out_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                            Exit Function
                    End Select
                    '------------------ 混載チェック
                    If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                        If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                            StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Zaitei_Out_Proc = False
                            Exit Function
                        End If
                    End If
                    '------------------ 棚マスタ読込み
                    Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                    Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                    Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                    Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Zaitei_Out_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                            Exit Function
                    End Select
            
                    '------------------ 禁止棚のチェック
                    If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                        Zaitei_Out_Proc = False
                        Exit Function
                    End If
            
                End If
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ 品目マスタ読込み
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Zaitei_Out_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                        End If
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Zaitei_Out_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                End Select
            
            Case LCD_Suryo          '数量（ここは無い）
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_Out_Proc = False
                    Exit Function
                End If
                
                QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                If QTY = 0 Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_Out_Proc = False
                    Exit Function
                End If
            
            
            Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_Out_Proc = False
                    Exit Function
                End If
                
                
                If Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD) = LCD_SUMI_Suryo Then
                    SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                Else
                    MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                End If
        
                If i = M_Gyo - 1 Then
                    If SUMI_QTY = 0 And MI_QTY = 0 Then
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "商品／未商品＝０", "数量入力ミス", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                        Zaitei_Out_Proc = False
                        Exit Function
                    End If
                End If
        
        End Select
    Next i
    '----------------------------------- データ更新処理開始 -----------
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Exit Function
    End If
                                        
    If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
        MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                
    Else
                
        MENU_NO = ""
    End If
                                        
                                        
                                        '出庫更新
    sts = Syuko_Update_Proc(RET_JGYOBU, _
                            RET_NAIGAI, _
                            Hinban, _
                            "", _
                            Tanaban, _
                            (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                            SUMI_QTY, _
                            MI_QTY, _
                            0, _
                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                            FILE_RETRY, , , , , , , MENU_NO)
    Select Case sts
        Case False
        
        Case True       '在庫不足時に発生
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "在庫数不足", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Zaitei_Out_Proc = False
            GoTo Abort_Tran
        Case SYS_CANCEL
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Zaitei_Out_Proc = False
            GoTo Abort_Tran
        Case SYS_ERR
            Sendbuf = Text_Create_Proc()
            Call Err_Send_Proc("システム異常発生", "", "", "", "")
            Zaitei_Out_Proc = SYS_ERR    'システム異常発生
            
            GoTo Abort_Tran
    End Select


End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        GoTo Abort_Tran
    End If
                                        '次の作業要求
    
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
        '   -------------------------------- エラーメッセージ作成
        Case Else
        '重要な要因なので未登録はシステム停止とする
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
        Exit Function
    End Select
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    If Sagyo_Send_Proc() Then
        Sendbuf = Text_Create_Proc()
        Exit Function
    End If
            
    Sendbuf = Text_Create_Proc()
    
    
    
    Zaitei_Out_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function

Private Function Ido_In_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『移動入庫指定時のチェック＆更新処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim From_Tanaban    As String * 8
Dim To_Tanaban      As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    Ido_In_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        To_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(To_Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Ido_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 混載チェック
                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Ido_In_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(To_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(To_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(To_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Ido_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Ido_In_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(To_Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            To_Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Ido_In_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_In_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
                                        'FROM 仮想棚番
            From_Tanaban = ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM & "01" & "01" & "01"
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, From_Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           'ここでは発生しない
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("在庫使用中", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_In_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("有効在庫無し", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                Ido_In_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = To_Tanaban       '棚番をセーブ
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品数量
                                                        
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '資材対応の事業部2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '資材対応の国内外2006.01.06
                                                        
                                                        
                                                        '商品化用倉庫区分をセーブ
            ID_KANRI_TBL(ING_No).GOODS_ON_F = StrConv(SOKOREC.GOODS_ON_F, vbUnicode)
            '数量付きの送信メッセージを作成する
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                            Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2))
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "08"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                            'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                            '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（商品／未商品数量）
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '数量（ここは無い）
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_In_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_In_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_In_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_In_Proc = False
                                Exit Function
                            
                            End If
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                            If MI_QTY > ID_KANRI_TBL(ING_No).Send_MI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_In_Proc = False
                                Exit Function
                            End If
                        End If
        
                        If i = M_Gyo - 1 Then       '最終行だったら
                            If SUMI_QTY = 0 And MI_QTY = 0 Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "商品／未商品＝０", "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                Ido_In_Proc = False
                                Exit Function
                            End If
                        End If
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                                        'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
            '2006.01.30
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
        
        
            sts = IDO_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM & "01" & "01" & "01"), _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , MENU_NO)

    
    
            Select Case sts
                Case False
        
                Case True       '在庫不足時に発生
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "在庫数不足", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                    Ido_In_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_In_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Ido_In_Proc = SYS_ERR    'システム異常発生
                    GoTo Abort_Tran
            End Select
    
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '出荷予定／在庫の予約解除
            sts = Data_Clear_Proc(0, Sendbuf)
            Select Case sts
                Case SYS_CANCEL
                    Call Err_Send_Proc("データ使用中", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
                                        
                                        
                                        '次の作業要求
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    Ido_In_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Ido_Out_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『移動出庫指定時のチェック＆更新処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim From_Tanaban    As String * 8
Dim To_Tanaban      As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    Ido_Out_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        From_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(From_Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Ido_Out_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(From_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(From_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(From_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Ido_Out_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Ido_Out_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(To_Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            From_Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Ido_Out_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_Out_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            
            '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
                                        'FROM 実棚番
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, From_Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           'ここでは発生しない
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("在庫使用中", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_Out_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("有効在庫無し", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                Ido_Out_Proc = False
                Exit Function
            End If
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = From_Tanaban     '棚番をセーブ
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品の数量
                                                            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '資材対応の事業部2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '資材対応の国内外2006.01.06
                                                            
                                                            
                                                            
                                                            '商品化用倉庫区分をセーブ
            ID_KANRI_TBL(ING_No).GOODS_ON_F = StrConv(SOKOREC.GOODS_ON_F, vbUnicode)
            '数量付きの送信メッセージを作成する
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                    Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                    Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "08"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                            'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                             '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                           '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（商品／未商品数量）
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '数量（ここは無い）
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_Out_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_Out_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_Out_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_Out_Proc = False
                                Exit Function
                            End If
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_Out_Proc = False
                                Exit Function
                            End If
                        End If
        
                        If i = M_Gyo - 1 Then       '最終行だったら
                            If SUMI_QTY = 0 And MI_QTY = 0 Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "商品／未商品＝０", "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                Ido_Out_Proc = False
                                Exit Function
                            End If
                        End If
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                                        'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
            '2006.01.30
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
            Else
                MENU_NO = ""
            End If
        
            sts = IDO_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM & "01" & "01" & "01"), _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , MENU_NO)

    
    
            Select Case sts
                Case False
        
                Case True       '在庫不足時に発生
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "在庫数不足", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                    Ido_Out_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_Out_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Ido_Out_Proc = SYS_ERR    'システム異常発生
                    GoTo Abort_Tran
            End Select
    
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
            '出荷予定／在庫の予約解除
            sts = Data_Clear_Proc(0, Sendbuf)
            Select Case sts
                Case SYS_CANCEL
                    Call Err_Send_Proc("データ使用中", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
                                        
                                        '次の作業要求
            
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    Ido_Out_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Tanahyoji_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『棚番表示処理』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13
Dim Tanaban     As String * 8
Dim ST_Tanaban  As String * 8


Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim Tanahyoji() As Tanahyoji_tag
Dim Tana_Cnt    As Integer


Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1


    Tanahyoji_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    '品目マスタ読み込み（標準棚番ＧＥＴ）
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        ST_Tanaban = (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Tanahyoji_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                End Select
                
                
                On Error Resume Next
                Kill (FullPath)             '送信用ファイル削除
                On Error GoTo 0
        
                Erase Tanahyoji
                Tana_Cnt = -1
        
                If Len(Trim(ST_Tanaban)) = 0 Then
                                            '標準棚番設定なし
                
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                
                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                Else
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            RET_JGYOBU, _
                                            RET_NAIGAI, _
                                            Hinban, _
                                            ST_Tanaban) Then

            
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Exit Function
                    End If
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                    Tanahyoji(Tana_Cnt).SUMI_QTY = SUMI_QTY
                    Tanahyoji(Tana_Cnt).MI_QTY = MI_QTY
                    
                    
                    
                End If
                
                Call UniCode_Conv(K6_ZAIKO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K6_ZAIKO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K6_ZAIKO.HIN_GAI, Hinban)
                Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                Call UniCode_Conv(K6_ZAIKO.Retu, "")
                Call UniCode_Conv(K6_ZAIKO.Ren, "")
                Call UniCode_Conv(K6_ZAIKO.Dan, "")
                    
                com = BtOpGetGreater
        
                SUMI_QTY = 0
                MI_QTY = 0
    
    
                Do
                    sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "在庫データ", 0)
                            Exit Function
                    End Select
                
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                        '事業部／国内外／品番ブレーク
                        Exit Do
                    End If
                    
                    If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) _
                        = ST_Tanaban Then
                        '標準棚番は対象外
                    Else
                        If Tana_Cnt = (-1) Then
                            Tana_Cnt = Tana_Cnt + 1
                            ReDim Tanahyoji(Tana_Cnt)
                            Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                Tanahyoji(Tana_Cnt).MI_QTY = 0
                            Else
                                Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            End If
                        
                        Else
                            For j = 0 To UBound(Tanahyoji)
                                If Tanahyoji(j).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                                    Exit For
                                End If
                            Next j
                        
                            If j <= UBound(Tanahyoji) Then
                            
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(j).SUMI_QTY = Tanahyoji(j).SUMI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                Else
                                    Tanahyoji(j).MI_QTY = Tanahyoji(j).MI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            
                            Else
                            
                                Tana_Cnt = Tana_Cnt + 1
                            
                                ReDim Preserve Tanahyoji(Tana_Cnt)
                                Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                                Else
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                    Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            End If
                        
                        End If
                    
                    End If
                    
                    com = BtOpGetNext
                    
                Loop
        
        
        
                FileNo = FreeFile           '送信用ファイルＯＰＥＮ
                Open FullPath For Binary As #FileNo
        
        
                SendFileRec.Title = "0"     'タイトル行
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Tana_Cnt + 1, "#0") & "件")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
        
        
                If Tana_Cnt > -1 Then
                                '集計テーブルよりデータ出力
                    For j = 0 To UBound(Tanahyoji)
                                
                        SendFileRec.Title = "1"
'                        Call UniCode_Conv(SendFileRec.LCD, Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
'                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
'                        SendFileRec.CRLF = vbCrLf
                        
                        If j = 0 Then
                            Call UniCode_Conv(SendFileRec.LCD, "*" & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        Else
                            Call UniCode_Conv(SendFileRec.LCD, " " & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        End If
                        SendFileRec.CRLF = vbCrLf
                                            
                        Put #FileNo, , SendFileRec
                                            
                                            
                        Call UniCode_Conv(SendFileRec.LCD, "  商：" & Format(Tanahyoji(j).SUMI_QTY, "#0") & "  未：" & Format(Tanahyoji(j).MI_QTY, "#0"))
                        SendFileRec.CRLF = vbCrLf
                                            
                                            
                        
                        Put #FileNo, , SendFileRec
                            
                    Next j
                End If
        
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '表示画面フラグ 参照画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '送信データファイル名
    Send_Text.fileName = B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------１～５行目
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX属性
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            '数値初期表示
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            
                                                            '初期カーソル位置
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '入力桁数
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            'メニュ―番号
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Tanahyoji_Proc = False
    

End Function
Private Function Tanahyoji_Kaso_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『棚番表示(仮想優先)処理』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13
Dim Tanaban     As String * 8
Dim ST_Tanaban  As String * 8


Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim Tanahyoji() As Tanahyoji_tag
Dim Tana_Cnt    As Integer


Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1


    Tanahyoji_Kaso_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    '品目マスタ読み込み（標準棚番ＧＥＴ）
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        ST_Tanaban = (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Tanahyoji_Kaso_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                End Select
                
                
                On Error Resume Next
                Kill (FullPath)             '送信用ファイル削除
                On Error GoTo 0
        
                Erase Tanahyoji
                Tana_Cnt = -1
        
                If Len(Trim(ST_Tanaban)) = 0 Then
                                            '標準棚番設定なし
                
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                
                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                Else
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            RET_JGYOBU, _
                                            RET_NAIGAI, _
                                            Hinban, _
                                            ST_Tanaban) Then

            
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Exit Function
                    End If
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                    Tanahyoji(Tana_Cnt).SUMI_QTY = SUMI_QTY
                    Tanahyoji(Tana_Cnt).MI_QTY = MI_QTY
                    
                    
                    
                End If
                
                Call UniCode_Conv(K6_ZAIKO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K6_ZAIKO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K6_ZAIKO.HIN_GAI, Hinban)
                Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                Call UniCode_Conv(K6_ZAIKO.Retu, "")
                Call UniCode_Conv(K6_ZAIKO.Ren, "")
                Call UniCode_Conv(K6_ZAIKO.Dan, "")
                    
                com = BtOpGetGreater
        
                SUMI_QTY = 0
                MI_QTY = 0
    
    
                Do
                    sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "在庫データ", 0)
                            Exit Function
                    End Select
                
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                        '事業部／国内外／品番ブレーク
                        Exit Do
                    End If
                    
                    If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) _
                        = ST_Tanaban Then
                        '標準棚番は対象外
                    Else
                        
                        '仮想倉庫を優先する
                        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_JITU)
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "倉庫ﾏｽﾀ", 0)
                                Exit Function
                        End Select
                        
                        
                        
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                        
                            If Tana_Cnt = (-1) Then
                                Tana_Cnt = Tana_Cnt + 1
                                ReDim Tanahyoji(Tana_Cnt)
                                Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                                Else
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                    Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            Else
                                For j = 0 To UBound(Tanahyoji)
                                    If Tanahyoji(j).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                                        Exit For
                                    End If
                                Next j
                            
                                If j <= UBound(Tanahyoji) Then
                                
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(j).SUMI_QTY = Tanahyoji(j).SUMI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Else
                                        Tanahyoji(j).MI_QTY = Tanahyoji(j).MI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                
                                Else
                                
                                    Tana_Cnt = Tana_Cnt + 1
                                
                                    ReDim Preserve Tanahyoji(Tana_Cnt)
                                    Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                        Tanahyoji(Tana_Cnt).MI_QTY = 0
                                    Else
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                        Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                End If
                            
                            End If
                        
                        End If
                    End If
                    
                    com = BtOpGetNext
                    
                Loop
        
        
                Call UniCode_Conv(K6_ZAIKO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K6_ZAIKO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K6_ZAIKO.HIN_GAI, Hinban)
                Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                Call UniCode_Conv(K6_ZAIKO.Retu, "")
                Call UniCode_Conv(K6_ZAIKO.Ren, "")
                Call UniCode_Conv(K6_ZAIKO.Dan, "")
                    
                com = BtOpGetGreater
        
                SUMI_QTY = 0
                MI_QTY = 0
    
    
                Do
                    sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "在庫データ", 0)
                            Exit Function
                    End Select
                
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                        '事業部／国内外／品番ブレーク
                        Exit Do
                    End If
                    
                    If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) _
                        = ST_Tanaban Then
                        '標準棚番は対象外
                    Else
                        
                        '仮想倉庫を優先する
                        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO)
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "倉庫ﾏｽﾀ", 0)
                                Exit Function
                        End Select
                        
                        
                        
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_JITU Then
                        
                            If Tana_Cnt = (-1) Then
                                Tana_Cnt = Tana_Cnt + 1
                                ReDim Tanahyoji(Tana_Cnt)
                                Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                                Else
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                    Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            Else
                                For j = 0 To UBound(Tanahyoji)
                                    If Tanahyoji(j).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                                        Exit For
                                    End If
                                Next j
                            
                                If j <= UBound(Tanahyoji) Then
                                
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(j).SUMI_QTY = Tanahyoji(j).SUMI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Else
                                        Tanahyoji(j).MI_QTY = Tanahyoji(j).MI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                
                                Else
                                
                                    Tana_Cnt = Tana_Cnt + 1
                                
                                    ReDim Preserve Tanahyoji(Tana_Cnt)
                                    Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                        Tanahyoji(Tana_Cnt).MI_QTY = 0
                                    Else
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                        Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                End If
                            
                            End If
                        
                        End If
                    End If
                    
                    com = BtOpGetNext
                    
                Loop
        
        
        
        
                FileNo = FreeFile           '送信用ファイルＯＰＥＮ
                Open FullPath For Binary As #FileNo
        
        
                SendFileRec.Title = "0"     'タイトル行
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Tana_Cnt + 1, "#0") & "件")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
        
        
                If Tana_Cnt > -1 Then
                                '集計テーブルよりデータ出力
                    For j = 0 To UBound(Tanahyoji)
                                
                        SendFileRec.Title = "1"
'                        Call UniCode_Conv(SendFileRec.LCD, Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
'                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
'                        SendFileRec.CRLF = vbCrLf
                        
                        If j = 0 Then
                            Call UniCode_Conv(SendFileRec.LCD, "*" & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        Else
                            Call UniCode_Conv(SendFileRec.LCD, " " & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        End If
                        SendFileRec.CRLF = vbCrLf
                                            
                        Put #FileNo, , SendFileRec
                                            
                                            
                        Call UniCode_Conv(SendFileRec.LCD, "  商：" & Format(Tanahyoji(j).SUMI_QTY, "#0") & "  未：" & Format(Tanahyoji(j).MI_QTY, "#0"))
                        SendFileRec.CRLF = vbCrLf
                                            
                                            
                        
                        Put #FileNo, , SendFileRec
                            
                    Next j
                End If
        
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '表示画面フラグ 参照画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '送信データファイル名
    Send_Text.fileName = B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------１～５行目
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX属性
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            '数値初期表示
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            
                                                            '初期カーソル位置
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '入力桁数
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            'メニュ―番号
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Tanahyoji_Kaso_Proc = False
    

End Function
Private Function Ave_Syuka_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『月平均出荷数表示処理』
'
'-------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim i               As Integer

Dim Hinban          As String * 13
Dim Tanaban         As String * 8

Dim AVE_SYUKA_ED    As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


    Ave_Syuka_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    '品目マスタ読み込み
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Ave_Syuka_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                End Select
                                '月平均出荷数読み込み
''''''''''''''''2006.01.06 資材に対応
'                Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, RET_NAIGAI)
''''''''''''''''2006.01.06 資材に対応
                Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, Hinban)
                sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "00000000")
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "月平均出荷数", 0)
                        Exit Function
                End Select
        
        End Select
    Next i
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 参照画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '送信データファイル名
    Send_Text.fileName = ""
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------１行目
                                                            'BOX属性
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
    Send_Text.Box_Type(0).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(0).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(0).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                            'メニュ―番号
    Send_Text.Box_Type(0).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
    
    '-----------------------------------------------２行目
                                                            'BOX属性
    Send_Text.Box_Type(1).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Hinban)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Hinban)
                                                            '数値初期表示
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(1).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(1).Max_Size = "13"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "13"
                                                            'メニュ―番号
    Send_Text.Box_Type(1).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '-----------------------------------------------３行目
                                                            'BOX属性
    Send_Text.Box_Type(2).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            
                                                            
    AVE_SYUKA_ED = Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#0")
    AVE_SYUKA_ED = "[" & Space(8 - Len(AVE_SYUKA_ED)) & AVE_SYUKA_ED & "]"
                                                             
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, AVE_SYUKA_ED)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, AVE_SYUKA_ED)
                                                            '数値初期表示
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(2).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(2).Max_Size = "08"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "08"
                                                            'メニュ―番号
    Send_Text.Box_Type(2).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
    '-----------------------------------------------４行目
                                                            'BOX属性
    Send_Text.Box_Type(3).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                            '数値初期表示
    Send_Text.Box_Type(3).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(3).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                            '入力桁数
    Send_Text.Box_Type(3).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                            'メニュ―番号
    Send_Text.Box_Type(3).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
    '-----------------------------------------------５行目
                                                            'BOX属性
    Send_Text.Box_Type(4).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                            '数値初期表示
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(4).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                            '入力桁数
    Send_Text.Box_Type(4).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                            'メニュ―番号
    Send_Text.Box_Type(4).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        

    Sendbuf = Text_Create_Proc()
    
    
    
    Ave_Syuka_Proc = False
    

End Function


Private Function Host_Zaiko_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『ホスト在庫（理論在庫）照会処理』
'
'-------------------------------------------------------
Dim sts             As Integer
Dim i               As Integer
Dim Hinban          As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1
    
    
    Host_Zaiko_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Host_Zaiko_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                
                End Select
                '------------------ 在庫集計データ読込み
                Call UniCode_Conv(K0_SUMZ.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_SUMZ.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_SUMZ.HIN_GAI, Hinban)
                sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "在庫集計データ", 0)
                        Exit Function
                End Select
            
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                 '送信データファイル名
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    
    '-----------------------------------------------１行目
                                                    'BOX属性
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                    '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                    '数値初期表示
    Send_Text.Box_Type(0).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                    
                                                    '初期カーソル位置
    Send_Text.Box_Type(0).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                    '入力桁数
    Send_Text.Box_Type(0).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                        
    Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
    
    
    '-----------------------------------------------２行目
                                                            'BOX属性
    Send_Text.Box_Type(1).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Hinban)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Hinban)
                                                            '数値初期表示
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
    Send_Text.Box_Type(1).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(1).Max_Size = "13"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "13"
                                                                                
    Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '-----------------------------------------------３行目
                                                            'BOX属性
    Send_Text.Box_Type(2).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "ホスト在庫:" & Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#0"))
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "ホスト在庫:" & Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#0"))
                                                            '数値初期表示
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(2).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(2).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
    '-----------------------------------------------４行目
                                                             'BOX属性
    Send_Text.Box_Type(3).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                            '数値初期表示
    Send_Text.Box_Type(3).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(3).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(3).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
       
    '-----------------------------------------------５行目
                                                             'BOX属性
    Send_Text.Box_Type(4).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                            '数値初期表示
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(4).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(4).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(4).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
    
    Sendbuf = Text_Create_Proc()
    
    
    
    Host_Zaiko_Proc = False
    

End Function
Private Function Tanaorosi_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『棚卸照会処理』
'
'-------------------------------------------------------
Dim sts             As Integer
Dim i               As Integer
Dim Hinban          As String
Dim Tanaban         As String
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim Sumi_ED         As String
Dim Mi_ED           As String
Dim Total_ED        As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

    Tanaorosi_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            
            Case LCD_Tanaban        '棚番
                
                Tanaban = Trim(Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1))
                
                
            Case LCD_Hinban         '品番
                
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
        
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Tanaorosi_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                
                End Select
        
        End Select
    Next i
    
    sts = Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            ID_KANRI_TBL(ING_No).JGYOBU, _
                            ID_KANRI_TBL(ING_No).NAIGAI, _
                            Hinban, _
                            Tanaban)
    If sts Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Exit Function
    End If
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                 '送信データファイル名
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------１行目
                                                        'BOX属性
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                        '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                        '数値初期表示
    Send_Text.Box_Type(0).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                        '初期カーソル位置
    Send_Text.Box_Type(0).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                        '入力桁数
    Send_Text.Box_Type(0).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
    Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
    '-----------------------------------------------２行目
                                                            'BOX属性
    Send_Text.Box_Type(1).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                            '数値初期表示
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(1).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(1).Max_Size = "13"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "13"
                                                                                
    Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '-----------------------------------------------３行目
                                                            'BOX属性
    Send_Text.Box_Type(2).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                            '数値初期表示
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(2).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(2).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
    '-----------------------------------------------４行目
                                                             'BOX属性
    Send_Text.Box_Type(3).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '表示内容
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "商 品/未 品/合 計")
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "商 品/未 品/合 計")
                                                            '数値初期表示
    Send_Text.Box_Type(3).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(3).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(3).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
       
    '-----------------------------------------------５行目
                                                             'BOX属性
    Send_Text.Box_Type(4).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                            '表示内容
    Sumi_ED = Space(5 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0") & "/"
    Mi_ED = Space(5 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0") & "/"
    Total_ED = Space(5 - Len(Format(SUMI_QTY + MI_QTY, "#0"))) & Format(SUMI_QTY + MI_QTY, "#0")
    
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Sumi_ED & Mi_ED & Total_ED)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Sumi_ED & Mi_ED & Total_ED)
                                                            '数値初期表示
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '初期カーソル位置
    Send_Text.Box_Type(4).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                            '入力桁数
    Send_Text.Box_Type(4).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(4).MENU = ""                     'メニュ―番号
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
    
    Sendbuf = Text_Create_Proc()
    
    
    
    Tanaorosi_Proc = False
    

End Function
Private Function St_Tanaban_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『標準棚番設定処理』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer
Dim Hinban      As String
Dim Tanaban     As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


    St_Tanaban_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            
            Case LCD_Tanaban        '棚番
                Tanaban = Trim(Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1))
                                '------------------ 倉庫マスタ読込み
                Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                        St_Tanaban_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                        Exit Function
                End Select
                    
                '------------------ 混載チェック    2006.01.06 品番ﾁｪｯｸ後に移動↓
'                If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
'                    If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
'                        StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
'                        Sendbuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                        St_Tanaban_Proc = False
'                        Exit Function
'                    End If
'                End If
                '------------------ 棚マスタ読込み
                Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        St_Tanaban_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                        Exit Function
                    End Select
   
            Case LCD_Hinban         '品番
                
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ 品目マスタ読込み
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        St_Tanaban_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                End Select
        
        
                '------------------ 混載チェック
                If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                    If StrConv(SOKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(SOKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Then
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        St_Tanaban_Proc = False
                        Exit Function
                    End If
                End If
        
        
        
        End Select
    Next i
    '----------------------------------- データ更新処理開始 -----------
    Call UniCode_Conv(K0_ITEM.JGYOBU, RET_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, RET_NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
    
    Do
        '------------------ 品目マスタ読込み
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            '   -------------------------------- エラーメッセージ作成
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
        
                St_Tanaban_Proc = False
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            '   -------------------------------- エラーメッセージ作成
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "他端末で使用中", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                St_Tanaban_Proc = False
                Exit Function
            Case Else
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                Exit Function
        End Select
    
    Loop
                                        '標準棚番設定
    Call UniCode_Conv(ITEMREC.ST_SOKO, Left(Tanaban, 2))
    Call UniCode_Conv(ITEMREC.ST_RETU, Mid(Tanaban, 3, 2))
    Call UniCode_Conv(ITEMREC.ST_REN, Mid(Tanaban, 5, 2))
    Call UniCode_Conv(ITEMREC.ST_DAN, Right(Tanaban, 2))
    Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
                                        
    Do
        '------------------ 品目マスタ更新
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            '   -------------------------------- エラーメッセージ作成
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "他端末で使用中", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                St_Tanaban_Proc = False
                Exit Function
            Case Else
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                Exit Function
        End Select
    Loop
                                        
                                        
                                        
                                        
                                        '次の作業要求
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
        '   -------------------------------- エラーメッセージ作成
        Case Else
        '重要な要因なので未登録はシステム停止とする
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
        Exit Function
    End Select
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    If Sagyo_Send_Proc() Then
        Sendbuf = Text_Create_Proc()
        Exit Function
    End If
            
    Sendbuf = Text_Create_Proc()
    
    
    
    St_Tanaban_Proc = False
    

End Function
Private Function Rireki_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『当日出庫推移表示処理』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag

Dim Data_Cnt    As Integer


    Rireki_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Rireki_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                
                End Select
                                    
                                    
                                    
                                    
                                    '空読みして件数カウント
                
                
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Format(Now, "YYYYMMDD"))
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                Data_Cnt = 0
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '事業部／内外／品番ブレーク？
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                Exit Do
                            End If
                        
                            '日付？
                            If StrConv(IDOREC.JITU_DT, vbUnicode) <> Format(Now, "YYYYMMDD") Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "在庫移動歴", 0)
                            Exit Function
                    End Select
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_ZAITEI_OUT Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                    
                    
                        Data_Cnt = Data_Cnt + 1
                    
                    End If
                    
                    com = BtOpGetPrev
                
                Loop
                
                
                On Error Resume Next
                Kill (FullPath)             '送信用ファイル削除
                On Error GoTo 0
        
                FileNo = FreeFile           '送信用ファイルＯＰＥＮ
                Open FullPath For Binary As #FileNo
        
                SendFileRec.Title = "0"     'タイトル行
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Data_Cnt, "#0") & "件")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                SendFileRec.Title = "0"     '品番
                Call UniCode_Conv(SendFileRec.LCD, Hinban)
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                    
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Format(Now, "YYYYMMDD"))
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '事業部／内外／品番ブレーク？
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                Exit Do
                            End If
                        
                            '日付？
                            If StrConv(IDOREC.JITU_DT, vbUnicode) <> Format(Now, "YYYYMMDD") Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "在庫移動歴", 0)
                            Exit Function
                    End Select
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_ZAITEI_OUT Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                    
                        '当日履歴出力
                        SendFileRec.Title = "1"
                        SUMI_QTY = CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
                        MI_QTY = CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
                        Call UniCode_Conv(SendFileRec.LCD, StrConv(IDOREC.RIRK_NAME, vbUnicode) & _
                                            Space(10 - Len(Format(SUMI_QTY + MI_QTY, "#0"))) & _
                                            Format(SUMI_QTY + MI_QTY, "#0"))
                        SendFileRec.CRLF = vbCrLf
                        Put #FileNo, , SendFileRec
                    End If
                    
                    com = BtOpGetPrev
                
                Loop
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '表示画面フラグ 参照画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '送信データファイル名
    Send_Text.fileName = B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------１～５行目
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX属性
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            
                                                            '数値初期表示
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            
                                                            '初期カーソル位置
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '入力桁数
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            'メニュ―番号
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Rireki_Proc = False
    

End Function
Private Function Suii_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『出荷推移表示処理』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim SUMI_QTY    As Long
Dim MI_QTY      As Long




Dim Start_YMD   As String * 8
Dim End_YMD     As String * 8
Dim Save_YMD    As String * 6
Dim SYUKA_QTY   As Long


Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag


Dim Data_Cnt    As Integer

    Suii_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 31 To 28 Step -1
        Start_YMD = Left(Format(DateAdd("m", -1, Now), "YYYYMMDD"), 6) & Format(i, "00")
        If IsDate(Left(Start_YMD, 4) & "/" & Mid(Start_YMD, 5, 2) & "/" & Right(Start_YMD, 2)) Then
            Exit For
        End If
    Next i

'    Start_YMD = Left(Format(DateAdd("m", -1, Now), "YYYYMMDD"), 6) & "31"
    
    End_YMD = Left(Format(DateAdd("m", -11, (Left(Start_YMD, 4) & "/" & Mid(Start_YMD, 5, 2) & "/" & Right(Start_YMD, 2))), "YYYYMMDD"), 6) & "01"

    Save_YMD = ""

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Suii_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                
                End Select
                                    
                                    '空読みしてデータ件数獲得
                
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Start_YMD)
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                Data_Cnt = 0
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '事業部／内外／品番ブレーク？
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    Data_Cnt = Data_Cnt + 1
                                End If
                                
                                Exit Do
                            
                            End If
                            '日付？
                            If StrConv(IDOREC.JITU_DT, vbUnicode) < End_YMD Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    Data_Cnt = Data_Cnt + 1
                                End If
                                
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            If Len(Trim(Save_YMD)) <> 0 Then
                                Data_Cnt = Data_Cnt + 1
                            End If
                            
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "在庫移動歴", 0)
                            Exit Function
                    End Select
                    
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                        
                        
                        If Len(Trim(Save_YMD)) = 0 Then
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                        End If
                        If Save_YMD <> Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6) Then
                            Data_Cnt = Data_Cnt + 1
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                        End If
                    
                    
                    
                    End If
                    
                    com = BtOpGetPrev
                Loop
                
                
                On Error Resume Next
                Kill (FullPath)             '送信用ファイル削除
                On Error GoTo 0
        
                FileNo = FreeFile           '送信用ファイルＯＰＥＮ
                Open FullPath For Binary As #FileNo
        
                SendFileRec.Title = "0"     'タイトル行
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Data_Cnt, "#0") & "件")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                SendFileRec.Title = "0"     '品番
                Call UniCode_Conv(SendFileRec.LCD, Hinban)
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                    
                    
                Save_YMD = ""
                    
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Start_YMD)
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '事業部／内外／品番ブレーク？
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    SendFileRec.Title = "1"
                                    Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                        Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                        Format(SYUKA_QTY, "#0"))
                                    SendFileRec.CRLF = vbCrLf
                                    Put #FileNo, , SendFileRec
                                End If
                                
                                Exit Do
                            
                            End If
                            '日付？
                            If StrConv(IDOREC.JITU_DT, vbUnicode) < End_YMD Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    SendFileRec.Title = "1"
                                    Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                        Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                        Format(SYUKA_QTY, "#0"))
                                    SendFileRec.CRLF = vbCrLf
                                    Put #FileNo, , SendFileRec
                                End If
                                
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            If Len(Trim(Save_YMD)) <> 0 Then
                                SendFileRec.Title = "1"
                                Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                    Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                    Format(SYUKA_QTY, "#0"))
                                SendFileRec.CRLF = vbCrLf
                                Put #FileNo, , SendFileRec
                            End If
                            
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "在庫移動歴", 0)
                            Exit Function
                    End Select
                    
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                        
                        
                        If Len(Trim(Save_YMD)) = 0 Then
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                            SYUKA_QTY = 0
                        End If
                        If Save_YMD <> Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6) Then
                                
                            SendFileRec.Title = "1"
                            Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                Format(SYUKA_QTY, "#0"))
                            SendFileRec.CRLF = vbCrLf
                            Put #FileNo, , SendFileRec
                            
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                            SYUKA_QTY = 0
                                                
                        End If
                    
                        SYUKA_QTY = SYUKA_QTY + (CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                    
                    
                    End If
                    
                    com = BtOpGetPrev
                Loop
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '送信メッセージを作成する
    Send_Text.sts = Sts_OK                                  'ステータス　OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '表示画面フラグ 参照画面
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '送信データファイル名
    Send_Text.fileName = B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------１～５行目
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX属性
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '表示内容
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            '数値初期表示
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            '初期カーソル位置
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '入力桁数
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            'メニュ―番号
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Suii_Proc = False
    

End Function
Private Function Hin_Shogo_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『品番別在庫照合処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MEMO            As String

Dim MENU_NO         As String

    Hin_Shogo_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then     '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Hin_Shogo_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 混載チェック
'                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
'                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
'                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
'                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
'                                    SendBuf = Text_Create_Proc()
'                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                    Hin_Shogo_Proc = False
'                                    Exit Function
'                                End If
'                            End If
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Hin_Shogo_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ 禁止棚のチェック
'                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
'
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚使用不可", "", "")
'
'                                SendBuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                Ido_In_Proc = False
'                                Exit Function
'                            End If
            
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Hin_Shogo_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Hin_Shogo_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            '   -------------------------------- 在庫数集計
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, Hinban) Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            '   -------------------------------- 送信テキスト作成
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '棚番をセーブ
            
            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '資材対応の事業部2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '資材対応の国内外2006.01.06
            
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品数量
                                                        
            '数量付きの送信メッセージを作成する
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "08"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                            'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                            '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（商品／未商品数量）
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '数量（ここは無い）
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                        End If
        
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                                        'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                        '品目マスタ読込み
            Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).S_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).S_NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).Hinban)
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "他で使用中", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Hin_Shogo_Proc = False
                        GoTo Abort_Tran
                    Case BtErrKeyNotFound
                    '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "品番未登録", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Hin_Shogo_Proc = False
                        GoTo Abort_Tran
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 0)
                        Hin_Shogo_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
        
            Loop
                                        '最終照合日付
            Call UniCode_Conv(ITEMREC.LAST_CHK_DT, Format(Date, "yyyymmdd"))
                                        '最終照合在庫数
            Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, Format(SUMI_QTY + MI_QTY, "00000000"))
                                        '品目マスタ書き込み
            Do
                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "他で使用中", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Hin_Shogo_Proc = False
                        GoTo Abort_Tran
            
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpUpdate, "品目マスタ", 0)
                        Hin_Shogo_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
        
            If (SUMI_QTY + MI_QTY) = (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                MEMO = B2_MEMO & StrConv(Format((SUMI_QTY + MI_QTY), "#0"), vbWide) & "[OK]"
            Else
                MEMO = B2_MEMO & StrConv(Format((ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_SUMI_QTY), "#0"), vbWide) & "[" & StrConv(Format(SUMI_QTY + MI_QTY, "#0"), vbWide) & "]"
            End If
            '2006.01.30
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
            Else
                MENU_NO = ""
            End If
                                
                                
            sts = IDOREKI_OUTPUT_PROC("", _
                                        "", _
                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                        ID_KANRI_TBL(ING_No).Hinban, _
                                        "", _
                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                        0, _
                                        0, _
                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        FILE_RETRY, _
                                        , _
                                        MEMO, , , , , MENU_NO)
            Select Case sts
                Case False      '正常終了
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Hin_Shogo_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
    
    
                
    
    
    
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        '次の作業要求
            
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    Hin_Shogo_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Tana_Hin_Shogo_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『棚番別品番別在庫照合処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MEMO            As String

Dim MENU_NO         As String * 2
    
    Tana_Hin_Shogo_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then     '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Tana_Hin_Shogo_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 混載チェック
                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Tana_Hin_Shogo_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Tana_Hin_Shogo_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ 禁止棚のチェック
'                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
'
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚使用不可", "", "")
'
'                                SendBuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                Ido_In_Proc = False
'                                Exit Function
'                            End If
            
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Tana_Hin_Shogo_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Tana_Hin_Shogo_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            '   -------------------------------- 在庫数集計
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    RET_JGYOBU, _
                                    RET_NAIGAI, _
                                    Hinban, _
                                    Tanaban) Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            '   -------------------------------- 送信テキスト作成
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '棚番をセーブ
            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '資材対応の事業部2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '資材対応の国内外2006.01.06
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品数量
                                                        
            '数量付きの送信メッセージを作成する
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "08"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                            'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                            '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（商品／未商品数量）
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '数量（ここは無い）
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Tana_Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Tana_Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                        End If
        
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                                        'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                        '品目マスタ読込み
'            Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'            Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'            Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).Hinban)
'            Do
'                sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                        Exit Do
'                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'                    '   -------------------------------- エラーメッセージ作成
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "他で使用中", "", "")
'
'                        SendBuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                        Hin_Shogo_Proc = False
'                        GoTo Abort_Tran
'                    Case BtErrKeyNotFound
'                    '   -------------------------------- エラーメッセージ作成
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "品番未登録", "", "")
'
'                        SendBuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                        Hin_Shogo_Proc = False
'                        GoTo Abort_Tran
'                    Case Else
'                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
'                        SendBuf = Text_Create_Proc()
'                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 0)
'                        Hin_Shogo_Proc = SYS_ERR
'                        GoTo Abort_Tran
'                End Select
'
'            Loop
'                                        '最終照合日付
'            Call UniCode_Conv(ITEMREC.LAST_CHK_DT, Format(Date, "yyyymmdd"))
'                                        '最終照合在庫数
'            Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, Format(SUMI_QTY + MI_QTY, "00000000"))
'                                        '品目マスタ書き込み
'            Do
'                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                        Exit Do
'                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "他で使用中", "", "")
'
'                        SendBuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                        Hin_Shogo_Proc = False
'                        GoTo Abort_Tran
'
'                    Case Else
'                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
'                        SendBuf = Text_Create_Proc()
'                        Call File_Error(sts, BtOpUpdate, "品目マスタ", 0)
'                        Hin_Shogo_Proc = SYS_ERR
'                        GoTo Abort_Tran
'                End Select
'            Loop
        
            If (SUMI_QTY + MI_QTY) = (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                MEMO = B8_MEMO & StrConv(Format((SUMI_QTY + MI_QTY), "#0"), vbWide) & "[OK]"
            Else
                MEMO = B8_MEMO & StrConv(Format((ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_SUMI_QTY), "#0"), vbWide) & "[" & StrConv(Format(SUMI_QTY + MI_QTY, "#0"), vbWide) & "]"
            End If
                                
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                                
                                
            sts = IDOREKI_OUTPUT_PROC(ID_KANRI_TBL(ING_No).Tanaban, _
                                        "", _
                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                        ID_KANRI_TBL(ING_No).Hinban, _
                                        "", _
                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                        0, _
                                        0, _
                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        FILE_RETRY, _
                                        , _
                                        MEMO, , , , , MENU_NO)
            Select Case sts
                Case False      '正常終了
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Tana_Hin_Shogo_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
    
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        '次の作業要求
            
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    Tana_Hin_Shogo_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function MTS_Dec_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『向け先宣言での出荷処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    MTS_Dec_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    MTS_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    MTS_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                MTS_Dec_Proc = False
                                Exit Function
                            End If
                        End If
                            
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            MTS_Dec_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                MTS_Dec_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            '------------------ 使用可能な出荷予定の予約を行い、出荷予定数を獲得する
            ID_NO = ""
            sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    Hinban, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE, _
                                    ID_KANRI_TBL(ING_No).SS_CODE, _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    Y_SYU_CNT, _
                                    ID_NO, _
                                    SYUKA_QTY, _
                                    DEN_NO, _
                                    KAN_KBN)
            Select Case sts
                Case False          '正常
                    If Y_SYU_CNT = 0 Then   '対象データなし
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "出荷予定無し", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        MTS_Dec_Proc = False
                        Exit Function
                    End If
                
                Case True
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "出荷予定使用中", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    MTS_Dec_Proc = False
                    Exit Function
            End Select
            '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           'ここでは発生しない
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("在庫使用中", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    MTS_Dec_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("有効在庫無し", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                MTS_Dec_Proc = False
                Exit Function
            End If
        
                
        
        
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '棚番をセーブ
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT      '該当伝票枚数
            ID_KANRI_TBL(ING_No).ID_NO = ID_NO              '伝票№
            ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO            '伝票№
            ID_KANRI_TBL(ING_No).YUKO_SUMI_QTY = SUMI_QTY   '使用可能な商品化済み在庫
            ID_KANRI_TBL(ING_No).YUKO_MI_QTY = MI_QTY       '使用可能な未商品在庫
        
        
            Select Case Y_SYU_CNT
                Case 1              '対象伝票が１枚
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                    ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                    '数量付きの送信メッセージを作成する
                    Send_Text.sts = Sts_OK                                      'ステータス　OK
                    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                    Send_Text.Display_Flg = Display_DEF                         '表示画面フラグ 通常入力画面
                    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                    Send_Text.End_Menu = Menu_Only                              '最終メニューフラグ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                    Send_Text.Menu_Suu = "05"                                   'メニュー項目数（05固定）
                    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                    Send_Text.fileName = ""                                     '送信データファイル名
                    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
                    Send_Text.Buzzer = Buzzer_DEF                               'ブザー音　標準
                    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                
                
                    '-----------------------------------------------１行目
                                                                                'BOX属性
                    Send_Text.Box_Type(0).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                    Send_Text.Box_Type(0).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                                                '初期カーソル位置
                    Send_Text.Box_Type(0).Start_Pos = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                    Send_Text.Box_Type(0).Max_Size = "00"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                    Send_Text.Box_Type(0).MENU = ""                             'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                    '-----------------------------------------------２行目
                                                                                'BOX属性
                    Send_Text.Box_Type(1).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                                            '数値初期表示
                    Send_Text.Box_Type(1).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(1).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                            '入力桁数
                    Send_Text.Box_Type(1).Max_Size = "08"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
                    Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                    '-----------------------------------------------３行目
                                                                            'BOX属性
                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                            '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                                            '数値初期表示
                    Send_Text.Box_Type(2).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(2).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                            '入力桁数
                    Send_Text.Box_Type(2).Max_Size = "13"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                    Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                    '-----------------------------------------------４行目
                                                                            'BOX属性
                    Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                                            
                    If SYUKA_QTY > (SUMI_QTY + MI_QTY) Then
                        SYUKA_QTY = SUMI_QTY + MI_QTY           '在庫数が少ない時は在庫数を送信
                    End If
                                                                            
                                                                            '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            
                                                                            
                                                                            '数値初期表示
                    Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                                            '入力桁数
                    Send_Text.Box_Type(3).Max_Size = "05"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
                    Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                    '-----------------------------------------------５行目
                                                                            'BOX属性
                    Send_Text.Box_Type(4).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                            '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                            '数値初期表示
                    Send_Text.Box_Type(4).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(4).Start_Pos = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                            '入力桁数
                    Send_Text.Box_Type(4).Max_Size = "00"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                    Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
                    Sendbuf = Text_Create_Proc()
                
                
                Case Else           '対象伝票が複数枚
            
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                    
                    '数量付きの送信メッセージを作成する
                    Send_Text.sts = Sts_OK                                      'ステータス　OK
                    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                    Send_Text.Display_Flg = Display_DEF                         '表示画面フラグ 通常入力画面
                    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                    Send_Text.End_Menu = Menu_Only                              '最終メニューフラグ
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                    Send_Text.Menu_Suu = "05"                                   'メニュー項目数（05固定）
                    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                    Send_Text.fileName = ""                                     '送信データファイル名
                    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
                    Send_Text.Buzzer = Buzzer_DEF                               'ブザー音　標準
                    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                
                
                    '-----------------------------------------------１行目
                                                                                'BOX属性
                    Send_Text.Box_Type(0).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                    Send_Text.Box_Type(0).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                                                '初期カーソル位置
                    Send_Text.Box_Type(0).Start_Pos = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                    Send_Text.Box_Type(0).Max_Size = "00"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                    Send_Text.Box_Type(0).MENU = ""                             'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                    '-----------------------------------------------２行目
                                                                                'BOX属性
                    Send_Text.Box_Type(1).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                                            '数値初期表示
                    Send_Text.Box_Type(1).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(1).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                            '入力桁数
                    Send_Text.Box_Type(1).Max_Size = "08"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
                    Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                    '-----------------------------------------------３行目
                                                                            'BOX属性
                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                            '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                                            '数値初期表示
                    Send_Text.Box_Type(2).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(2).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                            '入力桁数
                    Send_Text.Box_Type(2).Max_Size = "13"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                    Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                    '-----------------------------------------------４行目
                                                                            'BOX属性
                    Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                            '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_ID_No)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_ID_No)
                                                                            '数値初期表示
                    Send_Text.Box_Type(3).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(3).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                            '入力桁数
                    Send_Text.Box_Type(3).Max_Size = "08"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "08"
                                                                                
                    Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                    '-----------------------------------------------５行目
                                                                            'BOX属性
                    Send_Text.Box_Type(4).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                            '表示内容
                    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                            '数値初期表示
                    Send_Text.Box_Type(4).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                            '初期カーソル位置
                    Send_Text.Box_Type(4).Start_Pos = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                            '入力桁数
                    Send_Text.Box_Type(4).Max_Size = "00"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                    Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
                    Sendbuf = Text_Create_Proc()
            
            End Select
        
        Case Step_Sagyo2_RES        '２回目の受信（出荷数／伝票ＩＤ）
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Syuka      '出荷残数
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
                        '----------------------------------- データ更新処理開始 -----------
                                                            'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                    
                                    
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                                    
                                    '出庫処理
                        sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "他端末で使用中", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MTS_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
                                        'トランザクション終了
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                        '次の作業要求
                        
                        
                        '出荷予定／在庫の予約解除
                        sts = Data_Clear_Proc(0, Sendbuf)
                        Select Case sts
                            Case SYS_CANCEL
                                Call Err_Send_Proc("データ使用中", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                        
                        
                        Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                        Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                        Select Case sts
                            Case BtNoErr
                            '   -------------------------------- エラーメッセージ作成
                            Case Else
                            '重要な要因なので未登録はシステム停止とする
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                            Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
            
                        Sendbuf = Text_Create_Proc()
            
            
            
                    Case LCD_ID_No      '伝票ＩＤ
                
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN)
                        Select Case sts
                            Case False          '正常
                                If Y_SYU_CNT = 0 Then   '対象データなし
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "出荷予定無し", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    MTS_Dec_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                                Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "出荷予定使用中", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MTS_Dec_Proc = False
                                Exit Function
                        End Select
                
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO
                
                        '------------------ 確定した出荷予定の予定数を送信する
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                        '数量付きの送信メッセージを作成する
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
        
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                    '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                                                    Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                                                    Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2))
                                                                            '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                            '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                            '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "08"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                            'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                            '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban)
                                                                            '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                            '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                            '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                            'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                                            
                        If SYUKA_QTY > (ID_KANRI_TBL(ING_No).YUKO_SUMI_QTY + ID_KANRI_TBL(ING_No).YUKO_MI_QTY) Then
                            SYUKA_QTY = ID_KANRI_TBL(ING_No).YUKO_SUMI_QTY + ID_KANRI_TBL(ING_No).YUKO_MI_QTY
                        End If
                                                                            
                                                                            '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            '数値初期表示
                        Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                                                                            '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                                            '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------５行目
                                                                            'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                            '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                            '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                            '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                            '入力桁数
                        Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        
                        Sendbuf = Text_Create_Proc()
                
                End Select
            Next i
        Case Step_Sagyo3_RES        '３回目の受信（出荷数）
            For i = 0 To M_Gyo - 1
            
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Syuka      '出荷残数
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
                        '----------------------------------- データ更新処理開始 -----------
                                                            'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                                    '出庫処理
                        sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "他端末で使用中", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MTS_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
                                        'トランザクション終了
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                        '次の作業要求
                        
                        
                        Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                        Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                        Select Case sts
                            Case BtNoErr
                            '   -------------------------------- エラーメッセージ作成
                            Case Else
                            '重要な要因なので未登録はシステム停止とする
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                            Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
            
                        Sendbuf = Text_Create_Proc()
                End Select
            Next i
    End Select

    MTS_Dec_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function
Private Function DEN_ID_Dec_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『伝票ＩＤでの出荷処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8


Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim KAN_KBN         As String * 1

Dim MENU_NO         As String * 2

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


    DEN_ID_Dec_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（伝票ＩＤ）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No      '伝票ＩＤ
    
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            RET_JGYOBU = Left(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), 1)
                            ID_NO = Right(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), 12)
                        Else
                            RET_JGYOBU = ID_KANRI_TBL(ING_No).JGYOBU
'''                            If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                                ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                            Else
                                ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                            End If
                        End If
                        '------------------ 使用可能な出荷予定の予約を行い、出荷予定数を獲得する
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                                RET_JGYOBU, _
                                                RET_NAIGAI, _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN)
                        Select Case sts
                            Case False          '正常
                                If Y_SYU_CNT = 0 Then   '対象データなし
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定無し", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    DEN_ID_Dec_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                                Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定使用中", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                DEN_ID_Dec_Proc = False
                                Exit Function
                        End Select
                
                
                        If KAN_KBN <> KAN_KBN_UN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出庫処理済み", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                                                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '資材対応の事業部2006.01.06
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '資材対応の国内外2006.01.06
                           
                        
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        
                        '-----------------------------------------------１行目
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "09"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "09"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                                '入力桁数
                         Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        Case Step_Sagyo2_RES        '２回目の受信（棚番／品番／数量）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban    '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    DEN_ID_Dec_Proc = False
                                    Exit Function
                    
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    DEN_ID_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                DEN_ID_Dec_Proc = False
                                Exit Function
                            End If
                        End If
                            
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                        If ID_KANRI_TBL(ING_No).JGYOBU = ID_KANRI_TBL(ING_No).S_JGYOBU And _
                            ID_KANRI_TBL(ING_No).NAIGAI = ID_KANRI_TBL(ING_No).S_NAIGAI Then
                        
                            sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        
                        Else
                        
                            sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, ID_KANRI_TBL(ING_No).S_NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        
                        End If
                        
                        
                        
                        
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            DEN_ID_Dec_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                DEN_ID_Dec_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
                    
                    
                    
                        If Hinban <> ID_KANRI_TBL(ING_No).Hinban Then
                        
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                            DEN_ID_Dec_Proc = False
                            Exit Function
                        
                        End If
                    
                    
                    Case LCD_Syuka      '出荷残数
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                        
                        
                        
                        '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
                        Select Case sts
                            Case False
                            Case True           'ここでは発生しない
                                Exit Function
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc("在庫使用中", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                DEN_ID_Dec_Proc = False
                                Exit Function
                        End Select
                    
                        If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                            Call Err_Send_Proc("有効在庫無し", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                        
                        
                        
                        If SYUKA_QTY > (SUMI_QTY + MI_QTY) Then
                            Call Err_Send_Proc("出荷数不足", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                        
                        '----------------------------------- データ更新処理開始 -----------
                                                            'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                    
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                                    '出庫処理
                        sts = Syuko_Update_Proc(RET_JGYOBU, _
                                    RET_NAIGAI, _
                                    Hinban, _
                                    "", _
                                    Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, _
                                    MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "他端末で使用中", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                DEN_ID_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
                                        'トランザクション終了
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                        
                                        
                        '出荷予定／在庫の予約解除
                        sts = Data_Clear_Proc(0, Sendbuf)
                        Select Case sts
                            Case SYS_CANCEL
                                Call Err_Send_Proc("データ使用中", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                                        
                                        
                                        '次の作業要求
                                        
                        Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                        Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                        Select Case sts
                            Case BtNoErr
                            '   -------------------------------- エラーメッセージ作成
                            Case Else
                            '重要な要因なので未登録はシステム停止とする
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                            Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
            
                        Sendbuf = Text_Create_Proc()
                End Select
            Next i
    End Select
    DEN_ID_Dec_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If



End Function
Private Function SYUKO_HYO_Dec_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『出庫表での出荷処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    SYUKO_HYO_Dec_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（出庫表№）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_SYUKO_HYO_No   '出庫表№
    
                        
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出庫表使用不可", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                        
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "000000000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        '------------------ 使用可能な出荷予定の予約を行い、出荷予定数を獲得する
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                "", _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN)
                        Select Case sts
                            Case False          '正常
                                If Y_SYU_CNT = 0 Then   '対象データなし
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定無し", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    SYUKO_HYO_Dec_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                                Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定使用中", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                        End Select
                
                        If KAN_KBN <> KAN_KBN_UN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出庫処理済み", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "09"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "09"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                                '入力桁数
                         Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        Case Step_Sagyo2_RES        '２回目の受信（棚番／品番／数量）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban    '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    SYUKO_HYO_Dec_Proc = False
                                    Exit Function
                    
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    SYUKO_HYO_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                            End If
                        End If
                            
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            SYUKO_HYO_Dec_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
                    
                    
                    
                        If Hinban <> ID_KANRI_TBL(ING_No).Hinban Then
                        
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        
                        End If
                    
                    
                    Case LCD_Syuka      '出荷残数
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                        
                        '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
                        Select Case sts
                            Case False
                            Case True           'ここでは発生しない
                                Exit Function
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc("在庫使用中", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                        End Select
                    
                        If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                            Call Err_Send_Proc("有効在庫無し", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                        
                        
                        If SYUKA_QTY > (SUMI_QTY + MI_QTY) Then
                            Call Err_Send_Proc("出荷数不足", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                        
                        '----------------------------------- データ更新処理開始 -----------
                                                            'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                    
                                    
                                    
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                                    
                                    
                                    
                                    '出庫処理
                        sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    Hinban, _
                                    "", _
                                    Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "他端末で使用中", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SYUKO_HYO_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
                                        'トランザクション終了
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                        
                                        
                        '出荷予定／在庫の予約解除
                        sts = Data_Clear_Proc(0, Sendbuf)
                        Select Case sts
                            Case SYS_CANCEL
                                Call Err_Send_Proc("データ使用中", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                                        
                                        '次の作業要求
                                        
                        Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                        Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                        Select Case sts
                            Case BtNoErr
                            '   -------------------------------- エラーメッセージ作成
                            Case Else
                           '重要な要因なので未登録はシステム停止とする
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                            Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
            
                        Sendbuf = Text_Create_Proc()
                End Select
            Next i
    End Select
    SYUKO_HYO_Dec_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If



End Function
Private Function Out_Plan_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『計画外（データ無し）の出荷処理』
'
'-------------------------------------------------------
Dim i               As Integer
Dim Hinban          As String * 13
Dim Tanaban         As String * 8
Dim sts             As Integer

Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2
    
    Out_Plan_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Tanaban        '棚番
                Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                
                If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                Else
                '------------------ 倉庫マスタ読込み
                    Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                            Out_Plan_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                            Exit Function
                    End Select
                    '------------------ 混載チェック
                    If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                        If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                            StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                             
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                             
                             Out_Plan_Proc = False
                            Exit Function
                        End If
                    End If
                    '------------------ 棚マスタ読込み
                    Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                    Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                    Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                    Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- エラーメッセージ作成
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Out_Plan_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                            Exit Function
                    End Select
            
                    '------------------ 禁止棚のチェック
                    If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                        Out_Plan_Proc = False
                        Exit Function
                    End If
            
                End If
            Case LCD_Hinban         '品番
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ 品目マスタ読込み
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Out_Plan_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                        End If
                    Case BtErrKeyNotFound
                        '   -------------------------------- エラーメッセージ作成
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Out_Plan_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                        Exit Function
                End Select
            
            Case LCD_Suryo          '数量
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Out_Plan_Proc = False
                    Exit Function
                End If
                
                SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                If SYUKA_QTY = 0 Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Out_Plan_Proc = False
                    Exit Function
                End If
            
            
        
        End Select
    Next i
    '----------------------------------- データ更新処理開始 -----------
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Exit Function
    End If
                                        
    If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
        MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                
    Else
                
        MENU_NO = ""
    End If
                                        
                                        '出庫更新
    sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                            ID_KANRI_TBL(ING_No).NAIGAI, _
                            Hinban, _
                            "", _
                            Tanaban, _
                            (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                            SUMI_QTY, _
                            MI_QTY, _
                            SYUKA_QTY, _
                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                            FILE_RETRY, _
                            "", _
                            ID_KANRI_TBL(ING_No).CYU_KBN, _
                            ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                            Format(Now, "YYYYMMDD"), _
                            "", _
                            "", MENU_NO)

    Select Case sts
        Case False
        
        Case True       '在庫不足時に発生
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "在庫数不足", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Out_Plan_Proc = False
            GoTo Abort_Tran
        Case SYS_CANCEL
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Out_Plan_Proc = False
            GoTo Abort_Tran
        Case SYS_ERR
            Call Err_Send_Proc("システム異常発生", "", "", "", "")
            Sendbuf = Text_Create_Proc()
            Out_Plan_Proc = SYS_ERR    'システム異常発生
            
            GoTo Abort_Tran
    End Select


End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        GoTo Abort_Tran
    End If
                                        '次の作業要求
    
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    Select Case sts
        Case BtNoErr
        '   -------------------------------- エラーメッセージ作成
        Case Else
        '重要な要因なので未登録はシステム停止とする
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
        Exit Function
    End Select
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
    If Sagyo_Send_Proc() Then
        Sendbuf = Text_Create_Proc()
        Exit Function
    End If
            
    Sendbuf = Text_Create_Proc()
    
    
    
    Out_Plan_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Inspe_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『検品処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim JITU_QTY        As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim NAIGAI          As String * 1
Dim HIN_NO          As String * 13


Dim KAN_FLG         As String * 1

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim MENU_NO         As String * 2

    Inspe_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（伝票ＩＤ）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No      '伝票ＩＤ
    
    
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "検品作業不可", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
    
    
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        '------------------ 使用可能な出荷予定の予約を行い、出荷予定数を獲得する
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                NAIGAI, _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN, _
                                                JITU_QTY)
                        Select Case sts
                            Case False          '正常
                                If Y_SYU_CNT = 0 Then   '対象データなし
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定無し", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定使用中", "", "")
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc = False
                                Exit Function
                        End Select
                
                        '------------------ 向け先のチェック
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "向け先エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ 注文区分のチェック
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "注文区分ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ 出庫完了のチェック
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "作業未完了", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc = False
                                Exit Function
                            End If
                        End If
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT
                        ID_KANRI_TBL(ING_No).SYUKA_QTY = JITU_QTY
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Hinban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""                    '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '入力桁数
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '２回目の受信（品番）
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Hinban     '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc = False
                                Exit Function
                    
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                
                        End Select
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "品番エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                
                
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "出荷数：" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "出荷数：" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""                    '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '入力桁数
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
                
                
                
                End Select
            
            Next i
        Case Step_Sagyo3_RES        '３回目の受信（Any Key）
            '----------------------------------- データ更新処理開始 -----------
                                            'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                            '出荷予定の読み込み
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '事業部
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_KANRI_TBL(ING_No).ID_NO)   'ID№
    
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "出荷予定不明", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "出荷予定使用中", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc = False
                        GoTo Abort_Tran
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
        
            Loop
    
    
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                            
                                            '出荷予定書込み
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "出荷予定使用中", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc = False
                        GoTo Abort_Tran
                
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定", 0)
                        Inspe_Proc = SYS_ERR
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
            Loop
                                        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                                        
            '2004.07.16 ↓
            sts = IDOREKI_OUTPUT_PROC("", _
                                        "", _
                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                        ID_KANRI_TBL(ING_No).Hinban, _
                                        "", _
                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                        0, _
                                        0, _
                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        FILE_RETRY, _
                                        , _
                                        "", , , , , MENU_NO, _
                                        ID_KANRI_TBL(ING_No).MTS_CODE, _
                                        ID_KANRI_TBL(ING_No).SS_CODE, _
                                        ID_KANRI_TBL(ING_No).ID_NO)
            Select Case sts
                Case False      '正常終了
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Inspe_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
            '2004.07.16 ↑
                                        
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
            '次の作業要求
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                    Exit Function
            End Select
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        
            Sendbuf = Text_Create_Proc()
    
    
    End Select

    Inspe_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function
Private Function Inspe_Proc_MTS(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『検品処理（ＭＴＳ読み込みあり）』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim JITU_QTY        As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim NAIGAI          As String * 1
Dim HIN_NO          As String * 13


Dim KAN_FLG         As String * 1

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    Inspe_Proc_MTS = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（向け先）
        
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_MTS    '向け先
                                
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) < 16 Then
                                    '向け先（得意先）のみで向け先マスタ読み込み
                            Call UniCode_Conv(K2_MTS.MUKE_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
                            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
                            Select Case sts
                                Case BtNoErr
                                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                                    
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "出荷先エラー", "", "")
                    
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                        Inspe_Proc_MTS = False
                                        Exit Function
                                    
                                    End If
                                
                                Case BtErrKeyNotFound
                                
                                    Call UniCode_Conv(K3_MTS.SS_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
                                                        
                                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "出荷先エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                            Inspe_Proc_MTS = False
                                            Exit Function
                                        
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ", 0)
                                            Exit Function
                                    End Select
                        
                            End Select
                        
                            MTS_CODE = StrConv(MTSREC.MUKE_CODE, vbUnicode)
                            SS_CODE = StrConv(MTSREC.SS_CODE, vbUnicode)
                        
                        
                        Else
                            MTS_CODE = Left(ID_KANRI_TBL(ING_No).Recv_text(i), 8)
                            SS_CODE = Right(ID_KANRI_TBL(ING_No).Recv_text(i), 8)
                        
                                                '向け先マスタ読み込み
                            Call UniCode_Conv(K0_MTS.MUKE_CODE, MTS_CODE)
                            Call UniCode_Conv(K0_MTS.SS_CODE, SS_CODE)
                         
                            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, MTS_CODE & SS_CODE, "出荷先エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Inspe_Proc_MTS = False
                                    Exit Function
                            
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ", 0)
                                    Exit Function
                            End Select
                        
                        
                        End If
                         
                         
                         
                         
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, MTS_CODE & SS_CODE)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, MTS_CODE & SS_CODE)
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_ID_No)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_ID_No)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""                    '数値は５桁固定
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '入力桁数
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        
        
        Case Step_Sagyo2_RES        '２回目の受信（伝票ＩＤ）
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), _
                            Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - _
                            CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_ID_No      '伝票ＩＤ
    
    
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "検品作業不可", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
    
    
    
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        '------------------ 使用可能な出荷予定の予約を行い、出荷予定数を獲得する
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                NAIGAI, _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN, _
                                                JITU_QTY)
                        Select Case sts
                            Case False          '正常
                                If Y_SYU_CNT = 0 Then   '対象データなし
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定無し", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_MTS = False
                                    Exit Function
                                End If
                
                            Case True
                            Case SYS_ERR
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "出荷予定使用中", "", "")
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc_MTS = False
                                Exit Function
                        End Select
                
                        '------------------ 向け先のチェック
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "向け先エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ 注文区分のチェック
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "注文区分ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ 出庫完了のチェック
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "作業未完了", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_MTS = False
                                Exit Function
                            End If
                        End If
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT
                        ID_KANRI_TBL(ING_No).SYUKA_QTY = JITU_QTY
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                                                                                'BOX属性
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                                                                                '数値初期表示
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(1).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(1).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo3_RES        '３回目の受信（品番）
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Hinban     '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_MTS = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                
                        End Select
                        
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "品番エラー", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                
                
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
                        
                        
                        '-----------------------------------------------ヘッダー
                        Send_Text.sts = Sts_OK                                  'ステータス　OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '送信データファイル名
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------１行目
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '数値初期表示
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------２行目
                                                                                'BOX属性
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------３行目
                                                                                'BOX属性
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "伝票ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                '数値初期表示
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(2).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(2).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------４行目
                                                                                'BOX属性
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Hinban)
                                                                                '数値初期表示
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------５行目
                                                                                'BOX属性
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '表示内容
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "出荷数：" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "出荷数：" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
                                                                                '数値初期表示
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '初期カーソル位置
                        Send_Text.Box_Type(4).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '入力桁数
                        Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
                
                
                
                End Select
            
            Next i
        Case Step_Sagyo4_RES        '４回目の受信（Any Key）
            '----------------------------------- データ更新処理開始 -----------
                                            'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                            '出荷予定の読み込み
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '事業部
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_KANRI_TBL(ING_No).ID_NO)   'ID№
    
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "出荷予定不明", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc_MTS = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "出荷予定使用中", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc_MTS = False
                        GoTo Abort_Tran
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
        
            Loop
    
    
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                            
                                            '出荷予定書込み
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "出荷予定使用中", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc_MTS = False
                        GoTo Abort_Tran
                
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定", 0)
                        Inspe_Proc_MTS = SYS_ERR
                        Call Err_Send_Proc("システム異常発生", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
            Loop
                                        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                                        
                                        
                                        
            '2004.07.16 ↓
            sts = IDOREKI_OUTPUT_PROC("", _
                                        "", _
                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                        ID_KANRI_TBL(ING_No).Hinban, _
                                        "", _
                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                        0, _
                                        0, _
                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        FILE_RETRY, _
                                        , _
                                        "", , , , , MENU_NO, _
                                        ID_KANRI_TBL(ING_No).MTS_CODE, _
                                        ID_KANRI_TBL(ING_No).SS_CODE, _
                                        ID_KANRI_TBL(ING_No).ID_NO)
            Select Case sts
                Case False      '正常終了
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Inspe_Proc_MTS = SYS_ERR
                    GoTo Abort_Tran
            End Select
            '2004.07.16 ↑
                                        
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
            '次の作業要求
                Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                    Exit Function
                End Select
                
                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                
                '検品確認終了時は、伝票ＩＤ要求に戻す為の特殊処理
                
                
                '-----------------------------------------------ヘッダー
                Send_Text.sts = Sts_OK                                  'ステータス　OK
                ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
        
                Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
                ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
        
                Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
        
                Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
                ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
        
                Send_Text.fileName = ""                                 '送信データファイル名
                ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
        
                Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
                ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                
                '-----------------------------------------------１行目
                                                                'BOX属性
                Send_Text.Box_Type(0).Box_Type = TYPE_REF
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                '表示内容
                Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                        '数値初期表示
                Send_Text.Box_Type(0).INIT = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                        '初期カーソル位置
                Send_Text.Box_Type(0).Start_Pos = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                        '入力桁数
                Send_Text.Box_Type(0).Max_Size = "00"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                        
                Send_Text.Box_Type(0).MENU = ""                         'メニュ―番号
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                '-----------------------------------------------２行目
                                                                        'BOX属性
                Send_Text.Box_Type(1).Box_Type = TYPE_REF
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                        '表示内容
                Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                                                                        '数値初期表示
                Send_Text.Box_Type(1).INIT = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                        '初期カーソル位置
                Send_Text.Box_Type(1).Start_Pos = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                                        '入力桁数
                Send_Text.Box_Type(1).Max_Size = "00"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                        
                Send_Text.Box_Type(1).MENU = ""                         'メニュ―番号
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                '-----------------------------------------------３行目
                                                                        'BOX属性
                Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                        '表示内容
                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_ID_No)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_ID_No)
                                                                        '数値初期表示
                Send_Text.Box_Type(2).INIT = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                        '初期カーソル位置
                Send_Text.Box_Type(2).Start_Pos = "01"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                        '入力桁数
                Send_Text.Box_Type(2).Max_Size = "12"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "12"
                                                                        
                Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                '-----------------------------------------------４行目
                                                                        'BOX属性
                Send_Text.Box_Type(3).Box_Type = TYPE_REF
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                        '表示内容
                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                        '数値初期表示
                Send_Text.Box_Type(3).INIT = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                        '初期カーソル位置
                Send_Text.Box_Type(3).Start_Pos = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                        '入力桁数
                Send_Text.Box_Type(3).Max_Size = "00"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                        
                Send_Text.Box_Type(3).MENU = ""                         'メニュ―番号
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                '-----------------------------------------------４行目
                                                                        'BOX属性
                Send_Text.Box_Type(4).Box_Type = TYPE_REF
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                        '表示内容
                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                        '数値初期表示
                Send_Text.Box_Type(4).INIT = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                        '初期カーソル位置
                Send_Text.Box_Type(4).Start_Pos = ""                    '数値は５桁固定
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                        '入力桁数
                 Send_Text.Box_Type(4).Max_Size = "00"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                        
                Send_Text.Box_Type(4).MENU = ""                         'メニュ―番号
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                Sendbuf = Text_Create_Proc()


'                If Sagyo_Send_Proc() Then
'                    Sendbuf = Text_Create_Proc()
'                    Exit Function
'                End If
            
                Sendbuf = Text_Create_Proc()
    
    
    End Select

    Inspe_Proc_MTS = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function
Private Function Zaiko_Reserve_Proc(ID_NO As Integer, FROM_LOCATION As String, JGYOBU As String, NAIGAI As String, Hinban As String, SUMI_QTY As Long, MI_QTY As Long) As Integer
'-------------------------------------------------------
'
'   『在庫データの使用予約』
'
'-------------------------------------------------------
Dim sts             As Integer

    Zaiko_Reserve_Proc = True
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Zaiko_Reserve_Proc = SYS_ERR
        Exit Function
    End If

    sts = Zaiko_Lock_Proc(FROM_LOCATION, JGYOBU, NAIGAI, Hinban, Format(ID_NO, "000"), SUMI_QTY, MI_QTY, FILE_RETRY)
    If sts Then
        Zaiko_Reserve_Proc = sts
        GoTo Abort_Tran
    End If
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Zaiko_Reserve_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    Zaiko_Reserve_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Y_Syuka_Chek_Proc(Mode As String, _
                                        JGYOBU As String, _
                                        NAIGAI As String, _
                                        Hinban As String, _
                                        MTS_CODE As String, _
                                        SS_CODE As String, _
                                        CYU_KBN As String, _
                                        Y_SYU_CNT As Integer, _
                                        ID_NO As String, _
                                        SYUKA_QTY As Long, _
                                        DEN_NO As String, _
                                        KAN_KBN As String, _
                                        Optional JITU_QTY) As Integer
'-------------------------------------------------------
'
'   『単一出荷予定の使用予約』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer


Dim ans         As Integer

Dim RETRY_CNT   As Integer

Dim WK_ID_NO    As String * 12
Dim WK_DEN_NO   As String * 6




    Y_Syuka_Chek_Proc = True
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Y_Syuka_Chek_Proc = SYS_ERR
        Exit Function
    End If


    WK_ID_NO = ""
    WK_DEN_NO = ""

    Y_SYU_CNT = 0
    
    If Len(Trim(ID_NO)) <> 0 Then
        '伝票ＩＤ指定での処理
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)              '事業部
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_NO)            'ID№
    
        RETRY_CNT = 0
    
        Do
            DoEvents
            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    'データなし
                    Y_Syuka_Chek_Proc = False
                    GoTo Abort_Tran
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                        Y_Syuka_Chek_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                    Y_Syuka_Chek_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
                                                
                                                
        If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) = 0 And _
            Len(Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode))) = 0 Then
        Else
            If StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> Format(ID_KANRI_TBL(ING_No).ID, "000") Or _
                Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) <> StrConv(App.EXEName, vbUpperCase) Then
                                '他で使用中
                Y_Syuka_Chek_Proc = SYS_CANCEL
                GoTo Abort_Tran
            End If
        End If
            
        Call UniCode_Conv(Y_SYUREC.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
        Call UniCode_Conv(Y_SYUREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
                                            
        RETRY_CNT = 0
                                            
                                            '出荷予定書込み
        Do
            DoEvents
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                        Y_Syuka_Chek_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpUpdate, "出荷予定", 0)
                    Y_Syuka_Chek_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
                                
                                '1件のみ伝票№＆出荷数KEEP
        SYUKA_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
        
        WK_ID_NO = StrConv(Y_SYUREC.ID_NO, vbUnicode)
        WK_DEN_NO = Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6)
        Y_SYU_CNT = 1
        
        NAIGAI = StrConv(Y_SYUREC.NAIGAI, vbUnicode)
        Hinban = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
        MTS_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)
        SS_CODE = StrConv(Y_SYUREC.SS_CODE, vbUnicode)
        CYU_KBN = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
        KAN_KBN = StrConv(Y_SYUREC.KAN_KBN, vbUnicode)
        JITU_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
'        SYUKA_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
    
    
    
    Else
        '注文区分／向け先／品番での処理
        Call UniCode_Conv(K3_Y_SYU.JGYOBU, JGYOBU)              '事業部
        Call UniCode_Conv(K3_Y_SYU.KEY_CYU_KBN, CYU_KBN)        '注文区分
        Call UniCode_Conv(K3_Y_SYU.KEY_MUKE_CODE, MTS_CODE)     '得意先コード
        Call UniCode_Conv(K3_Y_SYU.KEY_SS_CODE, SS_CODE)        '得意先コード
        Call UniCode_Conv(K3_Y_SYU.NAIGAI, NAIGAI)              '国内外
        Call UniCode_Conv(K3_Y_SYU.KEY_HIN_NO, Hinban)          '品番
        Call UniCode_Conv(K3_Y_SYU.KEY_ID_NO, "")               'ID№
    
        com = BtOpGetGreaterEqual
    
    
        Do
            RETRY_CNT = 0
            Do
                DoEvents
                sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                Select Case sts
                    Case BtNoErr
                        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> JGYOBU Or _
                            StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> CYU_KBN Or _
                            Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) <> Trim(MTS_CODE) Or _
                            Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) <> Trim(SS_CODE) Or _
                            StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                            Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) <> Trim(Hinban) Then

                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "出荷予定", 0)
                                Y_Syuka_Chek_Proc = SYS_ERR
                                GoTo Abort_Tran
                            End If
                        
                            sts = BtErrEOF
                    
                        End If
                                        
                    
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > FILE_RETRY Then
                            Y_Syuka_Chek_Proc = SYS_CANCEL
                            GoTo Abort_Tran
                        End If
                   Case Else
                        Call File_Error(sts, com + BtSNoWait, "出荷予定", 0)
                        Y_Syuka_Chek_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
                                        
            If sts = BtErrEOF Then
                Exit Do
            End If
        
        
'            If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) <> KAN_KBN_FIN And _
'                  Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then
                                            
            If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) = 0 And _
                Len(Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode))) = 0 Then
            Else
                If StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> Format(ID_KANRI_TBL(ING_No).ID, "000") Or _
                    Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) <> StrConv(App.EXEName, vbUpperCase) Then
                                            '他で使用中
                    Y_Syuka_Chek_Proc = SYS_CANCEL
                    GoTo Abort_Tran
                End If
            End If
                    
            Call UniCode_Conv(Y_SYUREC.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
            Call UniCode_Conv(Y_SYUREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
                                            
            RETRY_CNT = 0
                                            
                                            '出荷予定書込み
            Do
                DoEvents
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > FILE_RETRY Then
                            Y_Syuka_Chek_Proc = SYS_CANCEL
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定", 0)
                        Y_Syuka_Chek_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
                                
            
            If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = Mode Then
                        
            
                Y_SYU_CNT = Y_SYU_CNT + 1
                If Y_SYU_CNT > 1 Then
                                            '複数伝票あり
                    Y_Syuka_Chek_Proc = False
                    GoTo Abort_Tran
            
                End If
                                '1件のみ伝票№＆出荷数KEEP
                SYUKA_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
        
                WK_ID_NO = StrConv(Y_SYUREC.ID_NO, vbUnicode)
                WK_DEN_NO = Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6)
                                   
            End If
            
            com = BtOpGetNext
    
        Loop
    
    End If


    If Len(Trim(WK_ID_NO)) <> 0 Then
        ID_NO = WK_ID_NO
        DEN_NO = WK_DEN_NO
    End If

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Y_Syuka_Chek_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    Y_Syuka_Chek_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If



End Function
Private Function Cancel_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『キャンセル処理（前画面検索）』
'
'-------------------------------------------------------
    
    Cancel_Proc = True
        
    
    Select Case ID_KANRI_TBL(ING_No).Step
    
        Case Step_Start         '子機電源ＯＮ
        Case Step_TANTO_REQ     '担当者要求
            
            Call Re_Send_Proc(Sendbuf)
                        
        Case Step_JGYOBU_REQ    '事業部要求

'            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'            ID_KANRI_TBL(ING_No).JGYOBU = ""
'            Call Start_Proc(Sendbuf)

            '事業部要求でループする
            ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
            ID_KANRI_TBL(ING_No).JGYOBU = ""
            ID_KANRI_TBL(ING_No).NAIGAI = ""
            
'            ID_KANRI_TBL(ING_No).MENU_GRP = ""
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            



        Case Step_NAIGAI_REQ    '国内外要求
            
            ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
            ID_KANRI_TBL(ING_No).JGYOBU = ""
            ID_KANRI_TBL(ING_No).NAIGAI = ""
            
'            ID_KANRI_TBL(ING_No).MENU_GRP = ""
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            
            Call Menu_Send_Proc(Sendbuf)


        Case Step_MENU1_REQ     'メニュー１要求
        
            If UBound(NAIGAI) = 0 Then
                '国内外の切り分けなし
'                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'                ID_KANRI_TBL(ING_No).JGYOBU = ""
'                Call Start_Proc(Sendbuf)
            
                ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            
            
                Call Menu_Send_Proc(Sendbuf)
            
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
                Call Menu_Send_Proc(Sendbuf)
            End If
        
        Case Step_MENU2_REQ     'メニュー２要求
        
            ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        
            Call Menu_Send_Proc(Sendbuf)
        
        
'2006.01.30        Case Step_MENU3_REQ     'メニュー３要求
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_MENU2_REQ
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30
'2006.01.30            Call Menu_Send_Proc(Sendbuf)

        Case Step_Sagyo1_REQ    '作業１要求
'2006.01.30            If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) <> 0 Then
'2006.01.30                ID_KANRI_TBL(ING_No).Step = Step_MENU3_REQ
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30                Call Menu_Send_Proc(Sendbuf)
'2006.01.30            Else
                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) <> 0 Then
                    ID_KANRI_TBL(ING_No).Step = Step_MENU2_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                    Call Menu_Send_Proc(Sendbuf)
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    Call Menu_Send_Proc(Sendbuf)
                End If
'2006.01.30            End If
                                                    '作業２／作業３／作業４要求
        Case Step_Sagyo2_REQ, Step_Sagyo3_REQ, Step_Sagyo4_REQ
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
        
    
    End Select
    
    Cancel_Proc = False


End Function

Private Function Data_Clear_Proc(Mode As Integer, Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   『出荷予定／在庫の予約キャンセル』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim RETRY_CNT   As Integer
    
    
    Data_Clear_Proc = True
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Data_Clear_Proc = SYS_ERR
        Exit Function
    End If
    
    If Mode = 0 Then
                                        '出荷予約の開放
        Call UniCode_Conv(K4_Y_SYU.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
        Call UniCode_Conv(K4_Y_SYU.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        com = BtOpGetGreaterEqual
    Else
        Call UniCode_Conv(K4_Y_SYU.WEL_ID, "")
        Call UniCode_Conv(K4_Y_SYU.PRG_ID, "")
        com = BtOpGetGreater
    End If

    Do
        DoEvents
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
            Select Case sts
                                
                Case BtNoErr
                    If Mode = 0 Then
                        If Format(ID_KANRI_TBL(ING_No).ID, "000") <> StrConv(Y_SYUREC.WEL_ID, vbUnicode) Or _
                             StrConv(App.EXEName, vbUpperCase) <> Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) Then
                            sts = BtErrEOF
                        
                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
                            If sts Then
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                                Data_Clear_Proc = SYS_ERR
                                GoTo Abort_Tran
                            End If
                        
                        End If
                    End If
                    
                    Exit Do
                
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("出荷使用中", "", "", "", "")
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        RETRY_CNT = 0
        
        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
        Do
        
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("出荷使用中", "", "", "", "")
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop


    If Mode = 0 Then
                                        '在庫の開放
        Call UniCode_Conv(K3_ZAIKO.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
        Call UniCode_Conv(K3_ZAIKO.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        com = BtOpGetGreaterEqual
    Else
        Call UniCode_Conv(K3_ZAIKO.WEL_ID, "")
        Call UniCode_Conv(K3_ZAIKO.PRG_ID, "")
        com = BtOpGetGreater
    End If
    
    Do
        DoEvents
        
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    
                    If Mode = 0 Then
                        If Format(ID_KANRI_TBL(ING_No).ID, "000") <> StrConv(ZAIKOREC.WEL_ID, vbUnicode) Or _
                             StrConv(App.EXEName, vbUpperCase) <> Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)) Then
                            sts = BtErrEOF
                        
                        
                            sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
                            If sts Then
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                                Data_Clear_Proc = SYS_ERR
                                GoTo Abort_Tran
                            End If
                        
                        
                        End If
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("在庫使用中", "", "", "", "")
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        RETRY_CNT = 0
        
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")
        Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)
        Do
        
            sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("在庫使用中", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop
End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Data_Clear_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    Data_Clear_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function
Private Function tmpZaiko_Clear_Proc() As Integer
'-------------------------------------------------------
'
'   『在庫データ（一時データ）の消去』
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim RETRY_CNT   As Integer
    
    
    tmpZaiko_Clear_Proc = True
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        tmpZaiko_Clear_Proc = SYS_ERR
        Exit Function
    End If
    
    com = BtOpGetFirst

    Do
        DoEvents
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                                
                Case BtNoErr
                    
                    Exit Do
                
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                        tmpZaiko_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                    tmpZaiko_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        RETRY_CNT = 0
        
        Do
        
            sts = BTRV(BtOpDelete, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                        tmpZaiko_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ（一時データ）")
                    tmpZaiko_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop

End_Tran:
                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        tmpZaiko_Clear_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    tmpZaiko_Clear_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function

Private Function Item_Read_Proc(JGYOBU As String, NAIGAI As String, Hinban As String, RET_JGYOBU As String, RET_NAIGAI As String) As Integer
'-------------------------------------------------------
'
'   『品目マスタ』の読み込み処理
'
'   「外部品番」⇒[Ｊａｎ]⇒「読み替えコード」順次に読み込む
'
'   返り値
'       BtNoErr             :正常終了
'       BtErrKeyNotFound    :未登録
'       上記以外            :Pervasive リターンコード
'
'-------------------------------------------------------
Dim sts As Integer

    
    '--------------------------------------------------外部品番
    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case BtErrKeyNotFound
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select
    '--------------------------------------------------Ｊａｎコード
    Call UniCode_Conv(K4_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K4_ITEM.JAN_CODE, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K4_ITEM, Len(K4_ITEM), 4)
    Select Case sts
        Case BtNoErr
            Hinban = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case BtErrKeyNotFound
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select
    '--------------------------------------------------読替えコード
    Call UniCode_Conv(K5_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K5_ITEM.HIN_CHANGE, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K5_ITEM, Len(K5_ITEM), 5)
    Select Case sts
        Case BtNoErr
            Hinban = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case BtErrKeyNotFound   '2006.01.06 '資材品番での読み替えを追加
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select
    '--------------------------------------------------資材品番で読み替え
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Hinban = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select

End Function


Private Function GOODS_ONOFF_Ono_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『商品化済み→未商品の切り替え（小野用）』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim MENU_NO         As String * 2

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

    GOODS_ONOFF_Ono_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    GOODS_ONOFF_Ono_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                GOODS_ONOFF_Ono_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            GOODS_ONOFF_Ono_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            
                                                
                            
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                GOODS_ONOFF_Ono_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           'ここでは発生しない
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("在庫使用中", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GOODS_ONOFF_Ono_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("有効在庫無し", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                GOODS_ONOFF_Ono_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '棚番をセーブ
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品数量
                                                        
            '数量付きの送信メッセージを作成する
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "08"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                            'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                            '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "総在庫：" & Format((SUMI_QTY + MI_QTY), "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "総在庫：" & Format((SUMI_QTY + MI_QTY), "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            
            
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（商品／未商品数量）
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '数量（ここは無い）
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Ono_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Ono_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Ono_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                        End If
        
                        If i = M_Gyo - 1 Then       '最終行だったら
'                            If SUMI_QTY = 0 And MI_QTY = 0 Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "商品／未商品＝０", "数量入力ミス", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                GOODS_ONOFF_Siga_Proc = False
'                                Exit Function
'                            End If
                
                
                            MI_QTY = (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) - SUMI_QTY
                
                
                            If (SUMI_QTY + MI_QTY) <> (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "総数量変更不可", "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                
                                
                                
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                GOODS_ONOFF_Ono_Proc = False
                                Exit Function
                    
                            End If
                        End If
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                                        'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
        
        
        
            '商品化←→未商品の切り替え更新
            sts = GOODS_ONOFF_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                            ID_KANRI_TBL(ING_No).NAIGAI, _
                                            ID_KANRI_TBL(ING_No).Hinban, _
                                            ID_KANRI_TBL(ING_No).Tanaban, _
                                            SUMI_QTY, _
                                            MI_QTY, _
                                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                            FILE_RETRY)
            
            Select Case sts
                Case False

                Case True       '在庫不足時に発生
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "在庫数不足", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                    GOODS_ONOFF_Ono_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GOODS_ONOFF_Ono_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GOODS_ONOFF_Ono_Proc = SYS_ERR    'システム異常発生
                    GoTo Abort_Tran
            End Select
   
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '出荷予定／在庫の予約解除
            sts = Data_Clear_Proc(0, Sendbuf)
            Select Case sts
                Case SYS_CANCEL
                    Call Err_Send_Proc("データ使用中", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
                                        
                                        
                                        '次の作業要求
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    GOODS_ONOFF_Ono_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function

Private Function GOODS_ONOFF_Siga_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『商品化済み→未商品の切り替え（滋賀用）』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

    GOODS_ONOFF_Siga_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    GOODS_ONOFF_Siga_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                GOODS_ONOFF_Siga_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            GOODS_ONOFF_Siga_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            
                                                
                            
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                GOODS_ONOFF_Siga_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           'ここでは発生しない
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("在庫使用中", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GOODS_ONOFF_Siga_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("有効在庫無し", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                GOODS_ONOFF_Siga_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '棚番をセーブ
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品数量
                                                        
            '数量付きの送信メッセージを作成する
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "08"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                            'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                            '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（商品／未商品数量）
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '数量（ここは無い）
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Siga_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Siga_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Siga_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                        End If
        
                        If i = M_Gyo - 1 Then       '最終行だったら
'                            If SUMI_QTY = 0 And MI_QTY = 0 Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "商品／未商品＝０", "数量入力ミス", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                GOODS_ONOFF_Siga_Proc = False
'                                Exit Function
'                            End If
                
                
                
                            If (SUMI_QTY + MI_QTY) <> (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "総数量変更不可", "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                
                                
                                
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                GOODS_ONOFF_Siga_Proc = False
                                Exit Function
                    
                            End If
                        End If
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                                        'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
        
            '商品化←→未商品の切り替え更新
            sts = GOODS_ONOFF_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                            ID_KANRI_TBL(ING_No).NAIGAI, _
                                            ID_KANRI_TBL(ING_No).Hinban, _
                                            ID_KANRI_TBL(ING_No).Tanaban, _
                                            SUMI_QTY, _
                                            MI_QTY, _
                                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                            FILE_RETRY)
            
            Select Case sts
                Case False

                Case True       '在庫不足時に発生
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "在庫数不足", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                    GOODS_ONOFF_Siga_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GOODS_ONOFF_Siga_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GOODS_ONOFF_Siga_Proc = SYS_ERR    'システム異常発生
                    GoTo Abort_Tran
            End Select
   
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '出荷予定／在庫の予約解除
            sts = Data_Clear_Proc(0, Sendbuf)
            Select Case sts
                Case SYS_CANCEL
                    Call Err_Send_Proc("データ使用中", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
                                        
                                        
                                        '次の作業要求
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    GOODS_ONOFF_Siga_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Private Function GOODS_ONOFF_Update_Proc(JGYOBU As String, _
                                        NAIGAI As String, _
                                        HIN_GAI As String, _
                                        LOCATION As String, _
                                        SUMI_JITU_QTY As Long, _
                                        MI_JITU_QTY As Long, _
                                        ID As String, _
                                        TANTO_CODE As String, _
                                        Optional RETRY As Integer = 10) As Integer
'****************************************************
'*      「商品化／未商品切り替え処理」在庫データ更新
'*
'*  在庫データの更新を行う。
'*  (引数の設定ミスはこちらではチェックしない)
'*  使用ﾌｧｲﾙ    :   在庫データ
'*                  在庫データ(一時ファイル)
'*  引数：  事業部（省略不可）
'*          国内外（省略不可）
'*          品番外部(省略不可)
'*          棚番（XXXXXXXX(倉庫№+列+連+段)省略不可）
'*          商品化済み実績数（何れか一方必須）
'*          未商品実績数　　（　　〃　　　　）
'*          ID(省略不可)
'*          担当者（省略不可）
'*          リトライ(省略可 １桁目:1=画面メッセージ有 0:無，２桁目:リトライ回数(0～9 0:無限))
'*  戻り値: false       :正常
'*          true        :継続可能な異常
'*          SYS_ERR     :継続できない異常
'*          SYS_CANCEL  :更新ｷｬﾝｾﾙ
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim Upd_com     As Integer


Dim RETRY_CNT   As Integer
Dim MESG_FLG    As Integer
Dim RETRY_SU    As Integer
    
Dim Zan_Qty     As Long
Dim WK_Qty      As Long
    
    

    GOODS_ONOFF_Update_Proc = True
                                                                      
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
    
'============================================================ 対象在庫データを在庫一時データに全件移動する。
    
    Call UniCode_Conv(K4_ZAIKO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K4_ZAIKO.Soko_No, Left(LOCATION, 2))
    Call UniCode_Conv(K4_ZAIKO.Retu, Mid(LOCATION, 3, 2))
    Call UniCode_Conv(K4_ZAIKO.Ren, Mid(LOCATION, 5, 2))
    Call UniCode_Conv(K4_ZAIKO.Dan, Right(LOCATION, 2))
    
    
    com = BtOpGetGreaterEqual
    
    RETRY_CNT = 0
    
    
    Do
        DoEvents
'------- 元在庫読込み
        Do
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
            Select Case sts
                Case BtNoErr
                    If JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                        LOCATION <> (StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                        StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                        StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                        StrConv(ZAIKOREC.Dan, vbUnicode)) Then
                        sts = BtErrEOF
                    
                    End If
                
                    Exit Do
                
                Case BtErrEOF
    
                    Exit Do
                
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                'リトライ回数チェック
                    If RETRY_SU <> 0 Then
    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, com + BtSNoWait, "在庫データ", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function
    
                        End If
    
                    End If
    
                    DoEvents
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫データ")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function
            End Select
    
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
'------- 在庫(一時データ)読込み
        Call UniCode_Conv(K0_tmpZAIKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.Retu, StrConv(ZAIKOREC.Retu, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.Ren, StrConv(ZAIKOREC.Ren, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.Dan, StrConv(ZAIKOREC.Dan, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
    
        
        Do
        
            sts = BTRV(BtOpGetEqual + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Upd_com = BtOpUpdate
                    Exit Do
                
                Case BtErrKeyNotFound
                    Upd_com = BtOpInsert
    
                    Exit Do
                
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                'リトライ回数チェック
                    If RETRY_SU <> 0 Then
    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, com + BtSNoWait, "在庫データ（一時データ）", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function
    
                        End If
    
                    End If
    
                    DoEvents
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫データ（一時データ）")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        
        
        Loop
'------- 在庫(一時データ)を出力
        Select Case Upd_com
        
            Case BtOpInsert
            '------- 新規追加
                Do
                    sts = BTRV(BtOpInsert, tmpZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                
                
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then
        
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, BtOpInsert, "在庫データ（一時データ）", 0)
                                    GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                    Exit Function
        
                                End If
        
                            End If
        
                            DoEvents
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ（一時データ）")
                            GOODS_ONOFF_Update_Proc = SYS_ERR
                            Exit Function
        
                    End Select
                
                Loop
    
        
            Case BtOpUpdate
            '------- 在庫数加算（更新）
                Call UniCode_Conv(tmpZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode)) + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "00000000"))
        
        
                Do
                    sts = BTRV(BtOpUpdate, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                
                
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'リトライ回数チェック
                            If RETRY_SU <> 0 Then
        
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '回数オーバー
                                    Call File_Error(sts, BtOpInsert, "在庫データ（一時データ）", 0)
                                    GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                    Exit Function
        
                                End If
        
                            End If
        
                            DoEvents
                        Case Else
                            Call File_Error(sts, Upd_com, "在庫データ（一時データ）")
                            GOODS_ONOFF_Update_Proc = SYS_ERR
                            Exit Function
        
                    End Select
                
                Loop
        
        End Select
   
'------- 元在庫削除
        Do
            sts = BTRV(BtOpDelete, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
        
        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                'リトライ回数チェック
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, BtOpDelete, "在庫データ", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    DoEvents
                Case Else
                    Call File_Error(sts, Upd_com, "在庫データ")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        
        Loop
    
    
        com = BtOpGetNext
    
    Loop
    
'============================================================ 商品化済みの処理(古い日付より引き当てる)
    If SUMI_JITU_QTY <> 0 Then
    
        Zan_Qty = SUMI_JITU_QTY




        Call UniCode_Conv(K0_tmpZAIKO.Soko_No, Left(LOCATION, 2))
        Call UniCode_Conv(K0_tmpZAIKO.Retu, Mid(LOCATION, 3, 2))
        Call UniCode_Conv(K0_tmpZAIKO.Ren, Mid(LOCATION, 5, 2))
        Call UniCode_Conv(K0_tmpZAIKO.Dan, Right(LOCATION, 2))
        Call UniCode_Conv(K0_tmpZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_tmpZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K0_tmpZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K0_tmpZAIKO.NYUKA_DT, "")

        
        com = BtOpGetGreaterEqual
        
        Do

            RETRY_CNT = 0
'------- 在庫（一時データ）読込み
            Do
                sts = BTRV(com + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
                        If JGYOBU <> StrConv(tmpZAIKOREC.JGYOBU, vbUnicode) Or _
                            NAIGAI <> StrConv(tmpZAIKOREC.NAIGAI, vbUnicode) Or _
                            Trim(HIN_GAI) <> Trim(StrConv(tmpZAIKOREC.HIN_GAI, vbUnicode)) Or _
                            LOCATION <> (StrConv(tmpZAIKOREC.Soko_No, vbUnicode) & _
                                            StrConv(tmpZAIKOREC.Retu, vbUnicode) & _
                                            StrConv(tmpZAIKOREC.Ren, vbUnicode) & _
                                            StrConv(tmpZAIKOREC.Dan, vbUnicode)) Then
                            sts = BtErrEOF
                        
                        End If
                    
                    
                        If Zan_Qty < CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            Upd_com = BtOpUpdate
                            WK_Qty = Zan_Qty
                        Else
                            Upd_com = BtOpDelete
                            WK_Qty = CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        End If

                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'リトライ回数チェック
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '回数オーバー
                                Call File_Error(sts, com + BtSNoWait, "在庫データ（一時データ）", 0)
                                GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                Exit Function

                            End If

                        End If

                        DoEvents
                    
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "在庫データ（一時データ）")
                        GOODS_ONOFF_Update_Proc = SYS_ERR
                        Exit Function
                End Select

            Loop

            If sts = BtErrEOF Then
                Exit Do
            End If

            If Upd_com = BtOpUpdate Then
                                                                            '有効在庫数
                Call UniCode_Conv(tmpZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode)) - WK_Qty, "00000000"))
            
            End If


            RETRY_CNT = 0
'------- 在庫（一時データ）消し込み
            Do
                sts = BTRV(Upd_com, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'リトライ回数チェック
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '回数オーバー
                                Call File_Error(sts, Upd_com, "在庫データ（一時データ）", 0)
                                GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                Exit Function

                            End If

                        End If
                        DoEvents
                    Case Else
                        Call File_Error(sts, Upd_com, "在庫データ（一時データ）")
                        GOODS_ONOFF_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            
            Loop
'============================================================ 実際の在庫データに移動
                                                '常に新規追加
            Call UniCode_Conv(ZAIKOREC.Soko_No, Left(LOCATION, 2))          '倉庫№
            Call UniCode_Conv(ZAIKOREC.Retu, Mid(LOCATION, 3, 2))           '列
            Call UniCode_Conv(ZAIKOREC.Ren, Mid(LOCATION, 5, 2))            '連
            Call UniCode_Conv(ZAIKOREC.Dan, Right(LOCATION, 2))             '段
            Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                      '事業部
            Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                      '内外
            Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                    '品番（外部）
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                       '商品／未商品
                                                                            '入荷日
            Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(tmpZAIKOREC.NYUKA_DT, vbUnicode))
                                                                            '入庫日
            Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(tmpZAIKOREC.NYUKO_DT, vbUnicode))
                                                                            '品番（内部）
'            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(tmpZAIKOREC.NYUKO_DT, vbUnicode))  2005.09.03
            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(tmpZAIKOREC.HIN_NAI, vbUnicode))
                                                                            '有効在庫数
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(WK_Qty, "00000000"))
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                   '排他フラグ
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                         '使用中子機ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                         '使用中ﾌﾟﾛｸﾞﾗﾑ
            Call UniCode_Conv(ZAIKOREC.GOODS_YMD, Format(Now, "YYYYMMDD"))  '商品化日付


            Call UniCode_Conv(ZAIKOREC.FILLER, "")

            RETRY_CNT = 0
'*------------------------------------------------------'在庫データ出力
            Do
                sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'リトライ回数チェック
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '回数オーバー
                                Call File_Error(sts, Upd_com, "在庫データ", 0)
                                GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                Exit Function

                            End If

                        End If
                        DoEvents
                    Case Else
                        Call File_Error(sts, BtOpInsert, "在庫データ")
                        GOODS_ONOFF_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            Loop

            Zan_Qty = Zan_Qty - WK_Qty

            If Zan_Qty <= 0 Then
                Exit Do                     '引き落とし終了（商品化済み分）
            End If

        Loop
                
    End If
'================================================================================
    '*
    '*--------------------  未商品化の処理(一時在庫に残っている分は全て未商品として計上)
    
    Call UniCode_Conv(K0_tmpZAIKO.Soko_No, Left(LOCATION, 2))
    Call UniCode_Conv(K0_tmpZAIKO.Retu, Mid(LOCATION, 3, 2))
    Call UniCode_Conv(K0_tmpZAIKO.Ren, Mid(LOCATION, 5, 2))
    Call UniCode_Conv(K0_tmpZAIKO.Dan, Right(LOCATION, 2))
    Call UniCode_Conv(K0_tmpZAIKO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_tmpZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_tmpZAIKO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K0_tmpZAIKO.NYUKA_DT, "")
    
    
    com = BtOpGetGreaterEqual
    
    Do


        DoEvents


        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                Case BtNoErr
                                        '棚＋品ブレーク
                    If LOCATION <> (StrConv(tmpZAIKOREC.Soko_No, vbUnicode) & _
                                        StrConv(tmpZAIKOREC.Retu, vbUnicode) & _
                                        StrConv(tmpZAIKOREC.Ren, vbUnicode) & _
                                        StrConv(tmpZAIKOREC.Dan, vbUnicode)) Or _
                        JGYOBU <> StrConv(tmpZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(tmpZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(tmpZAIKOREC.HIN_GAI, vbUnicode)) Then


                        sts = BtErrEOF

                    End If
                    Exit Do
                Case BtErrEOF
                    
                    Exit Do

                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                'リトライ回数チェック
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, com + BtSNoWait, "在庫データ（一時データ）", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    DoEvents
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "在庫データ（一時データ）")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function
            End Select

        Loop


        If sts = BtErrEOF Then
            Exit Do
        End If

'============================================================ 実際の在庫データに移動
                                                '常に新規追加
        Call UniCode_Conv(ZAIKOREC.Soko_No, Left(LOCATION, 2))          '倉庫№
        Call UniCode_Conv(ZAIKOREC.Retu, Mid(LOCATION, 3, 2))           '列
        Call UniCode_Conv(ZAIKOREC.Ren, Mid(LOCATION, 5, 2))            '連
        Call UniCode_Conv(ZAIKOREC.Dan, Right(LOCATION, 2))          '段
        Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                      '事業部
        Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                      '内外
        Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                    '品番（外部）
        Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                       '商品／未商品
                                                                        '入荷日
        Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(tmpZAIKOREC.NYUKA_DT, vbUnicode))
                                                                        '入庫日
        Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(tmpZAIKOREC.NYUKO_DT, vbUnicode))
                                                                        '品番（内部）
        Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(tmpZAIKOREC.HIN_NAI, vbUnicode))
                                                                        '有効在庫数
        Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode))
        Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                   '排他フラグ
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                         '使用中子機ID
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                         '使用中ﾌﾟﾛｸﾞﾗﾑ
        Call UniCode_Conv(ZAIKOREC.GOODS_YMD, Format(Now, "YYYYMMDD"))  '商品化日付


        Call UniCode_Conv(ZAIKOREC.FILLER, "")

        RETRY_CNT = 0
'*------------------------------------------------------'在庫データ出力
        Do
            sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                'リトライ回数チェック
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, Upd_com, "在庫データ", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If
                    DoEvents
                Case Else
                    Call File_Error(sts, BtOpInsert, "在庫データ")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        Loop

'*------------------------------------------------------'在庫（一時データ）削除
        Do
            sts = BTRV(BtOpDelete, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
        
        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                'リトライ回数チェック
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '回数オーバー
                            Call File_Error(sts, BtOpDelete, "在庫データ（一時データ）", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    DoEvents
                Case Else
                    Call File_Error(sts, BtOpDelete, "在庫データ（一時データ）")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        
        Loop


        com = BtOpGetNext
    
    Loop
                
    
    
    GOODS_ONOFF_Update_Proc = False

End Function

Private Function RETURNED_GOODS_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『良品返品』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    RETURNED_GOODS_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                RETURNED_GOODS_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
        
            SUMI_QTY = 0
            MI_QTY = 0
        
        
            '-----------------------------------------------送信テキスト作成
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban                    '品番をセーブ
                                                        
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '資材対応の事業部2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '資材対応の国内外2006.01.06
                                                        
                                                        
                                                        
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, Left(ID_KANRI_TBL(ING_No).YOIN_DNAME, 2) & "[" & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode) & "]")
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, Left(ID_KANRI_TBL(ING_No).YOIN_DNAME, 2) & "[" & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode) & "]")
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Hinban)

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Hinban)
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "08"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                                    'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                    '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                                                                    '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                    '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                    '入力桁数
            Send_Text.Box_Type(2).Max_Size = "09"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "09"
                                                                    
            Send_Text.Box_Type(2).MENU = ""                         'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#")))) & Format(SUMI_QTY, "#"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#")))) & Format(SUMI_QTY, "#"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#"))) & Format(SUMI_QTY, "#")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#"))) & Format(SUMI_QTY, "#")
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#")))) & Format(MI_QTY, "#"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#")))) & Format(MI_QTY, "#"))
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#"))) & Format(MI_QTY, "#")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#"))) & Format(MI_QTY, "#")
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（商品／未商品数量）
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
                    Case LCD_Tanaban        '棚番
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        If Trim(Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                        '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                    RETURNED_GOODS_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 混載チェック
                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    RETURNED_GOODS_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚番エラー", "", "")
                            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                    RETURNED_GOODS_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
                    
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                    
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "棚使用不可", "", "")
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                
                                RETURNED_GOODS_Proc = False
                                Exit Function
                            End If
                    
                    
                        End If
            
            
            
                    Case LCD_Suryo          '数量（ここは無い）
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            RETURNED_GOODS_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            RETURNED_GOODS_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '数量（商品化済み数量／未商品数量）
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            RETURNED_GOODS_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))

'                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                RETURNED_GOODS_Proc = False
'                                Exit Function
'
'                            End If
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
'                            If MI_QTY > ID_KANRI_TBL(ING_No).Send_MI_QTY Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "数量入力ミス", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                RETURNED_GOODS_Proc = False
'                                Exit Function
'                            End If
                        End If
        
                        If i = M_Gyo - 1 Then       '最終行だったら
                            If SUMI_QTY = 0 And MI_QTY = 0 Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "商品／未商品＝０", "数量入力ミス", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                RETURNED_GOODS_Proc = False
                                Exit Function
                            End If
                        End If
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                        
                                        
                                                'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                                
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                                                
                                                '入庫更新
            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    Format(Now, "YYYYMMDD"), _
                                    Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , , , , MENU_NO)
            Select Case sts
                Case False
                Case True           '入庫時は発生しない
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                    RETURNED_GOODS_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Sendbuf = Text_Create_Proc()
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    RETURNED_GOODS_Proc = SYS_ERR    'システム異常発生
                    
                    GoTo Abort_Tran
            End Select
                                        
                                        
                                        
End_Tran:
                                            'トランザクション終了
        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call Err_Send_Proc("システム異常発生", "", "", "", "")
            Sendbuf = Text_Create_Proc()
            Call File_Error(sts, BtOpEndTransaction, "", 0)
            GoTo Abort_Tran
        End If
                                        
                                        '次の作業要求
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    RETURNED_GOODS_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Location_Move_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   『棚移動指定時のチェック＆更新処理』
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim From_Tanaban    As String * 8
Dim To_Tanaban      As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2
    
    Location_Move_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '１回目の受信（棚番／品番）
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '棚番
                        From_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(From_Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Location_Move_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(From_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(From_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(From_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Location_Move_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Location_Move_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '品番
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ 品目マスタ読込み
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(From_Tanaban) = Loc_OK_Para Then
                                    '棚番OK時の棚番チェック
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            From_Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- エラーメッセージ作成
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "棚番エラー", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Location_Move_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- エラーメッセージ作成
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "品番エラー", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Location_Move_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ 在庫の使用予約を行い、有効在庫数を獲得する
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, From_Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           'ここでは発生しない
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("在庫使用中", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Location_Move_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("有効在庫無し", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                Location_Move_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = From_Tanaban       '棚番をセーブ
            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '資材対応の事業部2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '資材対応の国内外2006.01.06
            
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '品番をセーブ
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '送信する商品化済み数量
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '送信する未商品数量
            
            '数量付きの送信メッセージを作成する
            Send_Text.sts = Sts_OK                                  'ステータス　OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '表示画面フラグ 通常入力画面
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '最終メニューフラグ
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               'メニュー項目数（05固定）
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '送信データファイル名
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           'ブザー音　標準
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------１行目
                                                            'BOX属性
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '数値初期表示
            Send_Text.Box_Type(0).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(0).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(0).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(0).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
            '-----------------------------------------------２行目
                                                            'BOX属性
            Send_Text.Box_Type(1).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                            Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))
                                                            '数値初期表示
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '初期カーソル位置
            Send_Text.Box_Type(1).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(1).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(1).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
            '-----------------------------------------------３行目
                                                            'BOX属性
            Send_Text.Box_Type(2).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                            '数値初期表示
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(2).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(2).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------４行目
                                                            'BOX属性
            Send_Text.Box_Type(3).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Suryo & ":" & Format(SUMI_QTY + MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Suryo & ":" & Format(SUMI_QTY + MI_QTY, "#0"))
                                                            '数値初期表示
            Send_Text.Box_Type(3).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(3).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                            '入力桁数
            Send_Text.Box_Type(3).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            '-----------------------------------------------５行目
                                                            'BOX属性
            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                            '表示内容
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_To_Tanaban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_To_Tanaban)
                                                            '数値初期表示
            Send_Text.Box_Type(4).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '初期カーソル位置
            Send_Text.Box_Type(4).Start_Pos = "01"          '数値は５桁固定
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                            '入力桁数
            Send_Text.Box_Type(4).Max_Size = "09"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "09"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 'メニュ―番号
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '２回目の受信（移動先棚番）
            
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_To_Tanaban         '移動先棚番
                    
                    
                        To_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(From_Tanaban) = Loc_OK_Para Then '棚番OK
                        Else
                            '------------------ 倉庫マスタ読込み
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2), "倉庫エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Location_Move_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                    Exit Function
                            End Select
                            '------------------ 混載チェック
                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "混載エラー", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Location_Move_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ 棚マスタ読込み
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(To_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(To_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(To_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- エラーメッセージ作成
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚番エラー", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Location_Move_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "棚マスタ", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ 禁止棚のチェック
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "棚使用不可", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Location_Move_Proc = False
                                Exit Function
                            End If
            
                        End If
                    
                    
                    
                End Select
            Next i
            '----------------------------------- データ更新処理開始 -----------
                                                        'トランザクション開始
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
        
        
            sts = IDO_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    To_Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    ID_KANRI_TBL(ING_No).Send_SUMI_QTY, _
                                    ID_KANRI_TBL(ING_No).Send_MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , MENU_NO)

    
    
            Select Case sts
                Case False
        
                Case True       '在庫不足時に発生
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "在庫数不足", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                    Location_Move_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "処理中断", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Location_Move_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Location_Move_Proc = SYS_ERR    'システム異常発生
                    GoTo Abort_Tran
            End Select
    
                                        'トランザクション終了
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '出荷予定／在庫の予約解除
            sts = Data_Clear_Proc(0, Sendbuf)
            Select Case sts
                Case SYS_CANCEL
                    Call Err_Send_Proc("データ使用中", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                
                Case SYS_ERR
                    Call Err_Send_Proc("システム異常発生", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
            End Select
                                        
                                        
                                        '次の作業要求
            
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- エラーメッセージ作成
                Case Else
                '重要な要因なので未登録はシステム停止とする
                Call Err_Send_Proc("システム異常発生", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "要因マスタ", 0)
                Exit Function
            End Select
            
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
    
    End Select

    Location_Move_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("システム異常発生", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Private Function Dec_To_Bcd(DecStr As String) As Variant
Dim i           As Long
Dim BCDChr      As Variant

    Dec_To_Bcd = ""

    For i = 1 To Len(DecStr) Step 2
        BCDChr = Chr(Val(Mid(DecStr, i, 1)) * 16 Or Val(Mid(DecStr, i + 1, 1)))
        Dec_To_Bcd = Dec_To_Bcd & BCDChr
    Next i

End Function

Private Function Bcd_To_Dec(BcdStr As String) As Variant
Dim i           As Long
Dim DecLow      As Long

    Bcd_To_Dec = ""

    For i = 1 To Len(BcdStr)
        DecLow = Asc(Mid(BcdStr, i, 1)) Mod 16
        Bcd_To_Dec = Bcd_To_Dec & CStr((Asc(Mid(BcdStr, i, 1)) - DecLow) / 16) & CStr(DecLow)
    Next i

End Function

