VERSION 5.00
Begin VB.Form CONV20060515F 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データコンバート処理"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11070
   ControlBox      =   0   'False
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
   ScaleHeight     =   6495
   ScaleWidth      =   11070
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   8760
      MaxLength       =   2
      TabIndex        =   24
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   8160
      TabIndex        =   21
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   20
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "新移動歴削除"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   3720
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   555
      Index           =   2
      Left            =   7980
      TabIndex        =   18
      Top             =   5820
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｺﾝﾊﾞｰﾄ開始"
      Height          =   555
      Index           =   1
      Left            =   5160
      TabIndex        =   17
      Top             =   5820
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全選択"
      Height          =   1095
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   3540
      Width           =   435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "出荷予定　　　　＝"
      Height          =   375
      Index           =   6
      Left            =   1380
      TabIndex        =   15
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "入荷予定　　　　＝"
      Height          =   375
      Index           =   5
      Left            =   1380
      TabIndex        =   14
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "商品化指図(子)　＝"
      Height          =   375
      Index           =   4
      Left            =   1380
      TabIndex        =   13
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "作業実績ログ　　＝"
      Height          =   375
      Index           =   3
      Left            =   1380
      TabIndex        =   12
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "在庫移動歴　　　＝"
      Height          =   375
      Index           =   2
      Left            =   1380
      TabIndex        =   11
      Top             =   3720
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "発番マスタ　　　＝"
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   10
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "削除済み出荷予定＝"
      Height          =   375
      Index           =   0
      Left            =   1380
      TabIndex        =   9
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Height          =   375
      Left            =   7320
      TabIndex        =   26
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "日以前分"
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   25
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "月"
      Height          =   375
      Index           =   1
      Left            =   8520
      TabIndex        =   23
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "年"
      Height          =   375
      Index           =   0
      Left            =   7920
      TabIndex        =   22
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   3900
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   3900
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   3900
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3900
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3900
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3900
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3900
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label MsgLab 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "データコンバート(ID_NO 8→12桁)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "CONV20060515F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Convert_Proc() As Integer
Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim count           As Long

Dim DISP_INTERVAL   As Long

Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

Dim lngWk           As Double

    Convert_Proc = True

'---------------------------------------------  削除済み出荷予定のコンバート
Convert_P0:
    If Check1(0).Value <> 1 Then GoTo Convert_P1

    MsgLab(1) = "削除済み出荷予定コンバート処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(count, "#0")

    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), K0_O_DEL_SYU, Len(K0_O_DEL_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）削除済み出荷予定")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(DEL_SYUREC.WEL_ID, StrConv(O_DEL_SYUREC.WEL_ID, vbUnicode))   '使用子機ID
        Call UniCode_Conv(DEL_SYUREC.PRG_ID, StrConv(O_DEL_SYUREC.PRG_ID, vbUnicode))   '使用中ﾌﾟﾛｸﾞﾗﾑ
        Call UniCode_Conv(DEL_SYUREC.KAN_KBN, StrConv(O_DEL_SYUREC.KAN_KBN, vbUnicode)) '完了区分
        Call UniCode_Conv(DEL_SYUREC.DT_SYU, StrConv(O_DEL_SYUREC.DT_SYU, vbUnicode))   'ﾃﾞｰﾀ種別
        Call UniCode_Conv(DEL_SYUREC.JGYOBU, StrConv(O_DEL_SYUREC.JGYOBU, vbUnicode))   '事業部区分
        Call UniCode_Conv(DEL_SYUREC.KEY_CYU_KBN, _
                                        StrConv(O_DEL_SYUREC.KEY_CYU_KBN, vbUnicode))   '注文区分
        Call UniCode_Conv(DEL_SYUREC.KEY_ID_NO, _
                                        StrConv(O_DEL_SYUREC.KEY_ID_NO, vbUnicode))     'ID-NO(8桁→12桁)
        Call UniCode_Conv(DEL_SYUREC.NAIGAI, StrConv(O_DEL_SYUREC.NAIGAI, vbUnicode))   '国内外
        Call UniCode_Conv(DEL_SYUREC.KEY_HIN_NO, _
                                        StrConv(O_DEL_SYUREC.KEY_HIN_NO, vbUnicode))    '品目番号
        Call UniCode_Conv(DEL_SYUREC.KEY_MUKE_CODE, _
                                    StrConv(O_DEL_SYUREC.KEY_MUKE_CODE, vbUnicode))     '得意先コード
        Call UniCode_Conv(DEL_SYUREC.KEY_SS_CODE, _
                                    StrConv(O_DEL_SYUREC.KEY_SS_CODE, vbUnicode))       '直送先コード
        Call UniCode_Conv(DEL_SYUREC.KEY_SYUKA_YMD, _
                                    StrConv(O_DEL_SYUREC.KEY_SYUKA_YMD, vbUnicode))     '出荷日付
        Call UniCode_Conv(DEL_SYUREC.JGYOBA, StrConv(O_DEL_SYUREC.JGYOBA, vbUnicode))   '事業場
        Call UniCode_Conv(DEL_SYUREC.DATA_KBN, _
                                        StrConv(O_DEL_SYUREC.DATA_KBN, vbUnicode))      'データ区分
        Call UniCode_Conv(DEL_SYUREC.TORI_KBN, _
                                        StrConv(O_DEL_SYUREC.TORI_KBN, vbUnicode))      '取引区分
        Call UniCode_Conv(DEL_SYUREC.ID_NO, StrConv(O_DEL_SYUREC.ID_NO, vbUnicode))     'ID-NO(8桁→12桁)

'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.KAIKEI_JGYOBA, "")                         '会計用事業場ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.SHISAN_JGYOBA, "")                         '資産管理用事業場ｺｰﾄﾞ
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.HIN_NO, StrConv(O_DEL_SYUREC.HIN_NO, vbUnicode))   '品目番号
        Call UniCode_Conv(DEL_SYUREC.DEN_NO, StrConv(O_DEL_SYUREC.DEN_NO, vbUnicode))   '伝票番号
        Call UniCode_Conv(DEL_SYUREC.SURYO, StrConv(O_DEL_SYUREC.SURYO, vbUnicode))     '出庫数量
        Call UniCode_Conv(DEL_SYUREC.MUKE_CODE, _
                                        StrConv(O_DEL_SYUREC.MUKE_CODE, vbUnicode))     '得意先コード
        Call UniCode_Conv(DEL_SYUREC.SYUKO_SYUSI, _
                                        StrConv(O_DEL_SYUREC.SYUKO_SYUSI, vbUnicode))   '在庫収支

'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.SHISAN_SYUSI, "")                          '資産管理用在庫収支ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.HOJYO_SYUSI, "")                           '補助在庫収支ｺｰﾄﾞ
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.SYUKA_YMD, _
                                        StrConv(O_DEL_SYUREC.SYUKA_YMD, vbUnicode))     '出荷日付

'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.TANKA, _
                                    StrConv(O_DEL_SYUREC.TANKA, vbUnicode))     '実際単価
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.ODER_NO, _
                                        StrConv(O_DEL_SYUREC.ODER_NO, vbUnicode))       'オーダー番号
        Call UniCode_Conv(DEL_SYUREC.ITEM_NO, StrConv(O_DEL_SYUREC.ITEM_NO, vbUnicode)) 'アイテム番号

'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.ODER_NO_R, "")                             '注文管理番号略号
        Call UniCode_Conv(DEL_SYUREC.KOSO_KEITAI, "")                           '個装形態ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.SYUKO_YMD, "")                             '出荷予定日
        Call UniCode_Conv(DEL_SYUREC.TANABAN1, "")                              'ﾛｹｰｼｮﾝ1
        Call UniCode_Conv(DEL_SYUREC.TANABAN2, "")                              'ﾛｹｰｼｮﾝ2
        Call UniCode_Conv(DEL_SYUREC.TANABAN3, "")                              'ﾛｹｰｼｮﾝ3
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.MUKE_NAME, _
                                        StrConv(O_DEL_SYUREC.MUKE_NAME, vbUnicode))     '得意先名称
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN, StrConv(O_DEL_SYUREC.CYU_KBN, vbUnicode)) '注文区分
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN_NAME, _
                                        StrConv(O_DEL_SYUREC.CYU_KBN_NAME, vbUnicode))  '注文区分名称

'''        Call UniCode_Conv(DEL_SYUREC.EXPORT_KBN, _
'''                                        StrConv(O_DEL_SYUREC.EXPORT_KBN, vbUnicode))    '輸出出荷検査区分
'''        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_KBN, _
'''                                    StrConv(O_DEL_SYUREC.LABEL_ISSUE_KBN, vbUnicode))   '個装ﾗﾍﾞﾙ発行区分
'''        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_UNIT, _
'''                                    StrConv(O_DEL_SYUREC.LABEL_ISSUE_UNIT, vbUnicode))  '個装ﾗﾍﾞﾙ発行単位数
'''        Call UniCode_Conv(DEL_SYUREC.LABEL_TANKA_KBN, _
'''                                    StrConv(O_DEL_SYUREC.LABEL_TANKA_KBN, vbUnicode))   '個装ﾗﾍﾞﾙ単価表示区分
'''        Call UniCode_Conv(DEL_SYUREC.TANKA, StrConv(O_DEL_SYUREC.TANKA, vbUnicode))     '単価
'''        Call UniCode_Conv(DEL_SYUREC.KINGAKU, StrConv(O_DEL_SYUREC.KINGAKU, vbUnicode)) '金額
'''        Call UniCode_Conv(DEL_SYUREC.BIKOU2, StrConv(O_DEL_SYUREC.BIKOU2, vbUnicode))   '備考２
'''        Call UniCode_Conv(DEL_SYUREC.REBATE_KBN, _
'''                                        StrConv(O_DEL_SYUREC.REBATE_KBN, vbUnicode))    'リベート区分
'''        Call UniCode_Conv(DEL_SYUREC.CHOHA_KBN, _
'''                                        StrConv(O_DEL_SYUREC.CHOHA_KBN, vbUnicode))     '帳端区分
'''        Call UniCode_Conv(DEL_SYUREC.ATAISA_KBN, _
'''                                        StrConv(O_DEL_SYUREC.ATAISA_KBN, vbUnicode))    '値差区分
'''        Call UniCode_Conv(DEL_SYUREC.REP_KISHU, _
'''                                        StrConv(O_DEL_SYUREC.REP_KISHU, vbUnicode))     '代表機種
'''        Call UniCode_Conv(DEL_SYUREC.NS_KANRI_NO, _
'''                                        StrConv(O_DEL_SYUREC.NS_KANRI_NO, vbUnicode))   'ＮＳ管理番号
'''        Call UniCode_Conv(DEL_SYUREC.MTS_HIN_CODE, _
'''                                        StrConv(O_DEL_SYUREC.MTS_HIN_CODE, vbUnicode))  'ＭＴＳ部品コード
'''        Call UniCode_Conv(DEL_SYUREC.BIKOU1, StrConv(O_DEL_SYUREC.BIKOU1, vbUnicode))   '備考１
'''        Call UniCode_Conv(DEL_SYUREC.CHOKU_KBN, _
'''                                        StrConv(O_DEL_SYUREC.CHOKU_KBN, vbUnicode))     '直送区分
'''        Call UniCode_Conv(DEL_SYUREC.REBATE_RATE, _
'''                                        StrConv(O_DEL_SYUREC.REBATE_RATE, vbUnicode))   'リベート率
'''        Call UniCode_Conv(DEL_SYUREC.HIN_NAME, _
'''                                        StrConv(O_DEL_SYUREC.HIN_NAME, vbUnicode))      '品名
'''        Call UniCode_Conv(DEL_SYUREC.JGYOBA_GAI, _
'''                                        StrConv(O_DEL_SYUREC.JGYOBA_GAI, vbUnicode))    '対外事業場
'''        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, _
'''                                        StrConv(O_DEL_SYUREC.KISHU_CODE, vbUnicode))    '機種コード
'''        Call UniCode_Conv(DEL_SYUREC.SS_CODE, StrConv(O_DEL_SYUREC.SS_CODE, vbUnicode)) '直送先コード


'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(DEL_SYUREC.ORIGIN1, "")                              '原産国1
        Call UniCode_Conv(DEL_SYUREC.ORIGIN2, "")                              '原産国2
        Call UniCode_Conv(DEL_SYUREC.BIKOU2, _
                                        StrConv(O_DEL_SYUREC.BIKOU2, vbUnicode))    '備考2
        Call UniCode_Conv(DEL_SYUREC.HAN_KBN, "")                                   '販売区分
        Call UniCode_Conv(DEL_SYUREC.CHOKU_KBN, _
                                        StrConv(O_DEL_SYUREC.CHOKU_KBN, vbUnicode)) '直送指示区分
        Call UniCode_Conv(DEL_SYUREC.UNIT_ID_NO, "")                                   'ﾕﾆｯﾄ修正管理番号
        Call UniCode_Conv(DEL_SYUREC.ZAIKO_HIKIATE, "")                               '在庫引当順序
        Call UniCode_Conv(DEL_SYUREC.GOKON_KANRI_NO, "")                                  '合梱管理番号
        Call UniCode_Conv(DEL_SYUREC.JYUCHU_ZAN, "")                                '受注残数量
        Call UniCode_Conv(DEL_SYUREC.KYOKYU_KBN, "")                                '供給区分
        Call UniCode_Conv(DEL_SYUREC.SHOHIN_SYUSI, "")                             '商品化納品在庫収支ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.S_SHISAN_SYUSI, "")                            '商品化納品資産管理収支ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.S_HOJYO_SYUSI, "")                             '商品化納品補助収支ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.BIKOU1, _
                                        StrConv(O_DEL_SYUREC.BIKOU1, vbUnicode))    '備考1
        Call UniCode_Conv(DEL_SYUREC.CHOHA_KBN, _
                                        StrConv(O_DEL_SYUREC.CHOHA_KBN, vbUnicode)) '帳端区分
        Call UniCode_Conv(DEL_SYUREC.JYU_HIN_NO, "")                                '受付品目番号
        Call UniCode_Conv(DEL_SYUREC.HIN_NAME, _
                                        StrConv(O_DEL_SYUREC.HIN_NAME, vbUnicode))  '品名
        Call UniCode_Conv(DEL_SYUREC.HIN_CHANGE_KBN, "")                          '品目番号変更区分
        Call UniCode_Conv(DEL_SYUREC.MODULE_EXCHANGE, "")                                'ﾓｼﾞｭｰﾙ交換区分
        Call UniCode_Conv(DEL_SYUREC.ZAIKO_SYUSI, "")                           '残在庫まとめ在庫収支ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.ZAN_SHISAN_SYUSI, "")                          '残在庫まとめ資産管理収支ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.ZAN_HOJYO_SYUSI, "")                           '残在庫まとめ補助収支ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.NOUKI_YMD, "")                                     '指定納期
        Call UniCode_Conv(DEL_SYUREC.SERVICE_KANRI_NO, "")                                'ｻｰﾋﾞｽ会社管理番号
        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, _
                                    StrConv(O_DEL_SYUREC.KISHU_CODE, vbUnicode))    '機種品目ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, "")                                 '環境企画部品区分
        Call UniCode_Conv(DEL_SYUREC.SS_CODE, _
                                    StrConv(O_DEL_SYUREC.SS_CODE, vbUnicode))       '直送相手先ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.KEPIN_KAIJYO, "")                              '欠品解消区分
'------------------------------------------------------------------------------

        Call UniCode_Conv(DEL_SYUREC.HIN_NAI, StrConv(O_DEL_SYUREC.HIN_NAI, vbUnicode)) '品番（内部）
        Call UniCode_Conv(DEL_SYUREC.HTANABAN, _
                                        StrConv(O_DEL_SYUREC.HTANABAN, vbUnicode))      'ホスト棚番
        Call UniCode_Conv(DEL_SYUREC.PRINT_YMD, _
                                        StrConv(O_DEL_SYUREC.PRINT_YMD, vbUnicode))     '出庫表印刷日付
        Call UniCode_Conv(DEL_SYUREC.KAN_YMD, StrConv(O_DEL_SYUREC.KAN_YMD, vbUnicode)) '完了日付
        Call UniCode_Conv(DEL_SYUREC.KENPIN_YMD, _
                                        StrConv(O_DEL_SYUREC.KENPIN_YMD, vbUnicode))    '検品日付
        Call UniCode_Conv(DEL_SYUREC.TOK_KBN, StrConv(O_DEL_SYUREC.TOK_KBN, vbUnicode)) '特売り区分
        Call UniCode_Conv(DEL_SYUREC.JITU_SURYO, _
                                        StrConv(O_DEL_SYUREC.JITU_SURYO, vbUnicode))    '出庫実績数量
        Call UniCode_Conv(DEL_SYUREC.INS_NOW, StrConv(O_DEL_SYUREC.INS_NOW, vbUnicode)) '取込み日時
        Call UniCode_Conv(DEL_SYUREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<DEL_SYU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "削除済み出荷予定")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(0).Caption = Format(count, "#0")

'---------------------------------------------  発番マスタのコンバート
Convert_P1:
    If Check1(1).Value <> 1 Then GoTo Convert_P2

    MsgLab(1) = "発番マスタコンバート処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(count, "#0")

    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), K0_O_HATUBAN, Len(K0_O_HATUBAN), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）発番マスタ")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(1).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(HATUBANREC.JGYOBU, StrConv(O_HATUBANREC.JGYOBU, vbUnicode))   '事業部区分
        Call UniCode_Conv(HATUBANREC.NYK_KBN, StrConv(O_HATUBANREC.NYK_KBN, vbUnicode)) '入荷伝票№区分
        Call UniCode_Conv(HATUBANREC.NYK_DEN_NO, _
                                        StrConv(O_HATUBANREC.NYK_DEN_NO, vbUnicode))    '次入荷伝票№
        Call UniCode_Conv(HATUBANREC.SYK_KBN, StrConv(O_HATUBANREC.SYK_KBN, vbUnicode)) '出荷伝票№区分
        Call UniCode_Conv(HATUBANREC.SYK_DEN_NO, _
                                        StrConv(O_HATUBANREC.SYK_DEN_NO, vbUnicode))    '次出荷伝票№
        Call UniCode_Conv(HATUBANREC.NYK_ID_KBN, _
                                        StrConv(O_HATUBANREC.NYK_ID_KBN, vbUnicode))    '入荷ID№区分
        lngWk = Val(StrConv(O_HATUBANREC.NYK_ID_NO, vbUnicode))
        Call UniCode_Conv(HATUBANREC.NYK_ID_NO, Format(lngWk, String(11, "0")))         '次入荷ID№(8桁→11桁)
        Call UniCode_Conv(HATUBANREC.SYK_ID_KBN, _
                                        StrConv(O_HATUBANREC.SYK_ID_KBN, vbUnicode))    '出荷ID№区分
        lngWk = Val(StrConv(O_HATUBANREC.SYK_ID_NO, vbUnicode))
        Call UniCode_Conv(HATUBANREC.SYK_ID_NO, Format(lngWk, String(11, "0")))         '次出荷ID№(7桁→11桁)
        Call UniCode_Conv(HATUBANREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<HATUBAN.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "発番マスタ")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop


    Cnt(1).Caption = Format(count, "#0")

'---------------------------------------------  在庫移動歴のコンバート
Convert_P2:
    If Check1(2).Value <> 1 Then GoTo Convert_P3

    MsgLab(1) = "在庫移動歴コンバート処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_IDO_POS, O_IDOREC, Len(O_IDOREC), K0_O_IDO, Len(K0_O_IDO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）在庫移動歴")
                Exit Function
        End Select


        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(IDOREC.JITU_DT, StrConv(O_IDOREC.JITU_DT, vbUnicode))         '実績日付
        Call UniCode_Conv(IDOREC.JITU_TM, StrConv(O_IDOREC.JITU_TM, vbUnicode))         '実績時刻
        Call UniCode_Conv(IDOREC.JGYOBU, StrConv(O_IDOREC.JGYOBU, vbUnicode))           '事業部区分
        Call UniCode_Conv(IDOREC.NAIGAI, StrConv(O_IDOREC.NAIGAI, vbUnicode))           '国内外
        Call UniCode_Conv(IDOREC.HIN_GAI, StrConv(O_IDOREC.HIN_GAI, vbUnicode))         '品目（外部）
        Call UniCode_Conv(IDOREC.RIRK_ID, StrConv(O_IDOREC.RIRK_ID, vbUnicode))         '履歴種別
        Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, _
                                        StrConv(O_IDOREC.SUMI_JITU_QTY, vbUnicode))     '実績数量(商品化済み)
        Call UniCode_Conv(IDOREC.MI_JITU_QTY, StrConv(O_IDOREC.MI_JITU_QTY, vbUnicode)) '実績数量(実績数量(未商品))
        Call UniCode_Conv(IDOREC.FROM_SOKO, StrConv(O_IDOREC.FROM_SOKO, vbUnicode))     'From 倉庫№
        Call UniCode_Conv(IDOREC.FROM_RETU, StrConv(O_IDOREC.FROM_RETU, vbUnicode))     'From 列
        Call UniCode_Conv(IDOREC.FROM_REN, StrConv(O_IDOREC.FROM_REN, vbUnicode))       'From 連
        Call UniCode_Conv(IDOREC.FROM_DAN, StrConv(O_IDOREC.FROM_DAN, vbUnicode))       'From 段
        Call UniCode_Conv(IDOREC.TO_SOKO, StrConv(O_IDOREC.TO_SOKO, vbUnicode))         'TO 倉庫№
        Call UniCode_Conv(IDOREC.TO_RETU, StrConv(O_IDOREC.TO_RETU, vbUnicode))         'TO 列
        Call UniCode_Conv(IDOREC.TO_REN, StrConv(O_IDOREC.TO_REN, vbUnicode))           'TO 連
        Call UniCode_Conv(IDOREC.TO_DAN, StrConv(O_IDOREC.TO_DAN, vbUnicode))           'TO 段
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(O_IDOREC.DEN_DT, vbUnicode))           '伝票日付
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(O_IDOREC.DEN_NO, vbUnicode))           '伝票№
        Call UniCode_Conv(IDOREC.PRG_ID, StrConv(O_IDOREC.PRG_ID, vbUnicode))           '出力元プログラム
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(O_IDOREC.HIN_NAI, vbUnicode))         '品番（内部）
        Call UniCode_Conv(IDOREC.NYUKA_DT, StrConv(O_IDOREC.NYUKA_DT, vbUnicode))       '入荷日付
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(O_IDOREC.NYUKO_DT, vbUnicode))       '入庫日付
        Call UniCode_Conv(IDOREC.WEL_ID, StrConv(O_IDOREC.WEL_ID, vbUnicode))           '対象端末№
        Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(O_IDOREC.RIRK_NAME, vbUnicode))     '履歴種別名称
        Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(O_IDOREC.HIN_NAME, vbUnicode))       '品名
        Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, _
                                    StrConv(O_IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode))    '品目別在庫数（商品化済み）
        Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, _
                                    StrConv(O_IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))      '品目別在庫数（未商品）
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.SUMI_FROM_TANA_Zaiko_Qty, vbUnicode))  'FROM棚別品目別在庫数（商品化済み）
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.SUMI_TO_TANA_Zaiko_Qty, vbUnicode))    'TO棚別品目別在庫数（商品化済み）
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.MI_FROM_TANA_Zaiko_Qty, vbUnicode))    'FROM棚別品目別在庫数（未商品）
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, _
                                StrConv(O_IDOREC.MI_TO_TANA_Zaiko_Qty, vbUnicode))      'TO棚別品目別在庫数（未商品）
        Call UniCode_Conv(IDOREC.TOKU_MARK, StrConv(O_IDOREC.TOKU_MARK, vbUnicode))     '特売りマーク
        Call UniCode_Conv(IDOREC.MEMO, StrConv(O_IDOREC.MEMO, vbUnicode))               'メモ
        Call UniCode_Conv(IDOREC.TANTO_CODE, StrConv(O_IDOREC.TANTO_CODE, vbUnicode))                                       '担当者コード
        Call UniCode_Conv(IDOREC.TANTO_NAME, StrConv(O_IDOREC.TANTO_NAME, vbUnicode))                                        '担当者名称
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(O_IDOREC.MUKE_CODE, vbUnicode))     '得意先コード
        Call UniCode_Conv(IDOREC.MUKE_NAME, StrConv(O_IDOREC.MUKE_DNAME, vbUnicode))    '得意先名称
        Call UniCode_Conv(IDOREC.SS_CODE, StrConv(O_IDOREC.SS_CODE, vbUnicode))                                           '直送先コード
        Call UniCode_Conv(IDOREC.SS_NAME, StrConv(O_IDOREC.SS_NAME, vbUnicode))                                           '直送先名称
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(O_IDOREC.MUKE_DNAME, vbUnicode))   '得意先略称
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, StrConv(O_IDOREC.MUKE_CHG_CD, vbUnicode)) '向け先読替えコード
        Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(O_IDOREC.SUM_KBN, vbUnicode))         '集計区分
        Call UniCode_Conv(IDOREC.ID_NO, StrConv(O_IDOREC.ID_NO, vbUnicode))             'ID-NO(8桁→12桁)
        Call UniCode_Conv(IDOREC.Ins_DateTime, _
                                        StrConv(O_IDOREC.Ins_DateTime, vbUnicode))      '挿入日時
        Call UniCode_Conv(IDOREC.SHIIRE_CODE, StrConv(O_IDOREC.SHIIRE_CODE, vbUnicode)) '仕入先ｺｰﾄﾞ
        Call UniCode_Conv(IDOREC.SHIIRE_TANKA, _
                                        StrConv(O_IDOREC.SHIIRE_TANKA, vbUnicode))      '仕入単価(9(8)V99)
        Call UniCode_Conv(IDOREC.KEIJYO_YM, StrConv(O_IDOREC.KEIJYO_YM, vbUnicode))     '計上年月(YYYYMM)
        Call UniCode_Conv(IDOREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "在庫移動歴")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(2).Caption = Format(count, "#0")

'---------------------------------------------  作業実績ログのコンバート
Convert_P3:
    If Check1(3).Value <> 1 Then GoTo Convert_P4

    MsgLab(1) = "作業実績ログコンバート処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_P_SAGYO_LOG_POS, O_P_SAGYO_LOG_REC, Len(O_P_SAGYO_LOG_REC), K0_O_P_SAGYO_LOG, Len(K0_O_P_SAGYO_LOG), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）作業実績ログ")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(3).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_DT, _
                                StrConv(O_P_SAGYO_LOG_REC.JITU_DT, vbUnicode))      '実績日付
        Call UniCode_Conv(P_SAGYO_LOG_REC.JITU_TM, _
                                StrConv(O_P_SAGYO_LOG_REC.JITU_TM, vbUnicode))      '実績時刻
        Call UniCode_Conv(P_SAGYO_LOG_REC.TANTO_CODE, _
                                StrConv(O_P_SAGYO_LOG_REC.TANTO_CODE, vbUnicode))   '担当者コード
        Call UniCode_Conv(P_SAGYO_LOG_REC.WEL_ID, _
                                StrConv(O_P_SAGYO_LOG_REC.WEL_ID, vbUnicode))       '対象端末№
        Call UniCode_Conv(P_SAGYO_LOG_REC.JGYOBU, _
                                StrConv(O_P_SAGYO_LOG_REC.JGYOBU, vbUnicode))       '事業部区分
        Call UniCode_Conv(P_SAGYO_LOG_REC.NAIGAI, _
                                StrConv(O_P_SAGYO_LOG_REC.NAIGAI, vbUnicode))       '国内外
        Call UniCode_Conv(P_SAGYO_LOG_REC.MENU_NO, _
                                StrConv(O_P_SAGYO_LOG_REC.MENU_NO, vbUnicode))      'メニューグループ№
        Call UniCode_Conv(P_SAGYO_LOG_REC.RIRK_ID, _
                                StrConv(O_P_SAGYO_LOG_REC.RIRK_ID, vbUnicode))      '履歴種別
        Call UniCode_Conv(P_SAGYO_LOG_REC.ID_NO, _
                                StrConv(O_P_SAGYO_LOG_REC.ID_NO, vbUnicode))        'ID-NO
        Call UniCode_Conv(P_SAGYO_LOG_REC.HIN_GAI, _
                                StrConv(O_P_SAGYO_LOG_REC.HIN_GAI, vbUnicode))      '品番（外部）
        Call UniCode_Conv(P_SAGYO_LOG_REC.SUMI_JITU_QTY, _
                            StrConv(O_P_SAGYO_LOG_REC.SUMI_JITU_QTY, vbUnicode))    '実績数量(商品化済み)
        Call UniCode_Conv(P_SAGYO_LOG_REC.MI_JITU_QTY, _
                            StrConv(O_P_SAGYO_LOG_REC.MI_JITU_QTY, vbUnicode))      '実績数量(未商品)
        Call UniCode_Conv(P_SAGYO_LOG_REC.MUKE_CODE, _
                            StrConv(O_P_SAGYO_LOG_REC.MUKE_CODE, vbUnicode))        '得意先コード
        Call UniCode_Conv(P_SAGYO_LOG_REC.SS_CODE, _
                            StrConv(O_P_SAGYO_LOG_REC.SS_CODE, vbUnicode))          '直送先コード
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_SOKO, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_SOKO, vbUnicode))        'From 倉庫№
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_RETU, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_RETU, vbUnicode))        '   　列
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_REN, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_REN, vbUnicode))         '   　連
        Call UniCode_Conv(P_SAGYO_LOG_REC.FROM_DAN, _
                            StrConv(O_P_SAGYO_LOG_REC.FROM_DAN, vbUnicode))         '   　段
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_SOKO, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_SOKO, vbUnicode))          'ＴＯ 倉庫№
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_RETU, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_RETU, vbUnicode))          '   　列
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_REN, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_REN, vbUnicode))           '   　連
        Call UniCode_Conv(P_SAGYO_LOG_REC.TO_DAN, _
                            StrConv(O_P_SAGYO_LOG_REC.TO_DAN, vbUnicode))           '   　段
        Call UniCode_Conv(P_SAGYO_LOG_REC.PRG_ID, _
                            StrConv(O_P_SAGYO_LOG_REC.PRG_ID, vbUnicode))           '出力元プログラム
        Call UniCode_Conv(P_SAGYO_LOG_REC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SAGYO_LOG.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "作業実績ログ")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(3).Caption = Format(count, "#0")

'---------------------------------------------  商品化指図（子）のコンバート
Convert_P4:
    If Check1(4).Value <> 1 Then GoTo Convert_P5

    MsgLab(1) = "商品化指図（子）コンバート処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(4).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_P_SSHIJI_K_POS, O_P_SSHIJI_K_REC, Len(O_P_SSHIJI_K_REC), K0_O_P_SSHIJI_K, Len(K0_O_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）商品化指図（子）")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(4).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If


        Call UniCode_Conv(P_SSHIJI_K_REC.SHIJI_NO, _
                                    StrConv(O_P_SSHIJI_K_REC.SHIJI_NO, vbUnicode))      '指図票№
        Call UniCode_Conv(P_SSHIJI_K_REC.DATA_KBN, _
                                    StrConv(O_P_SSHIJI_K_REC.DATA_KBN, vbUnicode))      'ﾃﾞｰﾀ区分
        Call UniCode_Conv(P_SSHIJI_K_REC.SEQNO, _
                                    StrConv(O_P_SSHIJI_K_REC.SEQNO, vbUnicode))         '追番
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_SYUBETSU, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode))   '子 種別
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_JGYOBU, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))     '子 事業部
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_NAIGAI, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))     '子 国内外
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_HIN_GAI, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))    '子 品番
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_QTY, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_QTY, vbUnicode))        '子 員数(999V99)
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_SHIJI_QTY, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode))  '指示数(9(8)V99)
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_BIKOU, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_BIKOU, vbUnicode))      '子 備考
        Call UniCode_Conv(P_SSHIJI_K_REC.KO_ID_NO, _
                                    StrConv(O_P_SSHIJI_K_REC.KO_ID_NO, vbUnicode))      '子 ID_NO
        Call UniCode_Conv(P_SSHIJI_K_REC.CALCEL_F, _
                                    StrConv(O_P_SSHIJI_K_REC.CALCEL_F, vbUnicode))      'ｷｬﾝｾﾙF
        Call UniCode_Conv(P_SSHIJI_K_REC.CANCEL_DATETIME, _
                                StrConv(O_P_SSHIJI_K_REC.CANCEL_DATETIME, vbUnicode))   'ｷｬﾝｾﾙ日時
        Call UniCode_Conv(P_SSHIJI_K_REC.FILLER, "")
        Call UniCode_Conv(P_SSHIJI_K_REC.UPD_DATETIME, _
                                    StrConv(O_P_SSHIJI_K_REC.UPD_DATETIME, vbUnicode))  '更新 日時

        Do
            sts = BTRV(BtOpInsert, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<P_SSHIJI_K.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "商品化指図（子）")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(4).Caption = Format(count, "#0")

'    GoTo Convert_End

'---------------------------------------------  入荷予定のコンバート
Convert_P5:
    If Check1(5).Value <> 1 Then GoTo Convert_P6

    MsgLab(1) = "入荷予定コンバート処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(5).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), K0_O_Y_NYU, Len(K0_O_Y_NYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）入荷予定")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(5).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

        Call UniCode_Conv(Y_NYUREC.KAN_KBN, StrConv(O_Y_NYUREC.KAN_KBN, vbUnicode))     '完了区分
        Call UniCode_Conv(Y_NYUREC.DT_SYU, StrConv(O_Y_NYUREC.DT_SYU, vbUnicode))       'データ種別
        Call UniCode_Conv(Y_NYUREC.JGYOBU, StrConv(O_Y_NYUREC.JGYOBU, vbUnicode))       '事業部区分
        Call UniCode_Conv(Y_NYUREC.NAIGAI, StrConv(O_Y_NYUREC.NAIGAI, vbUnicode))       '国内外
        Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(O_Y_NYUREC.TEXT_NO, vbUnicode))     'テキスト№
        Call UniCode_Conv(Y_NYUREC.JGYOBA, StrConv(O_Y_NYUREC.JGYOBA, vbUnicode))       '事業場
        Call UniCode_Conv(Y_NYUREC.DATA_KBN, StrConv(O_Y_NYUREC.DATA_KBN, vbUnicode))   'データ区分
        Call UniCode_Conv(Y_NYUREC.TORI_KBN, StrConv(O_Y_NYUREC.TORI_KBN, vbUnicode))   '取引区分
        Call UniCode_Conv(Y_NYUREC.ID_NO, StrConv(O_Y_NYUREC.ID_NO, vbUnicode))         'ID-NO(8桁→12桁)
        Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(O_Y_NYUREC.HIN_NO, vbUnicode))       '品目番号
        Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(O_Y_NYUREC.DEN_NO, vbUnicode))       '伝票番号
        Call UniCode_Conv(Y_NYUREC.SURYO, StrConv(O_Y_NYUREC.SURYO, vbUnicode))         '出庫数量
        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, StrConv(O_Y_NYUREC.MUKE_CODE, vbUnicode)) '出庫先
        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, _
                                        StrConv(O_Y_NYUREC.SYUKO_SYUSI, vbUnicode))     '出庫収支
        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, StrConv(O_Y_NYUREC.SYUKO_YMD, vbUnicode)) '出庫日付
        Call UniCode_Conv(Y_NYUREC.TANKA, StrConv(O_Y_NYUREC.TANKA, vbUnicode))         '単価
        Call UniCode_Conv(Y_NYUREC.ODER_NO, StrConv(O_Y_NYUREC.ODER_NO, vbUnicode))     'オーダー番号
        Call UniCode_Conv(Y_NYUREC.ITEM_NO, StrConv(O_Y_NYUREC.ITEM_NO, vbUnicode))     'アイテム番号
        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, StrConv(O_Y_NYUREC.ODER_R_NO, vbUnicode)) 'オーダー略号
        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, _
                                        StrConv(O_Y_NYUREC.KOSO_KEITAI, vbUnicode))     '個装形態
        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(O_Y_NYUREC.SYUKA_YMD, vbUnicode)) '出荷日
        Call UniCode_Conv(Y_NYUREC.TANABAN1, StrConv(O_Y_NYUREC.TANABAN1, vbUnicode))   '棚番１
        Call UniCode_Conv(Y_NYUREC.TANABAN2, StrConv(O_Y_NYUREC.TANABAN2, vbUnicode))   '棚番２
        Call UniCode_Conv(Y_NYUREC.TANABAN3, StrConv(O_Y_NYUREC.TANABAN3, vbUnicode))   '棚番３
        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, StrConv(O_Y_NYUREC.MUKE_NAME, vbUnicode)) '出庫先名称
        Call UniCode_Conv(Y_NYUREC.CYU_KBN, StrConv(O_Y_NYUREC.CYU_KBN, vbUnicode))     '注文区分
        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, _
                                        StrConv(O_Y_NYUREC.CYU_KBN_NAME, vbUnicode))    '注文区分名称
        Call UniCode_Conv(Y_NYUREC.ORIGIN1, StrConv(O_Y_NYUREC.ORIGIN1, vbUnicode))     '原産国１
        Call UniCode_Conv(Y_NYUREC.ORIGIN2, StrConv(O_Y_NYUREC.ORIGIN2, vbUnicode))     '原産国２
        Call UniCode_Conv(Y_NYUREC.BIKOU2, StrConv(O_Y_NYUREC.BIKOU2, vbUnicode))       '備考２
        Call UniCode_Conv(Y_NYUREC.HAN_KBN, StrConv(O_Y_NYUREC.HAN_KBN, vbUnicode))     '販売区分
        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, StrConv(O_Y_NYUREC.CHOKU_KBN, vbUnicode)) '直送区分
        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, _
                                        StrConv(O_Y_NYUREC.UNIT_ID_NO, vbUnicode))      'ﾕﾆｯﾄ修理ID-NO
        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, _
                                        StrConv(O_Y_NYUREC.ZAIKO_HIKIATE, vbUnicode))   '在庫引当順序
        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, _
                                        StrConv(O_Y_NYUREC.GOKON_KANRI_NO, vbUnicode))  '合梱管理番号
        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, StrConv(O_Y_NYUREC.JUCHU_ZAN, vbUnicode)) '受注残数量
        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, _
                                        StrConv(O_Y_NYUREC.KYOKYU_KBN, vbUnicode))      '供給区分
        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, _
                                        StrConv(O_Y_NYUREC.SHOHIN_SYUSI, vbUnicode))    '商品化納入先収支
        Call UniCode_Conv(Y_NYUREC.BIKOU1, StrConv(O_Y_NYUREC.BIKOU1, vbUnicode))       '備考１
        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, StrConv(O_Y_NYUREC.CHOHA_KBN, vbUnicode)) '帳端区分
        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, _
                                        StrConv(O_Y_NYUREC.JYU_HIN_NO, vbUnicode))      '受注品目番号
        Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(O_Y_NYUREC.HIN_NAME, vbUnicode))   '品名
        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, _
                                        StrConv(O_Y_NYUREC.HIN_CHANGE_KBN, vbUnicode))  '品番変更区分
        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, _
                                        StrConv(O_Y_NYUREC.MODULE_EXCHANGE, vbUnicode)) 'モジュール交換区分
        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, _
                                        StrConv(O_Y_NYUREC.ZAIKO_SYUSI, vbUnicode))     '残在庫まとめ在庫収支コード
        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, StrConv(O_Y_NYUREC.NOUKI_YMD, vbUnicode)) '指定納期
        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, _
                                    StrConv(O_Y_NYUREC.SERVICE_KANRI_NO, vbUnicode))    'サービス会社管理番号
        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, StrConv(O_Y_NYUREC.KI_HIN_NO, vbUnicode)) '機種品目コード
        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, _
                                        StrConv(O_Y_NYUREC.ENVIRONMENT_KBN, vbUnicode)) '環境規格部品区分
        Call UniCode_Conv(Y_NYUREC.KAN_DT, StrConv(O_Y_NYUREC.KAN_DT, vbUnicode))       '完了日付
        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, _
                                        StrConv(O_Y_NYUREC.BEF_NYU_QTY, vbUnicode))     '先行入荷数
        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, _
                                        StrConv(O_Y_NYUREC.YOSAN_FROM, vbUnicode))      '予算単位（元）
        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(O_Y_NYUREC.YOSAN_TO, vbUnicode))   '予算単位（先）
        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(O_Y_NYUREC.HTANABAN, vbUnicode))   '標準棚番
        Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(O_Y_NYUREC.HIN_NAI, vbUnicode))     '品番（内部）
        Call UniCode_Conv(Y_NYUREC.FILLER, "")

        Do
            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_NYU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "入荷予定")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(5).Caption = Format(count, "#0")

'    GoTo Convert_End

'---------------------------------------------  出荷予定のコンバート
Convert_P6:
    If Check1(6).Value <> 1 Then GoTo Convert_End

    MsgLab(1) = "出荷予定データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    Cnt(6).Caption = Format(count, "#0")

    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), K0_O_Y_SYU, Len(K0_O_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）出荷予定データ")
                Exit Function
        End Select


        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(6).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If

                                                                '使用子機ID
        Call UniCode_Conv(Y_SYUREC.WEL_ID, StrConv(O_Y_SYUREC.WEL_ID, vbUnicode))
                                                                '使用中プログラム
        Call UniCode_Conv(Y_SYUREC.PRG_ID, StrConv(O_Y_SYUREC.PRG_ID, vbUnicode))
                                                                '完了区分
        Call UniCode_Conv(Y_SYUREC.KAN_KBN, StrConv(O_Y_SYUREC.KAN_KBN, vbUnicode))
                                                                'データ種別
        Call UniCode_Conv(Y_SYUREC.DT_SYU, StrConv(O_Y_SYUREC.DT_SYU, vbUnicode))
                                                                '事業部区分
        Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(O_Y_SYUREC.JGYOBU, vbUnicode))
                                                                '注文区分（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(O_Y_SYUREC.KEY_CYU_KBN, vbUnicode))
                                                                '伝票ＩＤ（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, StrConv(O_Y_SYUREC.KEY_ID_NO, vbUnicode))
                                                                '国内外
        Call UniCode_Conv(Y_SYUREC.NAIGAI, StrConv(O_Y_SYUREC.NAIGAI, vbUnicode))
                                                                '品目番号（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(O_Y_SYUREC.KEY_HIN_NO, vbUnicode))
                                                                '得意先コード（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(O_Y_SYUREC.KEY_MUKE_CODE, vbUnicode))
                                                                '直送先コード（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(O_Y_SYUREC.KEY_SS_CODE, vbUnicode))
                                                                '出荷日付（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(O_Y_SYUREC.KEY_SYUKA_YMD, vbUnicode))
                                                                '事業場
        Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(O_Y_SYUREC.JGYOBA, vbUnicode))
                                                                'データ区分
        Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(O_Y_SYUREC.DATA_KBN, vbUnicode))
                                                                '取引区分
        Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(O_Y_SYUREC.TORI_KBN, vbUnicode))
                                                                'ＩＤ№
        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(O_Y_SYUREC.ID_NO, vbUnicode))

'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")           '会計用事業場ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")           '資産管理用事業場ｺｰﾄﾞ
'------------------------------------------------------------------------------

                                                                '品目番号
        Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(O_Y_SYUREC.HIN_NO, vbUnicode))
                                                                '伝票番号
        Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(O_Y_SYUREC.DEN_NO, vbUnicode))
                                                                '出荷数量
        Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(O_Y_SYUREC.SURYO, vbUnicode))
                                                                '得意先コード
        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(O_Y_SYUREC.MUKE_CODE, vbUnicode))
                                                                '在庫収支
        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(O_Y_SYUREC.SYUKO_SYUSI, vbUnicode))

'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")            '資産管理用在庫収支ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")             '補助在庫収支ｺｰﾄﾞ
'------------------------------------------------------------------------------

                                                                '出荷日付
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(O_Y_SYUREC.SYUKA_YMD, vbUnicode))

'--- 追加項目 ------------------------------------------------------------------
                                                                '実際単価
        Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(O_Y_SYUREC.TANKA, vbUnicode))
'------------------------------------------------------------------------------

                                                                'オーダー番号
        Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(O_Y_SYUREC.ODER_NO, vbUnicode))
                                                                'アイテム番号
        Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(O_Y_SYUREC.ITEM_NO, vbUnicode))

'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")               '注文管理番号略号
        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")             '個装形態ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, "")               '出荷予定日
        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")                'ﾛｹｰｼｮﾝ1
        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")                'ﾛｹｰｼｮﾝ2
        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")                'ﾛｹｰｼｮﾝ3
'------------------------------------------------------------------------------

                                                                '得意先名称
        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(O_Y_SYUREC.MUKE_NAME, vbUnicode))
                                                                '注文区分
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(O_Y_SYUREC.CYU_KBN, vbUnicode))
                                                                '注文区分名称
        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(O_Y_SYUREC.CYU_KBN_NAME, vbUnicode))

'''                                                                '輸出出荷検査区分
'''        Call UniCode_Conv(Y_SYUREC.EXPORT_KBN, StrConv(O_Y_SYUREC.EXPORT_KBN, vbUnicode))
'''                                                                '個装ラベル発行区分
'''        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_KBN, StrConv(O_Y_SYUREC.LABEL_ISSUE_KBN, vbUnicode))
'''                                                                '個装ラベル発行単位数
'''        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_UNIT, StrConv(O_Y_SYUREC.LABEL_ISSUE_UNIT, vbUnicode))
'''                                                                '個装ラベル単価表示区分
'''        Call UniCode_Conv(Y_SYUREC.LABEL_TANKA_KBN, StrConv(O_Y_SYUREC.LABEL_TANKA_KBN, vbUnicode))
'''                                                                '単価
'''        Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(O_Y_SYUREC.TANKA, vbUnicode))
'''                                                                '金額
'''        Call UniCode_Conv(Y_SYUREC.KINGAKU, StrConv(O_Y_SYUREC.KINGAKU, vbUnicode))
'''                                                                '備考２
'''        Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(O_Y_SYUREC.BIKOU2, vbUnicode))
'''                                                                'リベート区分
'''        Call UniCode_Conv(Y_SYUREC.REBATE_KBN, StrConv(O_Y_SYUREC.REBATE_KBN, vbUnicode))
'''                                                                '帳端区分
'''        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(O_Y_SYUREC.CHOHA_KBN, vbUnicode))
'''                                                                '値差区分
'''        Call UniCode_Conv(Y_SYUREC.ATAISA_KBN, StrConv(O_Y_SYUREC.ATAISA_KBN, vbUnicode))
'''                                                                '代表機種
'''        Call UniCode_Conv(Y_SYUREC.REP_KISHU, StrConv(O_Y_SYUREC.REP_KISHU, vbUnicode))
'''                                                                'ＮＳ管理番号
'''        Call UniCode_Conv(Y_SYUREC.NS_KANRI_NO, StrConv(O_Y_SYUREC.NS_KANRI_NO, vbUnicode))
'''                                                                'ＭＴＳ部品コード
'''        Call UniCode_Conv(Y_SYUREC.MTS_HIN_CODE, StrConv(O_Y_SYUREC.MTS_HIN_CODE, vbUnicode))
'''                                                                '備考１
'''        Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(O_Y_SYUREC.BIKOU1, vbUnicode))
'''                                                                '直送区分
'''        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(O_Y_SYUREC.CHOKU_KBN, vbUnicode))
'''                                                                'リベート率
'''        Call UniCode_Conv(Y_SYUREC.REBATE_RATE, StrConv(O_Y_SYUREC.REBATE_RATE, vbUnicode))
'''                                                                '品名
'''        Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(O_Y_SYUREC.HIN_NAME, vbUnicode))
'''                                                                '対外事業場
'''        Call UniCode_Conv(Y_SYUREC.JGYOBA_GAI, StrConv(O_Y_SYUREC.JGYOBA_GAI, vbUnicode))
'''                                                                '機種コード
'''        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(O_Y_SYUREC.KISHU_CODE, vbUnicode))
'''                                                                '直送先コード
'''        Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(O_Y_SYUREC.SS_CODE, vbUnicode))


'--- 追加項目 ------------------------------------------------------------------
        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")                 '原産国1
        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")                 '原産国2
        Call UniCode_Conv(Y_SYUREC.BIKOU2, _
                    StrConv(O_Y_SYUREC.BIKOU2, vbUnicode))      '備考2
        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")                 '販売区分
        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, _
                    StrConv(O_Y_SYUREC.CHOKU_KBN, vbUnicode))   '直送指示区分
        Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")              'ﾕﾆｯﾄ修正管理番号
        Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")           '在庫引当順序
        Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")          '合梱管理番号
        Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")              '受注残数量
        Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")              '供給区分
        Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")            '商品化納品在庫収支ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")          '商品化納品資産管理収支ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")           '商品化納品補助収支ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.BIKOU1, _
                    StrConv(O_Y_SYUREC.BIKOU1, vbUnicode))      '備考1
        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, _
                    StrConv(O_Y_SYUREC.CHOHA_KBN, vbUnicode))   '帳端区分
        Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")              '受付品目番号
        Call UniCode_Conv(Y_SYUREC.HIN_NAME, _
                    StrConv(O_Y_SYUREC.HIN_NAME, vbUnicode))    '品名
        Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")          '品目番号変更区分
        Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")         'ﾓｼﾞｭｰﾙ交換区分
        Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")             '残在庫まとめ在庫収支ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")        '残在庫まとめ資産管理収支ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")         '残在庫まとめ補助収支ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")               '指定納期
        Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")        'ｻｰﾋﾞｽ会社管理番号
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, _
                StrConv(O_Y_SYUREC.KISHU_CODE, vbUnicode))      '機種品目ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")              '環境企画部品区分
        Call UniCode_Conv(Y_SYUREC.SS_CODE, _
                    StrConv(O_Y_SYUREC.SS_CODE, vbUnicode))     '直送相手先ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")            '欠品解消区分
'------------------------------------------------------------------------------

                                                                '品番（内部）
        Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(O_Y_SYUREC.HIN_NAI, vbUnicode))
                                                                'ホスト棚番
        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(O_Y_SYUREC.HTANABAN, vbUnicode))
                                                                '出庫表印刷日付
        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, StrConv(O_Y_SYUREC.PRINT_YMD, vbUnicode))
                                                                '完了日付
        Call UniCode_Conv(Y_SYUREC.KAN_YMD, StrConv(O_Y_SYUREC.KAN_YMD, vbUnicode))
                                                                '検品日付
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, StrConv(O_Y_SYUREC.KENPIN_YMD, vbUnicode))
                                                                '特売り区分
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, StrConv(O_Y_SYUREC.TOK_KBN, vbUnicode))
                                                                '実績出庫数
        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(O_Y_SYUREC.JITU_SURYO, vbUnicode))
                                                                '取込み日時
        Call UniCode_Conv(Y_SYUREC.INS_NOW, StrConv(O_Y_SYUREC.INS_NOW, vbUnicode))
                                                                        
        Call UniCode_Conv(Y_SYUREC.FILLER, "")


        Do
            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Cnt(6).Caption = Format(count, "#0")
                    DoEvents
                    Call File_Error(sts, BtOpInsert, "出荷予定")
                    Exit Function
            End Select
        Loop

        com = BtOpGetNext

    Loop

    Cnt(6).Caption = Format(count, "#0")


'---------------------------------------------  終了
Convert_End:
    
    Convert_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)
Dim ans     As Integer
Dim i       As Integer

    Select Case Index

        Case 0      '全選択
            For i = 0 To 6
                Check1(i).Value = 1
            Next i

        Case 1      'ｺﾝﾊﾞｰﾄ開始
            Beep
            ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                Command1(0).Enabled = False
                Command1(1).Enabled = False
                Command1(2).Enabled = False
                DoEvents

                If Convert_Proc() Then
                    Unload Me
                End If
            End If
            MsgBox "終了しました。"
            Unload Me

        Case 2      'ｷｬﾝｾﾙ
            Unload Me

    End Select

End Sub

Private Sub Command2_Click()

Dim yn      As Integer


    If Not IsNumeric(Text1(0).Text) Or _
        Not IsNumeric(Text1(1).Text) Or _
        Not IsNumeric(Text1(2).Text) Then
        MsgBox "日付ｴﾗｰ"
        Exit Sub
    End If
    
    Text1(1).Text = Format(CInt(Text1(1).Text), "00")
    Text1(2).Text = Format(CInt(Text1(2).Text), "00")
    
    yn = MsgBox("移動歴削除を行います？", vbYesNo + vbDefaultButton2, "注意！！")

    If yn = vbYes Then
        If IDO_DELETE_PROC() Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_DblClick()
    PrintForm
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

                                '削除済み出荷予定ＯＰＥＮ
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）削除済み出荷予定ＯＰＥＮ
    If O_DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '発番マスタＯＰＥＮ
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）発番マスタＯＰＥＮ
    If O_HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）在庫移動歴ＯＰＥＮ
    If O_IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '作業実績ログＯＰＥＮ
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）作業実績ログＯＰＥＮ
    If O_P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '商品化指図（子）ＯＰＥＮ
    If P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）商品化指図（子）ＯＰＥＮ
    If O_P_SSHIJI_K_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷予定ＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）入荷予定ＯＰＥＮ
    If O_Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）出荷予定ＯＰＥＮ
    If O_Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

                                            '削除済み出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "削除済み出荷予定")
        End If
    End If
                                            '(旧)削除済み出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, O_DEL_SYU_POS, O_DEL_SYUREC, Len(O_DEL_SYUREC), K0_O_DEL_SYU, Len(K0_O_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "（旧）削除済み出荷予定")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '(旧)在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, O_HATUBAN_POS, O_HATUBANREC, Len(O_HATUBANREC), K0_O_HATUBAN, Len(K0_O_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫データ")
        End If
    End If
    
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '(旧)在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, O_IDO_POS, O_IDOREC, Len(O_IDOREC), K0_O_IDO, Len(K0_O_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫移動歴")
        End If
    End If
                                            '作業実績ログＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "作業実績ログ")
        End If
    End If
                                            '(旧)作業実績ログＣＬＯＳＥ
    sts = BTRV(BtOpClose, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), K0_O_Y_NYU, Len(K0_O_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)作業実績ログ")
        End If
    End If
                                            '商品化指図（子）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "商品化指図（子）")
        End If
    End If
                                            '(旧)商品化指図（子）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), K0_O_Y_SYU, Len(K0_O_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)商品化指図（子）")
        End If
    End If
                                            '入荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
                                            '(旧)入荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, O_Y_NYU_POS, O_Y_NYUREC, Len(O_Y_NYUREC), K0_O_Y_NYU, Len(K0_O_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)出荷予定データ")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
                                            '(旧)出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, O_Y_SYU_POS, O_Y_SYUREC, Len(O_Y_SYUREC), K0_O_Y_SYU, Len(K0_O_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)出荷予定データ")
        End If
    End If


    sts = BTRV(BtOpReset, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV20060515F = Nothing

    End
End Sub

Private Function IDO_DELETE_PROC() As Integer
    
Dim count           As Long
Dim DISP_INTERVAL   As Long
Dim com             As Integer
Dim sts             As Integer
Dim ans             As Integer
    
Dim DEL_COUNT           As Long
Dim DISP_DEL_INTERVAL   As Long
    
    IDO_DELETE_PROC = True
    
    MsgLab(1) = "在庫移動歴削除処理中！！"
    Me.MousePointer = vbHourglass
    count = 0
    DISP_INTERVAL = 0
    DEL_COUNT = 0
    DISP_DEL_INTERVAL = 0
    Cnt(2).Caption = Format(count, "#0")


    com = BtOpGetFirst
    Do

        DoEvents

        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "在庫移動歴")
                Exit Function
        End Select

        count = count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(count, "#0")
            DISP_INTERVAL = 0
        End If



        If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text1(0).Text & Text1(1).Text & Text1(2).Text) Then
        Else

    
            If Trim(StrConv(IDOREC.TANTO_CODE, vbUnicode)) = "" Then
                DEL_COUNT = DEL_COUNT + 1
                
                DISP_DEL_INTERVAL = DISP_DEL_INTERVAL + 1
                If DISP_DEL_INTERVAL = 100 Then
                    Label2.Caption = Format(DEL_COUNT, "#0")
                    DISP_DEL_INTERVAL = 0
                End If
                
                
                
        
        
                Do
                    sts = BTRV(BtOpDelete, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<IDOREKI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpDelete, "在庫移動歴")
                            Exit Function
                    End Select
                Loop
            End If
        End If
        com = BtOpGetNext

    Loop
    
    Cnt(2).Caption = Format(count, "#0")
    Label2.Caption = Format(DEL_COUNT, "#0")
    
    MsgLab(1) = ""
    Me.MousePointer = vbDefault
    IDO_DELETE_PROC = False

End Function
