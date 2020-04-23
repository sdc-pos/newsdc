VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F1040171 
   Caption         =   "[緊急処理]品目マスタ前月金額入れ替え処理"
   ClientHeight    =   6675
   ClientLeft      =   2025
   ClientTop       =   -3510
   ClientWidth     =   8985
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
   ScaleHeight     =   6675
   ScaleWidth      =   8985
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   495
      Index           =   3
      Left            =   5145
      TabIndex        =   6
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ＬＯＧ"
      Height          =   495
      Index           =   2
      Left            =   1995
      TabIndex        =   5
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "実行"
      Height          =   495
      Index           =   1
      Left            =   420
      TabIndex        =   4
      Top             =   240
      Width           =   960
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   420
      TabIndex        =   3
      Top             =   1800
      Width           =   8310
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8085
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "参照"
      Height          =   495
      Index           =   0
      Left            =   7560
      TabIndex        =   2
      Top             =   1200
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   4845
   End
   Begin VB.Label Label2 
      Height          =   375
      Index           =   1
      Left            =   5565
      TabIndex        =   8
      Top             =   5640
      Width           =   1065
   End
   Begin VB.Label Label2 
      Height          =   375
      Index           =   0
      Left            =   315
      TabIndex        =   7
      Top             =   5640
      Width           =   5160
   End
   Begin VB.Label Label1 
      Caption         =   "取り込みデータ"
      Height          =   255
      Left            =   630
      TabIndex        =   0
      Top             =   1320
      Width           =   1800
   End
End
Attribute VB_Name = "F1040171"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'------------------------------------   'テキスト定義

Private Const ptxTanto_Code% = 0            '担当者コード
Private Const ptxTanto_Name% = 1            '担当者名称
Private Const ptxHin_Gai% = 2               '品番
Private Const ptxHin_Name% = 3              '品名

Private Const ptxST_SOKO% = 4               '標準棚番　 倉庫
Private Const ptxST_RETU% = 5               '標準棚番   列
Private Const ptxST_REN% = 6                '標準棚番　 連
Private Const ptxST_DAN% = 7                '標準棚番　 段

Private Const ptxBEF_SEI_LOT% = 8           '変更前　   ロット数
Private Const ptxBEF_SEI_RATE% = 9          '           分レート
Private Const ptxBEF_S_KOUSU% = 10          '           分レート
Private Const ptxBEF_S_KOUSU_GENKA% = 11    '           (原価)商品化工料
Private Const ptxBEF_S_KOUSU_BAIKA% = 12    '           (売価)商品化工料
Private Const ptxBEF_S_SHIZAI_GENKA% = 13   '           (原価)箱代
Private Const ptxBEF_S_SHIZAI_BAIKA% = 14   '           (売価)箱代

Private Const ptxBEF_S_GAISO_TANKA% = 165   '           外装単価
Private Const ptxBEF_S_PPSC_KAKO_KOSU% = 161 '          PPSC加工単価
Private Const ptxBEF_S_BU_KAKO_KOSU% = 162  '           BU加工単価




Private Const ptxBEF_S_KOUSU_SET_DATE% = 15 '          設定日
Private Const ptxBEF_SEI_TANKA_TANTO% = 16  '          担当者
Private Const ptxBEF_SE_TANKA_MEMO% = 17    '          メモ

Private Const ptxAFT_SEI_LOT% = 18          '変更後　   ロット数
Private Const ptxAFT_SEI_RATE% = 19         '           分レート
Private Const ptxAFT_S_KOUSU% = 20          '           工数
Private Const ptxAFT_S_KOUSU_GENKA% = 21    '           (原価)商品化工料
Private Const ptxAFT_S_KOUSU_BAIKA% = 22    '           (売価)商品化工料
Private Const ptxAFT_S_SHIZAI_GENKA% = 23   '           (原価)箱代
Private Const ptxAFT_S_SHIZAI_BAIKA% = 24   '           (売価)箱代




Private Const ptxAFT_S_GAISO_TANKA% = 166   '           外装単価
Private Const ptxAFT_S_PPSC_KAKO_KOSU% = 163 '          PPSC加工単価
Private Const ptxAFT_S_BU_KAKO_KOSU% = 164  '           BU加工単価


Private Const ptxAFT_S_KOUSU_SET_DATE% = 25 '          設定日
Private Const ptxAFT_SEI_TANKA_TANTO% = 26  '          担当者
Private Const ptxAFT_SE_TANKA_MEMO% = 27    '          メモ


Private Const ptxZEN_AVE% = 28              '月平均出荷数   前年度　平均
Private Const ptxZEN_SYUKAQTY04% = 29       '月平均出荷数   前年度　4月
Private Const ptxZEN_SYUKAQTY05% = 30       '　                     5月
Private Const ptxZEN_SYUKAQTY06% = 31       '　                     6月
Private Const ptxZEN_SYUKAQTY07% = 32       '　                     7月
Private Const ptxZEN_SYUKAQTY08% = 33       '　                     8月
Private Const ptxZEN_SYUKAQTY09% = 34       '　                     9月
Private Const ptxZEN_SYUKAQTY10% = 35       '　                     10月
Private Const ptxZEN_SYUKAQTY11% = 36       '　                     11月
Private Const ptxZEN_SYUKAQTY12% = 37       '　                     12月
Private Const ptxZEN_SYUKAQTY01% = 38       '　                     1月
Private Const ptxZEN_SYUKAQTY02% = 39       '　                     2月
Private Const ptxZEN_SYUKAQTY03% = 40       '　                     3月

Private Const ptxTOU_AVE% = 41              '月平均出荷数   今年度　平均
Private Const ptxTOU_SYUKAQTY04% = 42       '月平均出荷数   今年度　4月
Private Const ptxTOU_SYUKAQTY05% = 43       '　                     5月
Private Const ptxTOU_SYUKAQTY06% = 44       '　                     6月
Private Const ptxTOU_SYUKAQTY07% = 45       '　                     7月
Private Const ptxTOU_SYUKAQTY08% = 46       '　                     8月
Private Const ptxTOU_SYUKAQTY09% = 47       '　                     9月
Private Const ptxTOU_SYUKAQTY10% = 48       '　                     10月
Private Const ptxTOU_SYUKAQTY11% = 49       '　                     11月
Private Const ptxTOU_SYUKAQTY12% = 50       '　                     12月
Private Const ptxTOU_SYUKAQTY01% = 51       '　                     1月
Private Const ptxTOU_SYUKAQTY02% = 52       '　                     2月
Private Const ptxTOU_SYUKAQTY03% = 53       '　                     3月





Private Const ptxBEF_KOUTEI_TANI01% = 54    '前工程01　 単位
Private Const ptxBEF_KOUTEI_QTY01% = 55     '           数量
Private Const ptxBEF_KOUTEI_KOUSU01% = 56   '           工数
Private Const ptxBEF_KOUTEI_TANI02% = 57    '前工程02　 単位
Private Const ptxBEF_KOUTEI_QTY02% = 58     '           数量
Private Const ptxBEF_KOUTEI_KOUSU02% = 59   '           工数
Private Const ptxBEF_KOUTEI_TANI03% = 60    '前工程03　 単位
Private Const ptxBEF_KOUTEI_QTY03% = 61     '           数量
Private Const ptxBEF_KOUTEI_KOUSU03% = 62   '           工数
Private Const ptxBEF_KOUTEI_TANI04% = 63    '前工程04　 単位
Private Const ptxBEF_KOUTEI_QTY04% = 64     '           数量
Private Const ptxBEF_KOUTEI_KOUSU04% = 65   '           工数
Private Const ptxBEF_KOUTEI_TANI05% = 66    '前工程05　 単位
Private Const ptxBEF_KOUTEI_QTY05% = 67     '           数量
Private Const ptxBEF_KOUTEI_KOUSU05% = 68   '           工数
Private Const ptxBEF_KOUTEI_TANI06% = 69    '前工程06　 単位
Private Const ptxBEF_KOUTEI_QTY06% = 70     '           数量
Private Const ptxBEF_KOUTEI_KOUSU06% = 71   '           工数
Private Const ptxBEF_KOUTEI_TANI07% = 72    '前工程07　 単位
Private Const ptxBEF_KOUTEI_QTY07% = 73     '           数量
Private Const ptxBEF_KOUTEI_KOUSU07% = 74   '           工数
Private Const ptxBEF_KOUTEI_TANI08% = 75    '前工程08　 単位
Private Const ptxBEF_KOUTEI_QTY08% = 76     '           数量
Private Const ptxBEF_KOUTEI_KOUSU08% = 77   '           工数
Private Const ptxBEF_KOUTEI_TANI09% = 78    '前工程09　 単位
Private Const ptxBEF_KOUTEI_QTY09% = 79     '           数量
Private Const ptxBEF_KOUTEI_KOUSU09% = 80   '           工数

Private Const ptxBEF_KOUTEI_KEI1% = 81      '前工程計   計

Private Const ptxBEF_KOUTEI_R_RATE% = 82    '前工程計   余裕率

Private Const ptxBEF_KOUTEI_KEI2% = 83      '前工程計   (秒／個)
Private Const ptxBEF_KOUTEI_KEI3% = 84      '前工程計   (分／個)
Private Const ptxBEF_KOUTEI_KEI4% = 85      '前工程計   (円／個)

Private Const ptxMAIN_KOUTEI_TANI01% = 86   '作業工程01 単位
Private Const ptxMAIN_KOUTEI_QTY01% = 87    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU01% = 88  '           工数
Private Const ptxMAIN_KOUTEI_TANI02% = 89   '作業工程02 単位
Private Const ptxMAIN_KOUTEI_QTY02% = 90    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU02% = 91  '           工数
Private Const ptxMAIN_KOUTEI_TANI03% = 92   '作業工程03 単位
Private Const ptxMAIN_KOUTEI_QTY03% = 93    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU03% = 94  '           工数
Private Const ptxMAIN_KOUTEI_TANI04% = 95   '作業工程04 単位
Private Const ptxMAIN_KOUTEI_QTY04% = 96    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU04% = 97  '           工数
Private Const ptxMAIN_KOUTEI_TANI05% = 98   '作業工程05 単位
Private Const ptxMAIN_KOUTEI_QTY05% = 99    '           数量
Private Const ptxMAIN_KOUTEI_KOUSU05% = 100 '           工数
Private Const ptxMAIN_KOUTEI_TANI06% = 101  '作業工程06 単位
Private Const ptxMAIN_KOUTEI_QTY06% = 102   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU06% = 103 '           工数
Private Const ptxMAIN_KOUTEI_TANI07% = 104  '作業工程07 単位
Private Const ptxMAIN_KOUTEI_QTY07% = 105   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU07% = 106 '           工数
Private Const ptxMAIN_KOUTEI_TANI08% = 107  '作業工程08 単位
Private Const ptxMAIN_KOUTEI_QTY08% = 108   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU08% = 109 '           工数
Private Const ptxMAIN_KOUTEI_TANI09% = 110  '作業工程09 単位
Private Const ptxMAIN_KOUTEI_QTY09% = 111   '           数量
Private Const ptxMAIN_KOUTEI_KOUSU09% = 112 '           工数

Private Const ptxMAIN_KOUTEI_KEI1% = 113    '作業工程計 計

Private Const ptxMAIN_KOUTEI_R_RATE% = 114  '作業工程計   余裕率


Private Const ptxMAIN_KOUTEI_KEI2% = 115    '作業工程計  (秒／個)
Private Const ptxMAIN_KOUTEI_KEI3% = 116    '作業工程計  (分／個)
Private Const ptxMAIN_KOUTEI_KEI4% = 117    '作業工程計  (円／個)

Private Const ptxAFT_KOUTEI_TANI01% = 118   '後工程01   単位
Private Const ptxAFT_KOUTEI_QTY01% = 119    '           数量
Private Const ptxAFT_KOUTEI_KOUSU01% = 120  '           工数
Private Const ptxAFT_KOUTEI_TANI02% = 121   '後工程02   単位
Private Const ptxAFT_KOUTEI_QTY02% = 122    '           数量
Private Const ptxAFT_KOUTEI_KOUSU02% = 123  '           工数
Private Const ptxAFT_KOUTEI_TANI03% = 124   '後工程03   単位
Private Const ptxAFT_KOUTEI_QTY03% = 125    '           数量
Private Const ptxAFT_KOUTEI_KOUSU03% = 126  '           工数
Private Const ptxAFT_KOUTEI_TANI04% = 127   '後工程04   単位
Private Const ptxAFT_KOUTEI_QTY04% = 128    '           数量
Private Const ptxAFT_KOUTEI_KOUSU04% = 129  '           工数
Private Const ptxAFT_KOUTEI_TANI05% = 130   '後工程05   単位
Private Const ptxAFT_KOUTEI_QTY05% = 131    '           数量
Private Const ptxAFT_KOUTEI_KOUSU05% = 132  '           工数
Private Const ptxAFT_KOUTEI_TANI06% = 133   '後工程06   単位
Private Const ptxAFT_KOUTEI_QTY06% = 134    '           数量
Private Const ptxAFT_KOUTEI_KOUSU06% = 135  '           工数
Private Const ptxAFT_KOUTEI_TANI07% = 136   '後工程07   単位
Private Const ptxAFT_KOUTEI_QTY07% = 137    '           数量
Private Const ptxAFT_KOUTEI_KOUSU07% = 138  '           工数
Private Const ptxAFT_KOUTEI_TANI08% = 139   '後工程08   単位
Private Const ptxAFT_KOUTEI_QTY08% = 140    '           数量
Private Const ptxAFT_KOUTEI_KOUSU08% = 141  '           工数
Private Const ptxAFT_KOUTEI_TANI09% = 142   '後工程09   単位
Private Const ptxAFT_KOUTEI_QTY09% = 143    '           数量
Private Const ptxAFT_KOUTEI_KOUSU09% = 144  '           工数

Private Const ptxAFT_KOUTEI_KEI1% = 145     '後工程計   計

Private Const ptxAFT_KOUTEI_R_RATE% = 146   '後工程計   余裕率



Private Const ptxAFT_KOUTEI_KEI2% = 147     '後工程計   (秒／個)
Private Const ptxAFT_KOUTEI_KEI3% = 148     '後工程計   (分／個)
Private Const ptxAFT_KOUTEI_KEI4% = 149     '後工程計   (円／個)


Private Const ptxKOUTEI_KEI1% = 150         '工程計   計

Private Const ptxKOUTEI_R_RATE% = 151       '工程計   余裕率


Private Const ptxKOUTEI_KEI2% = 152         '工程計   (秒／個)
Private Const ptxKOUTEI_KEI3% = 153         '工程計   (分／個)
Private Const ptxKOUTEI_KEI4% = 154         '工程計   (円／個)


Private Const ptxS_CLASS_CODE% = 155        '商品化ｸﾗｽ
Private Const ptxF_CLASS_CODE% = 156        '付加ｸﾗｽ
Private Const ptxN_CLASS_CODE% = 157        '内職ｸﾗｽ

Private Const ptxIO_TANKA_No% = 158         '棚区分
Private Const ptxSE_Name% = 159             '棚区分名称





Private Const ptxSHIYOU_NO% = 167           '仕様書№       2009.06.02
Private Const ptxMITSUMORI_KBN% = 168       '見積り区分     2009.06.02
'Private Const ptxTANKA_KIRIKAE_DT% = 169    '単価切替日付   2009.06.02
Private Const ptxKIRIKAE_KBN% = 170         '切替区分       2009.06.02
    







'------2009.07.24
Private Const ptxOLD_S_KOUSU_BAIKA% = 171       ' 旧  (売価)商品化工料
Private Const ptxOLD_S_SHIZAI_BAIKA% = 172      ' 旧  (売価)箱代

Private Const ptxOLD_S_GAISO_TANKA% = 173       ' 旧  外装単価
Private Const ptxOLD_S_PPSC_KAKO_KOSU% = 174    ' 旧  PPSC加工単価
Private Const ptxOLD_S_BU_KAKO_KOSU% = 175      ' 旧  BU加工単価
Private Const ptxTANKA_KIRIKAE_DT% = 176        ' 旧  単価切替日付
'------2009.07.24
Private Const ptxPLUS_KOUSU% = 177              ' プラス分工数  2009.09.17




'------------------------------------   'コンボ定義
Private Const pcmbSHIMUKE% = 0          '仕向け先


'------------------------------------   'リッチテキストボックス定義
Private Const prchBIKOU% = 0            '備考

Private Const prchM_BIKOU% = 1          '見積書備考         2009.06.02



'------------------------------------   '構成品
Private Const pGrdKOUSEI% = 0

Dim KOUSEI      As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row   As Integer                'グリッド最大表示件数

Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 13             '最大列数

Private Const ColKO_JGYOBU% = 0         '事業部
Private Const ColKO_NAIGAI% = 1         '国内外


Private Const ColKO_SYUBETSU% = 2       '種別
Private Const ColKO_HIN_GAI% = 3        '品番
Private Const ColKO_HIN_NAME% = 4       '品名
Private Const ColKO_QTY% = 5            '員数
Private Const ColG_ST_SHITAN% = 6       '仕入＠
Private Const ColG_ST_URITAN% = 7       '売上＠
Private Const ColG_ST_SHIKIN% = 8       '仕入金額
Private Const ColG_ST_URIKIN% = 9       '売上金額
Private Const ColS_KOUSU% = 10          '作業時間
Private Const ColSEI_SYU_KON% = 11      '集合梱包
Private Const ColKO_BIKOU% = 12         '備考


                                        '草津 金額出力用
Private Const ColG_ST_URIKIN_KUSATU% = 13



'-----------------------------------    ドロップダウン
Dim SYUBETSU        As New XArrayDB


'-----------------------------------

Dim KOSOU_KBN       As String * 2       '個装区分
Dim GAISO_KBN       As String * 2       '外装区分


Dim INV_IO_TANKA_No As String * 2       '標準棚未登録時の出庫区分
Dim HIN_INV         As Boolean          '未登録品番の登録可否


Dim KUSATU_F        As Boolean          '対象センター　草津 OR 草津以外


Dim SHIZAI_T        As Variant          '資材対象
Dim DOUKON_T        As Variant          '同梱対象
Dim KAKOU_T         As Variant          '加工対象

Dim BU_T            As Variant          'BU付加対象
Dim PPSC_T          As Variant          'PPSC付加対象

Private Const KUSATU_ETC$ = "その他"


Dim svHin_Gai       As String           '品番
Dim svSHIMUKE_CODE  As String           '仕向け先


Dim FUTAI_KBN       As String * 2       '付帯作業 2009.09.05

'-----------------------------------    ＥＸＣＥＬ 宛名＆住所

Dim EX_NAME1        As String           '宛名１
Dim EX_NAME2        As String           '宛名２

Dim EX_SYAMEI       As String           '自社　名称
Dim EX_ADDR1        As String           '自社　住所１
Dim EX_ADDR2        As String           '自社　住所２


Dim EX_CENTER_NAME  As String           'センター   名称
Dim EX_CENTER_ADDR1 As String           'センター   住所１
Dim EX_CENTER_ADDR2 As String           'センター   住所２

Dim EX_BIKOU1       As String           '備考１
Dim EX_BIKOU2       As String           '備考２



'Dim EX_JIGYOBU      As String

'2009.06.02
Dim EX_SHIZAI_T     As Variant          '資材対象
Dim EX_SHIZAI_F     As Boolean          '資材対象

Dim EX_DOUKON_T     As Variant          '同梱対象
Dim EX_DOUKON_F     As Boolean          '同梱対象

Dim EX_FUKA_T       As Variant          '付加作業
Dim EX_FUKA_F       As Boolean          '付加作業


'2009.06.02

Dim EX_BCR_CODE     As String           'ﾊﾞｰｺｰﾄﾞﾗﾍﾞﾙｺｰﾄﾞ


Dim EXCEL_TEMPLATE  As String           'EXCELﾃﾝﾌﾟﾚｰﾄ

Private Const LAST_UPDATE_DAY$ = "2009.12.17 09:00"








Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer
Dim i       As Integer

Dim MESG    As String

    Select Case Index
        Case 0
            
            CommonDialog1.ShowOpen
            Text1(0).Text = Trim(CommonDialog1.fileName)
        
        
        
        
        Case 1
        
        
        
            ans = MsgBox("品目マスタ前月繰越金額入れ替え処理を実行しますか？", vbYesNo, "確認入力")
            
            If ans = vbYes Then
            
            
            
                If Update_Proc() Then
                    Unload Me
                End If
            
            
            
            End If
        
        
        
        
        
        
        
        
        
        
        
        
        Case 2
        
            
            Label2(0).Caption = "ログ出力中"
            Label2(1).Caption = ""
            
            
            For i = 0 To List1.ListCount - 1
            
            
            
            
                Call Log_Out(LOG_F, List1.List(i))
            
            
            
            
            
            
            
            
            
            Next i
        
        
            Label2(0).Caption = "ログ出力終了"
            Label2(1).Caption = ""
        
        Case 3
            Unload Me
    End Select
                    
    
    






End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer






    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
                                
                                
                                '在庫データＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

    F1040171.Caption = F1040171.Caption & " " & LAST_UPDATE_DAY


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
    
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    F1040171.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040171)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(F1040171)


    F1040171.MousePointer = vbDefault

End Sub




Private Function Update_Proc() As Integer
 
Dim In_File     As String
Dim tmp_File    As String
 
Dim i           As Long
Dim j           As Long
 
Dim rec         As String
 
Dim exl         As Object
 
 
Dim FileNo      As Long
 
Dim In_Rec      As String
Dim In_wk       As Variant
 
Dim sts         As Integer
 
Dim List_Wk     As String
 
 
Dim cnt         As Long
 
    Update_Proc = True
 
    Label2(0).Caption = "データ変換中"
    cnt = 0
    
    In_File = Trim(Text1(0).Text)
    tmp_File = "c:\F104017.txt"
    
    Set exl = CreateObject("Excel.Application")
    '  exl.Application.Visible = True
    
    
    On Error GoTo Error_Proc
    exl.Application.Workbooks.Open fileName:=In_File
    
    FileNo = FreeFile
    Open tmp_File For Output As FileNo
    For j = 2 To 65536
        If exl.Cells(j, 1) = "" Then Exit For
            
            
            cnt = cnt + 1
            Label2(1).Caption = cnt
            DoEvents
            rec = ""
            For i = 1 To 256
                If exl.Cells(j, i) = "" Then Exit For
                rec = rec & exl.Cells(j, i) & vbTab
            Next
            Print #1, rec
    Next
    Close FileNo
    '  exl.Application.DisplayAlerts = False
    exl.Application.Quit
        
        
    cnt = 0
    FileNo = FreeFile
    Open tmp_File For Input As FileNo
    
    
    
    List1.Clear
    Label2(0).Caption = "品目マスタ更新"
    
    Do While Not EOF(FileNo)
        
        DoEvents
        
        Line Input #FileNo, In_Rec
        
        In_wk = Split(In_Rec, vbTab, -1)
    
        cnt = cnt + 1
        Label2(1).Caption = cnt
        DoEvents
    
        If UBound(In_wk) < 7 Then
        Else
        
            If In_wk(7) <> "*" Then
            Else
        
                If IsNumeric(In_wk(5)) Then
        
        
        
                    
            
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(CStr(In_wk(0))))
                    
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(CStr(In_wk(1))))
            
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Trim(CStr(In_wk(2))))
            
            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            
                            
                            If Val(StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode)) = Val(CLng(Trim(CStr(In_wk(4))))) Then
                                            
                                            
                                List_Wk = CStr(In_wk(0)) & " " & CStr(In_wk(1)) & " " & CStr(In_wk(2)) & " " & StrConv(ITEMREC.G_ZEN_ZAIKO_KIN, vbUnicode) & "->" & Format(CLng(CStr(In_wk(5))), "00000000000")
                                
                            
                                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(CLng(Trim(CStr(In_wk(5)))), "00000000000"))
                            
                            
                            
                                List1.AddItem List_Wk
                                DoEvents
                            
                                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts Then
                                    Exit Function
                                End If
                        
                        
                            End If
                        
                        Case BtErrKeyNotFound
                        Case Else
                            Exit Function
                    End Select
            
            
                End If
            End If
        
        
        
        
        
        
        
        
        
        End If
    Loop
    
    Close FileNo
        
        
        
        
        
        
        
    Kill tmp_File
        

    Label2(0).Caption = "処理終了"
    Label2(1).Caption = ""



    Update_Proc = False
    
    Exit Function
Error_Proc:
 
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
Dim ans     As Integer
    
    
    Select Case Err.Number
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("ドライブを確認して下さい", vbYesNo + vbExclamation + vbDefaultButton1, "確認入力")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("ドライブまたはパスが見つかりません" & In_File, vbExclamation)
        Case ErrNotFound, 1004
            Beep
            ans = MsgBox("ファイルが見つかりません" & In_File, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("エラー  [" & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select


End Function
