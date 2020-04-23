VERSION 5.00
Begin VB.Form F1020101 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  '可変ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "入出荷予定データ取込み '2019/12/13 滋賀DC 収支R8対応 "
   ClientHeight    =   4170
   ClientLeft      =   1920
   ClientTop       =   2280
   ClientWidth     =   8580
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.ListBox LBox_Hin 
      Height          =   300
      Left            =   1560
      TabIndex        =   25
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6360
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   23
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   21
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   20
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5760
      TabIndex        =   19
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5760
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "／"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '右揃え
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
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
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "F1020101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WS_NO       As String               'ﾜｰｸｽﾃｰｼｮﾝ番号

Private FileName    As String               'テキストファイル名
Private FileNo      As Integer              'ファイル№


Private ER_IN_FileName As String            'ｴﾗｰ　テキストファイル名    2015.11.19
Private ER_IN_FileNo   As Integer           'ｴﾗｰ　ファイル№            2015.11.19
Private ER_OUT_FileName As String           'ｴﾗｰ　テキストファイル名    2015.11.19
Private ER_OUT_FileNo   As Integer          'ｴﾗｰ　ファイル№            2015.11.19

Private TP_IN_FileName As String            'ｴﾗｰ　テキストファイル名    2015.11.19
Private TP_IN_FileNo   As Integer           'ｴﾗｰ　ファイル№            2015.11.19
Private TP_OUT_FileName As String           'ｴﾗｰ　テキストファイル名    2015.11.19
Private TP_OUT_FileNo   As Integer          'ｴﾗｰ　ファイル№            2015.11.19


Private KASO_NYUKA_SOKO      As String * 2  '仮想入荷倉庫番号
Private KASO_SMODOSHI_SOKO   As String * 2  '仮想支給戻し倉庫番号

Private Proc_F      As Integer              '品番＆在庫有無　判定フラグ
Private Last_Proc_F As Integer              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有無フラグ
                                            
Private Type YUKO_SOKO_TBL                  '有効ﾎｽﾄ倉庫取り込みテーブル
    HS_SOKO             As String * 3
    NAIGAI              As String * 1
End Type

Dim Soko_T()            As YUKO_SOKO_TBL  '倉庫情報

'-                                          2005.12.30
Private Type SHIMUKE_TBL
    SHIMUKE_CODE            As String * 2   '仕向け先
    JGYOBU                  As String * 1   '事業部
    NAIGAI                  As String * 1   '国内外
End Type

Private SHIMUKE_T()         As SHIMUKE_TBL

Private SHIMUKE_Flg         As Boolean

'-                                          2005.12.30


Private HS_IN_SIJ   As String               '入庫データファイル名
Private HS_OUT_SIJ  As String               '出庫データファイル名
Private New_HS_OUT_SIJ  As String           '新ﾚｲｱｳﾄ出庫データファイル名2006.05.23


Private In_Cnt      As Integer              'データ読み込み件数
Private Out_Cnt     As Integer              'データ出力件数

Private Const In_Mode% = 1                  '入荷処理
Private Const Out_Mode% = 2                 '出荷処理

                                            
Dim NormalFont As New StdFont               '印刷フォント

Private Const LMAX% = 46                    '頁内最大行数
Private Const MGN_L% = 1                    '明細印刷開始桁位置（１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）
Private Pdate As String                     '印刷開始日付（ﾍｯﾀﾞｰ用）
Private Ptime As String                     '印刷開始時刻（ﾍｯﾀﾞｰ用）


Private Const NAI_CHANGE% = 1
Private Const GAI_CHANGE% = 2
Private Const NOT_GAI_CHANGE% = 3


Private ETC_MTS_NAI As String * 8             'その他向け先(国内)
Private ETC_SS_NAI  As String * 8             'その他直送先(国内)

Private ETC_MTS_GAI As String * 8             'その他向け先(海外)
Private ETC_SS_GAI  As String * 8             'その他直送先(海外)

Dim DUP_SYUKA_DATA  As String                 '出荷データフルパス


                                        
Dim MyCenter        As String

Private MENU_NO     As String * 2       '実績ログ出力用ﾒﾆｭｰ№   2007.11.06

Dim Err_FLg         As Boolean          '2008.10.07

Dim TANA_SPACE      As Boolean          '2009.03.07

Dim KAMOKU_FURIKAE      As String * 2       '科目振替要因 2009.06.26

'2010.07.20 ▽
'Private Const GENSANKOKU_ON% = 1
'Private Const GENSANKOKU_OFF% = 0
'2010.07.20 △


'商品化計画支援 2011.07.07
Dim NOT_Hin_Name    As Variant          '除外品名
Dim NOT_Hin_Name_F  As Boolean          '除外品名有無
'商品化計画支援 2011.07.07


Dim GOODS_F         As String * 1       '商品化有無　ﾃﾞﾌｫﾙﾄ 2012.12.20




Dim GENSAN_T()      As String * 1       '原産国更新有無 2016.12.28


'Private Const Last_Update_Day$ = "[F102010] 2019.03.06 11:55"
'Private Const Last_Update_Day$ = "[F102010] 2019.04.15 09:30"
Private Const Last_Update_Day$ = "[F102010] 2019.12.13 17:00 エアコン 収支R8取込対応"






Private Function Nyuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   「入荷予定データ」更新処理
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Skip_Flg    As Boolean
    
Dim WK_Y_QTY    As Long     '出荷数ワーク
Dim WK_Qty      As Long     '前借残ワーク
Dim WK_E_QTY    As Long     '先行出荷数ワーク

Dim SUMI_QTY    As Long     '商品化済みとして登録
Dim MI_QTY      As Long     '未商品として登録



Dim WORK_SOKO   As String * 2
    
Dim sts         As Integer
Dim ans         As Integer
Dim Not_SHUSI   As Boolean
    
Dim wkText      As String
Dim Length      As Integer
    
    
Dim NAIGAI      As String * 1   '2007.06.15
    
    
Dim TEXT_NO     As String * 9           'ﾃｷｽﾄ№
Dim JGYOBU_Code As String * 1           '事業部区分
Dim CYOK_KBN    As String * 1           '直送区分
Dim DEN_DT      As String * 8           '伝票日付
Dim IO_KBN      As String * 1           '入出庫区分
Dim PM_KBN      As String * 1           '赤黒区分
Dim DEN_SYU     As String * 1           '伝票種別
Dim DEN_NO      As String * 6           '伝票№
Dim CYU_KBN     As String * 1           '注文区分
'Dim HIN_GAI     As String * 13          '品番（外部）  '2016.04.26
'Dim HIN_NAI     As String * 13          '品番（内部）  '2016.04.26
Dim HIN_GAI     As String * 20          '品番（外部）   '2016.04.26
Dim HIN_NAI     As String * 20          '品番（内部）   '2016.04.26
Dim HIN_NAME    As String * 25          '品名
Dim YOTEI_QTY   As String * 6           '数量
Dim YOSAN_FROM  As String * 5           '予算単位（元）
Dim YOSAN_TO    As String * 5           '予算単位（先）
Dim HOST_SOKO   As String * 8           '倉庫区分（ﾎｽﾄ）
Dim HOST_TANA   As String * 8           '棚番（ﾎｽﾄ）
Dim SYUK_CODE   As String * 5           '支給先／出荷先
Dim SYUK_NAME   As String * 20          '支給先／出荷先名
Dim REC_END     As String * 1           'ﾚｺｰﾄﾞ終端ﾏｰｸ(@)
    
    
    
Dim GENSANKOKU          As String * 20  '原産国名
Dim GEN_GENSANKOKU      As String * 20  '現物表示原産国名
Dim SHIIRE_WORK_CENTER  As String * 8   '資材仕入先ﾜｰｸｾﾝﾀｰ
Dim KANKYO_KBN          As String * 3   '環境種類区分
Dim KANKYO_KBN_ST       As String * 8   '環境種類区分適用開始
Dim KANKYO_KBN_SURYO    As String * 10  '環境種類区分数量
Dim ID_NO2              As String * 12  'ID_NO
Dim AITESAKI_CODE       As String * 16  '相手先ｺｰﾄﾞ
Dim JYUCHU_YMD          As String * 8   '受注年月日
Dim SHITEI_NOUKI_YMD    As String * 8   '指定納期年月日
    
    
    
    
Dim GENSANKOKU_CHG_F    As Boolean      '原産国変更F
    
    
'出荷予定 編集前処理 ################################################################# 2005/05/16 Add ↓
Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
Dim INS_NOW         As String * 14
Dim wkStr           As String
    
Dim wkMUKE_CODE     As String
    
'2010.11.01
Dim DUP_FLG         As Boolean

'2011.01.19
Dim Loop_Cnt        As Integer


'2011.03.23
Dim MOTO_TEXT_NO    As String * 9

    
Dim Rec_LENG        As Long         '2016.04.19
    
    
Dim MAEGARI_FLG     As Boolean      '2018.11.16
    
    
    Fast_Flg = True

    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
'#################################################################################### 2005/05/16 Add ↑
    
    Nyuka_Update_Proc = True


    Do Until EOF(FileNo)
        Line Input #FileNo, wkText
    
    
    
    
    
'        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 And LenB(StrConv(wkText, vbFromUnicode)) <> 251 Then    '2016.04.26
        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 And LenB(StrConv(wkText, vbFromUnicode)) <> 265 Then     '2016.04.26
            
'            Call NG_File_Make_Proc
             Err_FLg = True
           
    Call LOG_OUT(LOG_F, wkText)
            Exit Do
        End If
    
    
        Rec_LENG = LenB(StrConv(wkText, vbFromUnicode)) '2016.04.19
    
    
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
    
    
    
        MAEGARI_FLG = False     '2018.11.16
    
                                                                    'ﾃｷｽﾄ№
        Length = 1
        TEXT_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(TEXT_NO)), vbUnicode)
                                                                    '事業部区分
        Length = Length + Len(TEXT_NO)
        JGYOBU_Code = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JGYOBU_Code)), vbUnicode)
                                                                    '直送区分
        Length = Length + Len(JGYOBU_Code)
        CYOK_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYOK_KBN)), vbUnicode)
                                                                    '伝票日付
        Length = Length + Len(CYOK_KBN)
        DEN_DT = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_DT)), vbUnicode)
                                                                    '入出庫区分
        Length = Length + Len(DEN_DT)
        IO_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(IO_KBN)), vbUnicode)
                                                                    '赤黒区分
        Length = Length + Len(IO_KBN)
        PM_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(PM_KBN)), vbUnicode)
                                                                    '伝票種別
        Length = Length + Len(PM_KBN)
        DEN_SYU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_SYU)), vbUnicode)
                                                                    '伝票№
        Length = Length + Len(DEN_SYU)
        DEN_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_NO)), vbUnicode)
                                                                    '注文区分
        Length = Length + Len(DEN_NO)
        CYU_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYU_KBN)), vbUnicode)
                                                                    '品番（外部）
        Length = Length + Len(CYU_KBN)
        HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI)), vbUnicode)
                                                                    '品番（内部）
        Length = Length + Len(HIN_GAI)
        HIN_NAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAI)), vbUnicode)
                                                                    '品名
        Length = Length + Len(HIN_NAI)
        HIN_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAME)), vbUnicode)
                                                                    '数量
        Length = Length + Len(HIN_NAME)
        YOTEI_QTY = Trim(StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOTEI_QTY)), vbUnicode))
                                                                    '予算単位（元）
        Length = Length + Len(YOTEI_QTY)
        YOSAN_FROM = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_FROM)), vbUnicode)
                                                                    '予算単位（先）
        Length = Length + Len(YOSAN_FROM)
        YOSAN_TO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_TO)), vbUnicode)
                                                                    '倉庫区分（ﾎｽﾄ）
        Length = Length + Len(YOSAN_TO)
        HOST_SOKO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_SOKO)), vbUnicode)
                                                                    '棚番（ﾎｽﾄ）
        Length = Length + Len(HOST_SOKO)
        HOST_TANA = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_TANA)), vbUnicode)
                                                                    '支給先／出荷先
        Length = Length + Len(HOST_TANA)
        SYUK_CODE = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_CODE)), vbUnicode)
                                                                    '支給先／出荷先名
        Length = Length + Len(SYUK_CODE)
        SYUK_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_NAME)), vbUnicode)
    
    
'        If LenB(StrConv(wkText, vbFromUnicode)) = 251 Then     '2016.04.26
        If LenB(StrConv(wkText, vbFromUnicode)) = 265 Then      '2016.04.26
                                                                    '原産国
            Length = Length + Len(SYUK_NAME)
            GENSANKOKU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(GENSANKOKU)), vbUnicode)
                                                                    '現物表示原産国名
            Length = Length + Len(GENSANKOKU)
            GEN_GENSANKOKU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(GEN_GENSANKOKU)), vbUnicode)
                                                                    '資材仕入先ﾜｰｸｾﾝﾀｰ
            Length = Length + Len(GEN_GENSANKOKU)
            SHIIRE_WORK_CENTER = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SHIIRE_WORK_CENTER)), vbUnicode)
                                                                    '環境種類区分
            Length = Length + Len(SHIIRE_WORK_CENTER)
            KANKYO_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(KANKYO_KBN)), vbUnicode)
                                                                    '環境種類区分適用開始
            Length = Length + Len(KANKYO_KBN)
            KANKYO_KBN_ST = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(KANKYO_KBN_ST)), vbUnicode)
                                                                    '環境種類区分数量
            Length = Length + Len(KANKYO_KBN_ST)
            KANKYO_KBN_SURYO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(KANKYO_KBN_SURYO)), vbUnicode)
                                                                    'ID_NO
            Length = Length + Len(KANKYO_KBN_SURYO)
            ID_NO2 = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(ID_NO2)), vbUnicode)
                                                                    '相手先
            Length = Length + Len(ID_NO2)
            AITESAKI_CODE = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(AITESAKI_CODE)), vbUnicode)
                                                                    '受注年月日
            Length = Length + Len(AITESAKI_CODE)
            JYUCHU_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JYUCHU_YMD)), vbUnicode)
                                                                    '指定納期年月日
            Length = Length + Len(JYUCHU_YMD)
            SHITEI_NOUKI_YMD = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SHITEI_NOUKI_YMD)), vbUnicode)
        
        
        
        Else
            
            GENSANKOKU = ""             '原産国名
            GEN_GENSANKOKU = ""         '現物表示原産国名
            SHIIRE_WORK_CENTER = ""     '資材仕入先ﾜｰｸｾﾝﾀｰ
            KANKYO_KBN = ""             '環境種類区分
            KANKYO_KBN_ST = ""          '環境種類区分適用開始
            KANKYO_KBN_SURYO = ""       '環境種類区分数量
            ID_NO2 = ""                 'ID_NO
            AITESAKI_CODE = ""          '相手先ｺｰﾄﾞ
            JYUCHU_YMD = ""             '受注年月日
            SHITEI_NOUKI_YMD = ""       '指定納期年月日
        End If
'        Length = 1
'        TEXT_NO = Mid(wkText, Length, Len(TEXT_NO))                 'ﾃｷｽﾄ№
        
'        Length = Length + Len(TEXT_NO)
'        JGYOBU_Code = Mid(wkText, Length, Len(JGYOBU_Code))         '事業部区分
    
'        Length = Length + Len(JGYOBU_Code)
'        CYOK_KBN = Mid(wkText, Length, Len(CYOK_KBN))               '直送区分
    
'        Length = Length + Len(CYOK_KBN)
'        DEN_DT = Mid(wkText, Length, Len(DEN_DT))                   '伝票日付
    
'        Length = Length + Len(DEN_DT)
'        IO_KBN = Mid(wkText, Length, Len(IO_KBN))                   '入出庫区分
    
'        Length = Length + Len(IO_KBN)
'        PM_KBN = Mid(wkText, Length, Len(PM_KBN))                   '赤黒区分
    
'        Length = Length + Len(PM_KBN)
'        DEN_SYU = Mid(wkText, Length, Len(DEN_SYU))                 '伝票種別
    
'        Length = Length + Len(DEN_SYU)
'        DEN_NO = Mid(wkText, Length, Len(DEN_NO))                   '伝票№
    
'        Length = Length + Len(DEN_NO)
'        CYU_KBN = Mid(wkText, Length, Len(CYU_KBN))                 '注文区分
    
'        Length = Length + Len(CYU_KBN)
'        HIN_GAI = Mid(wkText, Length, Len(HIN_GAI))                 '品番（外部）
    
'        Length = Length + Len(HIN_GAI)
'        HIN_NAI = Mid(wkText, Length, Len(HIN_NAI))                 '品番（内部）
    
'        Length = Length + Len(HIN_NAI)
'        HIN_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAME)), vbUnicode)             '品名
    
'        Length = Length + Len(HIN_NAME)
'        YOTEI_QTY = Trim(Mid(wkText, Length, Len(YOTEI_QTY)))       '数量
    
'        Length = Length + Len(YOTEI_QTY)
'        YOSAN_FROM = Mid(wkText, Length, Len(YOSAN_FROM))           '予算単位（元）
    
'        Length = Length + Len(YOSAN_FROM)
'        YOSAN_TO = Mid(wkText, Length, Len(YOSAN_TO))               '予算単位（先）
    
'        Length = Length + Len(YOSAN_TO)
'        HOST_SOKO = Mid(wkText, Length, Len(HOST_SOKO))             '倉庫区分（ﾎｽﾄ）
    
'        Length = Length + Len(HOST_SOKO)
'        HOST_TANA = Mid(wkText, Length, Len(HOST_TANA))             '棚番（ﾎｽﾄ）
        
'        Length = Length + Len(HOST_TANA)
'        SYUK_CODE = Mid(wkText, Length, Len(SYUK_CODE))             '支給先／出荷先
        
'        Length = Length + Len(SYUK_CODE)
'        SYUK_NAME = Mid(wkText, Length, Len(SYUK_NAME))             '支給先／出荷先名
    
    
    
    
    
    
        Skip_Flg = True
        Not_SHUSI = False
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If JGYOBU = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(HOST_SOKO) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
        If Skip_Flg Then
            Not_SHUSI = True
        End If
    
    
    
    
    
    
    
    
'-----------------------------------------  照合用入荷予定の出力処理    2007.06.15
        '照合用入荷予定重複チェック
        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
        
'2010.11.01
        DUP_FLG = False
'2010.11.01
        
        
        sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
'2010.11.01
                DUP_FLG = True
'2010.11.01
            
            Case BtErrKeyNotFound
            Case Else
                'Call File_Error(sts, BtOpGetEqual, "照合用入荷予定", 0)                '2016.06.23
                Call File_Error(sts, BtOpGetEqual, "照合用入荷予定", 1, Y_GLICS_ID)     '2016.06.23
'                Exit Function      '2015.11.19
                GoTo Abort_Tran     '2015.11.19
        End Select



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 発生元によるデータ確認  2011.03.23
        MOTO_TEXT_NO = ""
        
        If DUP_FLG Then
            If StrConv(App.EXEName, vbUpperCase) = Trim(StrConv(Y_GLICSREC.MOTO_PROG_ID, vbUnicode)) Then
            Else
                If Trim(TEXT_NO) = Trim(StrConv(Y_GLICSREC.MOTO_TEXT_NO, vbUnicode)) Then
                Else
                    MOTO_TEXT_NO = TEXT_NO
                    Mid(TEXT_NO, 5, 1) = "A"
                    DUP_FLG = False
                End If
            
            End If
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 発生元によるデータ確認  2011.03.23

















        If Not_SHUSI Then
            NAIGAI = "1"
        Else
            NAIGAI = Soko_T(i, j).NAIGAI
        End If

'        If sts = BtErrKeyNotFound Then
        If Not DUP_FLG Then

'            If Y_GLICS_PUT_PROC(JGYOBU, NAIGAI, INS_NOW, _
'                                TEXT_NO, _
'                                JGYOBU_Code, _
'                                CYOK_KBN, _
'                                DEN_DT, _
'                                IO_KBN, _
'                                PM_KBN, _
'                                DEN_SYU, _
'                                DEN_NO, _
'                                CYU_KBN, _
'                                HIN_GAI, _
'                                HIN_NAI, _
'                                HIN_NAME, _
'                                YOTEI_QTY, _
'                                YOSAN_FROM, _
'                                YOSAN_TO, _
'                                HOST_SOKO, _
'                                HOST_TANA, _
'                                SYUK_CODE, _
'                                SYUK_NAME) Then
                
                
                
    '2010.07.20 ▽
            If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                If Item_Check_Proc(In_Mode, JGYOBU, "1", HIN_GAI, HIN_NAI, HIN_NAME, GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO) Then
'                    GoTo Abort_Tran
'                    Exit Function      '2015.11.19
                    GoTo Abort_Tran     '2015.11.19
                End If
            End If
    '2010.07.20 △
                

                
                
'''''''''''''''''''''2011.03.23 引数追加
'            If Y_GLICS_PUT_PROC(JGYOBU, NAIGAI, INS_NOW, _
'                                TEXT_NO, _
'                                JGYOBU_Code, _
'                                CYOK_KBN, _
'                                DEN_DT, _
'                                IO_KBN, _
'                                PM_KBN, _
'                                DEN_SYU, _
'                                DEN_NO, _
'                                CYU_KBN, _
'                                HIN_GAI, _
'                                HIN_NAI, _
'                                HIN_NAME, _
'                                YOTEI_QTY, _
'                                YOSAN_FROM, _
'                                YOSAN_TO, _
'                                HOST_SOKO, _
'                                HOST_TANA, _
'                                SYUK_CODE, _
'                                SYUK_NAME, _
'                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD) Then
            If Y_GLICS_PUT_PROC(JGYOBU, NAIGAI, INS_NOW, _
                                TEXT_NO, _
                                JGYOBU_Code, _
                                CYOK_KBN, _
                                DEN_DT, _
                                IO_KBN, _
                                PM_KBN, _
                                DEN_SYU, _
                                DEN_NO, _
                                CYU_KBN, _
                                HIN_GAI, _
                                HIN_NAI, _
                                HIN_NAME, _
                                YOTEI_QTY, _
                                YOSAN_FROM, _
                                YOSAN_TO, _
                                HOST_SOKO, _
                                HOST_TANA, _
                                SYUK_CODE, _
                                SYUK_NAME, _
                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD, MOTO_TEXT_NO) Then
'''''''''''''''''''''2011.03.23 引数追加
'               Exit Function      '2015.11.19
                GoTo Abort_Tran     '2015.11.19
            End If

        End If



'-----------------------------------------  照合用入荷予定の出力処理    2007.06.15
    
    
    
    
    
    
    
    
    
        If IO_KBN <> "1" Then
            
            If IO_KBN = "4" And Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" And Trim(HOST_SOKO) = "11B" Then
            Else
                Skip_Flg = True
            End If
        End If
    
    
        If PM_KBN = "-" Then
            Skip_Flg = True
        End If
    
        'NOPOS  2006.05.01
        If Trim(DEN_NO) = "NOPOS" Then
            Skip_Flg = True
        End If
    
        '予算元＝36003除外  2006.07.15
        If Trim(YOSAN_FROM) = "36003" Then
            Skip_Flg = True
        End If
    
        '予算元＝PP除外  2008.01.10
        If Left(YOSAN_FROM, 2) = "PP" Then
            
            If Trim(YOSAN_FROM) = "PPP4" And JGYOBU = SHOKUSEN Then     '2017.02.17
            Else                                                        '2017.02.17
                Skip_Flg = True
            End If                                                      '2017.02.17
        End If
    
    
    
    
        WORK_SOKO = KASO_NYUKA_SOKO
    
    
    
        Select Case JGYOBU
            
''            Case SENTAKU                        '洗濯機
''
''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
''                    Skip_Flg = True
''                End If
''
''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
''                    Skip_Flg = True
''                End If
                            
            
            
            
            Case SOJIKI                         '掃除機
                
            
                If Left(YOSAN_FROM, 2) = "KM" Then
                    Skip_Flg = True
                End If

                If Left(YOSAN_FROM, 2) = "KK" Then
                    Skip_Flg = True
                End If

                If Left(YOSAN_FROM, 2) = "GG" Then
                    Skip_Flg = True
                End If

                If Left(YOSAN_FROM, 2) = "SS" Then
                    Skip_Flg = True
                End If

                '2005.04.07 収支追加
                If Left(YOSAN_FROM, 5) = "0090K" Then
                    Skip_Flg = True
                End If
                '2006.07.27 収支追加
                If Left(YOSAN_FROM, 5) = "0092H" Then
                    Skip_Flg = True
                End If
                '2006.07.27 収支追加
                If Left(YOSAN_FROM, 2) = "AA" Then
                    Skip_Flg = True
                End If
            
                '2009.08.28 収支追加
                If Left(YOSAN_FROM, 2) = "ZZ" Then
                    Skip_Flg = True
                End If
            
            
            
            
                If Trim(YOSAN_FROM) <> "91H" Then
                    WORK_SOKO = KASO_SMODOSHI_SOKO
                End If
            
            
            
            Case DENKA, SUIHAN, SENTAKU, BLBU        '電化、炊飯、洗濯機（アイロン）    BLBU追加 2012.04.06
            
            
                Select Case MyCenter
                    
                    Case "O"
                
                        If Left(YOSAN_FROM, 2) = "01" Then
                            Skip_Flg = True
                        End If
                    
                        If Left(YOSAN_FROM, 3) = "H33" Then    '2004.07.16
                            Skip_Flg = True
                        End If
                        If Left(YOSAN_FROM, 3) = "H22" Then    '2004.07.16
                            Skip_Flg = True
                        End If
        
                        If Left(YOSAN_FROM, 2) = "05" Then
                            Skip_Flg = True
                        End If
        
                        '2006.08.17
                        If Left(YOSAN_FROM, 2) = "08" Then
                            Skip_Flg = True
                        End If
                        
                        '2006.10.13 電化調理は予算元="02"のみ対象
                        If JGYOBU = DENKA Then
                            
                            '2008.01.07 "02"-->"0201"に変更 2008.01.08 Left(Trim(YOSAN_FROM), 2) <> "02" に変更
                            If Left(Trim(YOSAN_FROM), 2) <> "02" And _
                            Trim(YOSAN_FROM) <> "G11" And _
                            Trim(YOSAN_FROM) <> "G22" And _
                            Trim(YOSAN_FROM) <> "KA01" Then         '2012.08.31 KA01 追加
                                Skip_Flg = True
                            End If
                        End If
        
        
        
        
                        '2012.08.31
                        
        
        
        
                        '2006.11.22 炊飯/アイロンの除外条件追加 2012.04.06 BLBU追加
                        If (JGYOBU = SUIHAN Or _
                            JGYOBU = SENTAKU Or _
                            JGYOBU = BLBU) Then
                            If (Left(YOSAN_FROM, 2) = "P3" Or _
                                Left(YOSAN_FROM, 2) = "S3") Then
                                Skip_Flg = True
                            End If
                        End If
        
        
        
                        '2007.10.25 炊飯/アイロンの除外条件追加 2012.04.06 BLBU追加
                        If (JGYOBU = SUIHAN Or _
                            JGYOBU = SENTAKU Or _
                            JGYOBU = BLBU) Then
                            If Left(YOSAN_FROM, 2) = "RO" Then
                                Skip_Flg = True
                            End If
                        End If
        
                        '2007.12.06 炊飯/アイロンの除外条件追加  2012.04.06 BLBU追加
                        If (JGYOBU = SUIHAN Or _
                            JGYOBU = SENTAKU Or _
                            JGYOBU = BLBU) Then
                            If Left(YOSAN_FROM, 2) = "07" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
                        '2008.06.26 炊飯の除外条件追加 2012.04.06 BLBU追加
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "04" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
                        '2008.10.14 炊飯の除外条件追加 2012.04.06 BLBU追加
        
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "NC" Or Left(YOSAN_FROM, 2) = "99" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
        
                        '2016.06.17 炊飯,BLBUの除外条件追加
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "RX" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
                        '2017.07.22 炊飯,BLBUの除外条件追加
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "RZ" Then
                                Skip_Flg = True
                            End If
                        End If
        
        
        
                        If Left(YOSAN_FROM, 3) = "G22" Then
                            WORK_SOKO = "80"
                        End If
        
                        If Left(YOSAN_FROM, 3) = "G11" Then
                            WORK_SOKO = "81"
                        End If
                    
                        '2006.04.29用
                        If Left(YOSAN_FROM, 2) = "S1" And _
                            Left(YOSAN_TO, 2) = "S3" Then
                            WORK_SOKO = "87"
                        End If
                        '2006.05.01
                        If Trim(DEN_NO) = "POS87" Then
                            WORK_SOKO = "87"
                        End If
                
                                
                
                
                
                
                
                
                        '2008.06.26 炊飯の倉庫番号の設定追加  2012.04.06 BLBU追加
                        If JGYOBU = SUIHAN Or JGYOBU = BLBU Then
                            If Left(YOSAN_FROM, 2) = "02" And Left(YOSAN_TO, 3) = "SDC" Then
                                WORK_SOKO = "90"
                            End If
                        End If
                
                
                
                
                        '2009.06.01 65番倉庫出力追加 2012.04.06 BLBU追加
                        If (JGYOBU = SUIHAN Or JGYOBU = DENKA Or JGYOBU = BLBU) Then
                            If IO_KBN = "4" Then
                                If Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" Then
                                    
                                    If Trim(HOST_SOKO) = "11B" Then
                                        WORK_SOKO = "65"
                                    End If
                                End If
                            End If
                        End If
                        
                        
                
                
                
                
                
                
                
                
                    Case "F"
            
            
            
                        If Left(YOSAN_FROM, 2) = "P2" Then
                            Skip_Flg = True
                        End If

                        If Left(YOSAN_FROM, 2) = "U2" Then      '2008.01.11
                            Skip_Flg = True
                        End If


                        If Left(YOSAN_FROM, 3) <> "904" Then
                            If Left(YOSAN_FROM, 1) = "9" Then
                              Skip_Flg = True
                            End If
                        End If
                        
                        
                        
                        '予算元＝PP袋井のみ  2009.11.10
                        If Left(YOSAN_FROM, 2) = "PP" Then
                            
                            
                            If Not Not_SHUSI Then
                            
                                Skip_Flg = False
                            End If
                        End If
                        
                        If Left(YOSAN_FROM, 2) = "S1" And _
                            Left(YOSAN_TO, 2) = "S2" Then
                            WORK_SOKO = "88"
                        End If
            
                        '2006.05.01
                        If Trim(DEN_NO) = "POS88" Then
                            WORK_SOKO = "88"
                        End If
            
                End Select
             Case AIRCON                     'エアコン
                '除外倉庫に「CA」を追加 2006.07.27
                If Trim(HOST_SOKO) = "J4" Or _
                   Trim(HOST_SOKO) = "JG" Or _
                    Trim(HOST_SOKO) = "JW" Or _
                    Trim(HOST_SOKO) = "JV" Or _
                    Trim(HOST_SOKO) = "HY" Or _
                    Trim(HOST_SOKO) = "CA" Then
                    Skip_Flg = True
                End If
        
        
                If Left(YOSAN_FROM, 2) = "SH" Then
                    Skip_Flg = True
                End If
        
        
                If Left(YOSAN_FROM, 2) = "S1" Then
                    If Trim(HOST_SOKO) = "OS" Then
                      Skip_Flg = True
                    End If
                End If
        
                If Not Skip_Flg Then
                    'S2を追加 2009.11.04
                    'SSを追加 2010.03.08

'                    If Trim(HOST_SOKO) = "S8" Then
                    If Trim(HOST_SOKO) = "S8" Or Trim(HOST_SOKO) = "S2" Or Trim(HOST_SOKO) = "SS" Then
                        
                        
                        WORK_SOKO = "80"
                    Else
                        If CYU_KBN = "A" Then
                        Else
                            If CYU_KBN = "D" Then
                                WORK_SOKO = "70"
                            Else
                            End If
                        End If
                    End If
                End If
           
        
        
            Case OVEN           '電子レンジ 2012.09.28
                
        
    '6　1    ※     SDC    ※     90
    '6　1    001    SDC    ※     70←追加 仮入庫・ユニット
    '6  1    0102   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0201   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0601   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0602   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0701   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0702   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0801   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0802   SDC    ※     - ←追加 収支振替分は除外
    '6  1    0899   SDC    ※     - ←追加 収支振替分は除外
    '6  1    9101   SDC    ※     - ←追加 収支振替分は除外
    '6  1    9102   SDC    ※     - ←追加 収支振替分は除外
    '6  1    9301   SDC    ※     - ←追加 収支振替分は除外
    '6  1    9601   SDC    ※     - ←追加 収支振替分は除外
    '6  1    9602   SDC    ※     - ←追加 収支振替分は除外
    '6  1    9901   SDC    ※     - ←追加 収支振替分は除外
    '6  1    9902   SDC    ※     - ←追加 収支振替分は除外


        
                WORK_SOKO = "90"
                Select Case Trim(YOSAN_FROM)
                
                    Case "001"
                        WORK_SOKO = "70"
                

                        MAEGARI_FLG = True      '2018.11.16

                
                
                
                
                    Case "WP555"                    '2017.05.16
                        WORK_SOKO = "WP"            '2017.05.16
                    
                    
                        MAEGARI_FLG = True      '2018.11.16
                    
                    
                    
                    Case "0102"
                        Skip_Flg = True
                    Case "0201"
                        Skip_Flg = True
                    Case "0601"
                        Skip_Flg = True
                    Case "0602"
                        Skip_Flg = True
                    Case "0701"
                        If Trim(HOST_SOKO) = "01" Then          '2019.01.11
                        Else                                    '2019.01.11
                            Skip_Flg = True
                        End If                                  '2019.01.11
                    Case "0702"
                        Skip_Flg = True
                    Case "0801"
                        Skip_Flg = True
                    Case "0802"
                        Skip_Flg = True
                    Case "0899"
                        Skip_Flg = True
                    Case "9101"
                        Skip_Flg = True
                    Case "9102"
                        Skip_Flg = True
                    Case "9301"
                        Skip_Flg = True
                    Case "9601"
                        Skip_Flg = True
                    Case "9602"
                        Skip_Flg = True
                    Case "9901"
                        Skip_Flg = True
                    Case "9902"
                        Skip_Flg = True
                
                
                
                    Case "ZA071"                                    '2018.12.07
                        
                        
                        If (Trim(HOST_SOKO) = "01" Or Trim(HOST_SOKO) = "02" Or Trim(HOST_SOKO) = "99") Then    '2018.12.11
                        Else                                                                                    '2018.12.11
                            If Trim(HOST_SOKO) <> "06" Then     '2018.12.07
                                Skip_Flg = True                 '2018.12.07
                            End If                              '2018.12.07
                        End If                                                                                  '2018.12.11
                
                End Select
        
        
                If Trim(HOST_SOKO) = "06" And Trim(YOSAN_FROM) = "ZA071" Then     '2018.12.07
                Else                                                        '2018.12.70
        
        
                    If Trim(YOSAN_TO) <> "SDC" Then
                        Skip_Flg = True
                    End If
        
                End If                                                      '2018.12.7
        
        
                If Trim(HOST_SOKO) = "06" Then                  '2018.12.12
                    If (Trim(YOSAN_FROM) <> "ZA071" And Trim(YOSAN_FROM) <> "WP555") Then           '2018.12.12,2019.02.07
                        Skip_Flg = True                         '2018.12.12
                    End If                                      '2018.12.12
                End If                                          '2018.12.12
        
        
                '>> 2019.03.06
                If Trim(YOSAN_FROM) = "WP555" Then
                    If Trim(HOST_SOKO) = "01" Or Trim(HOST_SOKO) = "02" Or Trim(HOST_SOKO) = "06" Or Trim(HOST_SOKO) = "93" Or Trim(HOST_SOKO) = "99" Then
                        If Trim(YOSAN_TO) = "SDC" Then
                        
Debug.Print
                        
                        Else
                            Skip_Flg = True
                        End If
                    Else
                        Skip_Flg = True
                    End If
                End If
                '>> 2019.03.06
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>2015.10.16     食洗追加
            Case SHOKUSEN
                                                                                                                '"903" 追加 2015.10.21
'                If Trim(YOSAN_FROM) <> "S1S4" And Trim(YOSAN_FROM) <> "S1P4" And Trim(YOSAN_FROM) <> "904" And Trim(YOSAN_FROM) <> "903" Then
                                                                                                                '"PPP4" 追加 2017.02.17 "906"　追加　2019.04.15
'                If Trim(YOSAN_FROM) <> "S1S4" And Trim(YOSAN_FROM) <> "S1P4" And Trim(YOSAN_FROM) <> "904" And Trim(YOSAN_FROM) <> "903" And Trim(YOSAN_FROM) <> "PPP4" Then
                If Trim(YOSAN_FROM) <> "S1S4" And Trim(YOSAN_FROM) <> "S1P4" And Trim(YOSAN_FROM) <> "904" And Trim(YOSAN_FROM) <> "903" And Trim(YOSAN_FROM) <> "PPP4" _
                    And Trim(YOSAN_FROM) <> "906" Then
                    Skip_Flg = True
                End If
                
                If Trim(YOSAN_FROM) = "904" Then
'                   If Trim(HOST_SOKO) <> "S4" Then                                '2017.08.04
                    If Trim(HOST_SOKO) <> "S4" And Trim(HOST_SOKO) <> "P4" Then     '2017.08.04
                        Skip_Flg = True
                    End If
                End If
                
                If Trim(YOSAN_FROM) = "903" Then            '2015.10.21
                    If Trim(HOST_SOKO) <> "P4" Then         '2015.10.21
                        Skip_Flg = True                     '2015.10.21
                    End If                                  '2015.10.21
                End If                                      '2015.10.21
                
                
                If Trim(YOSAN_FROM) = "906" Then            '2019.04.15
                    If Trim(HOST_SOKO) <> "P4" Then         '2019.04.15
                        Skip_Flg = True                     '2019.04.15
                    End If                                  '2019.04.15
                End If                                      '2019.04.15
                
                
                
                
                
                If Trim(YOSAN_FROM) = "PPP4" And Trim(HOST_SOKO) = "P4" Then            '2017.02.17
                    WORK_SOKO = "81"                                                    '2017.02.17
                End If                                                                  '2017.02.17
                
                
                If Trim(YOSAN_TO) <> "SDC" Then
                    Skip_Flg = True
                End If
            '>>>>>>>>>>>>>>>>>>>>>>>>>2015.10.16     食洗追加
                
        
        End Select
            
            
            
            
        
    
    
        If Not Skip_Flg Then
                                        
                                        
            
                
                                        '入荷予定重複チェック
            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
    
            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
                    Skip_Flg = True
                Case BtErrKeyNotFound
                Case Else
                    'Call File_Error(sts, BtOpGetEqual, "入荷予定", 0)              '2016.06.23
                    Call File_Error(sts, BtOpGetEqual, "入荷予定", 1, Y_NYU_ID)     '2016.06.23
'                    Exit Function      '2015.11.19
                    GoTo Abort_Tran     '2015.11.19
            End Select
        
        
        
        
        
            If Not Skip_Flg Then
                                                'トランザクション開始
                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
'                    Exit Function      '2015.11.19
                    GoTo Abort_Tran     '2015.11.19
                End If
                                            '品目マスタチェック
'                If Item_Check_Proc(In_Mode, JGYOBU, Soko_T(i, j).NAIGAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
'                    GoTo Abort_Tran
'                End If
                                            
                If Item_Check_Proc(In_Mode, JGYOBU, Soko_T(i, j).NAIGAI, HIN_GAI, HIN_NAI, HIN_NAME, , , , KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO) Then
                    GoTo Abort_Tran
                End If
                                            
                                            
                '2012.12.20
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "0" And StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "1" Then
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_F)
                End If
                '2012.12.20
                                            
                                            '入荷データ作成
                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
        
        
                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANKA, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
        
        
                Last_Proc_F = True              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り
        
        
                '入荷ﾁｪｯｸﾃﾞｰﾀ更新
                
                
                If MAEGARI_FLG Then                                         '2018.11.16
                
                    
                    If WORK_SOKO = "70" And JGYOBU = OVEN Then
                    
                        'Call LOG_OUT(LOG_F, "HIN_GAI=" & HIN_GAI & " YOTEI_QTY=" & YOTEI_QTY)
                        If MAEGARI_PROC(JGYOBU_Code, HIN_GAI, YOTEI_QTY) Then   '2018.11.16
                            Unload Me                                           '2018.11.16
                        End If                                                  '2018.11.16
                
                    End If
                
                
                    WK_E_QTY = 0                                            '2018.11.16
                    
                Else                                                        '2018.11.16
                
                
                
                
                
                    If JGYOBU = OVEN And Trim(YOSAN_FROM) <> "4HHK" Then        '2019.02.08
                        WK_E_QTY = 0                                            '2019.02.08
                    Else                                                        '2019.02.08
                
                
                        Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
                        Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
                        Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_GAI)
            
                        WK_Y_QTY = CLng(YOTEI_QTY)
            
            
                        Loop_Cnt = 0
            
                        Do
                            sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
                                        WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
                                        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
                                
                                        Loop_Cnt = 0
                                        Do
                                        
                                            sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
        '                                            Beep
        '                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
        '                                            If ans = vbCancel Then
        ''                                                Exit Function
        '                                                GoTo Abort_Tran
        '                                            End If
                                                
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                                
                                                Case Else
                                                    'Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)            '2016.06.23
                                                    Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ", 1, J_NYU_ID)   '2016.06.23
        '                                            Exit Function
                                                    GoTo Abort_Tran
                                            End Select
                                        
                                        Loop
                                        WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                    Else
                                        
                                        
                                        Loop_Cnt = 0
                                                                        
                                        Do
                                            sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
        '                                            Beep
        '                                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
        '                                            If ans = vbCancel Then
        ''                                                Exit Function
        '                                                GoTo Abort_Tran
        '                                            End If
                                                
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                                
                                                
                                                Case Else
                                                    'Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)            '2016.06.23
                                                    Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ", 1, J_NYU_ID)   '2016.06.23
        '                                            Exit Function
                                                    GoTo Abort_Tran
                                            End Select
                                        Loop
                                        WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                                    End If
                            
                                    Exit Do
                                Case BtErrKeyNotFound
                                    WK_E_QTY = 0
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
        '                            Beep
        '                            ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
        '                            If ans = vbCancel Then
        ''                                Exit Function
        '                                GoTo Abort_Tran
        '                           End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                
                                
                                Case Else
                                    'Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)              '2016.06.23
                                    Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ", 1, J_NYU_ID)     '2016.06.23
        '                            Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
                    End If                                          '2019.02.08
                End If
                                    
                                    
                                    '先行入荷数（入荷実績数）
                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
        
                                    '予算単位元
                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                                    '予算単位先
                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                                    '標準棚番
                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
                                    'H倉庫 2006.10.17
                Call UniCode_Conv(Y_NYUREC.H_SOKO, HOST_SOKO)

                                    '入荷リスト出力フラグ   2007.06.12
                Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, " ")








                '----------------   2010.07.08 ▽
                
                
                If Trim(GENSANKOKU) = "" And Trim(GEN_GENSANKOKU) = "" And Trim(SHIIRE_WORK_CENTER) = "" Then
                
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))                    '原産国名
                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, StrConv(ITEMREC.TORI_GEN_GENSANKOKU, vbUnicode))            '現物表示原産国名
                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))    '資材仕入先ﾜｰｸｾﾝﾀｰ
                
                Else
                
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)                    '原産国名
                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, GEN_GENSANKOKU)            '現物表示原産国名
                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)    '資材仕入先ﾜｰｸｾﾝﾀｰ
                End If
                
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, KANKYO_KBN)                          '環境種類区分
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, KANKYO_KBN_ST)                    '環境種類区分適用開始
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)              '環境種類区分数量
                Call UniCode_Conv(Y_NYUREC.ID_NO2, ID_NO2)                                  'ID_NO
                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, AITESAKI_CODE)                    '相手先ｺｰﾄﾞ
                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, JYUCHU_YMD)                          '受注年月日
                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)              '指定納期年月日
                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "0")                             '入庫関連ﾘｽﾄ出力F
                    
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "8")                           '入庫管理ﾘｽﾄ出力F
                If StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode) <> "" And Mid(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode), 1, 1) > " " Then
                    
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                        
                    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                            Call UniCode_Conv(K0_GENSAN.NAIGAI, Soko_T(i, j).NAIGAI)
                            Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                            Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")
                            sts = BTRV(BtOpGetGreaterEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                        
                            Select Case sts
                                Case BtNoErr
                                
                                    If Trim(HIN_GAI) = Trim(StrConv(GENSANREC.HIN_GAI, vbUnicode)) Then
                                        Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "0")                   '入庫管理ﾘｽﾄ出力F
                                    End If
                                
                                
                                
                                Case BtErrEOF
                        
                                Case Else
                                  
                                    'Call File_Error(sts, BtOpGetGreaterEqual, "原産国ﾏｽﾀ", 0)              '2016.06.23
                                    Call File_Error(sts, BtOpGetGreaterEqual, "原産国ﾏｽﾀ", 1, GENSAN_ID)    '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                    End Select
                End If

                Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "0")                       '入庫ﾁｪｯｸﾘｽﾄ出力F
                
                
                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, (WORK_SOKO & _
                                                            "01" & "01" & "01"))        '入庫棚番
                                                                                        '前借相殺数
                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, Format(WK_E_QTY, "00000000"))
                
                
                
                
                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
                
                
                
                '2011.03.23 発生元プログラム
                Call UniCode_Conv(Y_NYUREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
                '2011.03.23 元テキスト№
                If Trim(MOTO_TEXT_NO) = "" Then
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, "")
                Else
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
                End If
                
                '----------------   2010.07.08 △








                Call UniCode_Conv(Y_NYUREC.FILLER, "")
                
                
                Loop_Cnt = 0
                            
                Do
                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                            Beep
'                            ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                            If ans = vbCancel Then
''                                Exit Function
'                                GoTo Abort_Tran
'                            End If
                        
                            Loop_Cnt = Loop_Cnt + 1
                            If Loop_Cnt > 5 Then
                                GoTo Abort_Tran
                            End If
                        
                            DoEvents
                            Sleep (500)
                        
                        Case Else
                            'Call File_Error(sts, BtOpInsert, "入荷予定", 0)            '2016.06.23
                            Call File_Error(sts, BtOpInsert, "入荷予定", 1, Y_NYU_ID)   '2016.06.23
'                            Exit Function
                            GoTo Abort_Tran
                    End Select
                Loop
            
            
                '----------------   2010.07.08 ▽
                '原産国のチェック＆登録
                If StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode) <> "" And Mid(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode), 1, 1) > " " Then
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, NAIGAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_GAI)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            
                            '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.12.28
                            Call UniCode_Conv(GENSANREC.UPD_TANTO, "2010")
                            Call UniCode_Conv(GENSANREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                            
                            Loop_Cnt = 0
                            
                            Do
                                sts = BTRV(BtOpUpdate, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    
                                        Loop_Cnt = Loop_Cnt + 1
                                        If Loop_Cnt > 5 Then
                                            GoTo Abort_Tran
                                        End If
                                    
                                        DoEvents
                                        Sleep (500)
                                    
                                    
                                    Case Else
                                        'Call File_Error(sts, BtOpInsert, "原産国ﾏｽﾀ", 0)               '2016.06.23
                                        Call File_Error(sts, BtOpInsert, "原産国ﾏｽﾀ", 1, GENSAN_ID)     '2016.06.23
                                        GoTo Abort_Tran
                                End Select
                            Loop
                            '>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.12.28
                        
                        
                        
                        Case BtErrKeyNotFound
                        
                            Call UniCode_Conv(GENSANREC.JGYOBU, JGYOBU)
                            Call UniCode_Conv(GENSANREC.NAIGAI, NAIGAI)
                            Call UniCode_Conv(GENSANREC.HIN_GAI, HIN_GAI)
                            Call UniCode_Conv(GENSANREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                            Call UniCode_Conv(GENSANREC.FILLER, "")
                            Call UniCode_Conv(GENSANREC.INS_TANTO, "2010")
                            Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
                            Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
                            Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
                            
                            Loop_Cnt = 0
                            
                            Do
                                sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                        If ans = vbCancel Then
''                                            Exit Function
'                                            GoTo Abort_Tran
'                                        End If
                                    
                                        Loop_Cnt = Loop_Cnt + 1
                                        If Loop_Cnt > 5 Then
                                            GoTo Abort_Tran
                                        End If
                                    
                                        DoEvents
                                        Sleep (500)
                                    
                                    
                                    Case Else
                                        'Call File_Error(sts, BtOpInsert, "原産国ﾏｽﾀ", 0)               '2016.06.23
                                        Call File_Error(sts, BtOpInsert, "原産国ﾏｽﾀ", 1, GENSAN_ID)     '2016.06.23
'                                        Exit Function
                                        GoTo Abort_Tran
                                End Select
                            Loop
                        Case Else
                            'Call File_Error(sts, BtOpGetEqual, "原産国ﾏｽﾀ", 0)                 '2016.06.23
                            Call File_Error(sts, BtOpGetEqual, "原産国ﾏｽﾀ", 1, GENSAN_ID)       '2016.06.23
'                            Exit Function
                            GoTo Abort_Tran
                    End Select
                End If
                '----------------   2010.07.08 △
            
            
            
            
            
            
            
'------------ 2005.12.30
                Select Case JGYOBU
                    Case AIRCON, SENTAKU
                        Call UniCode_Conv(K0_SOKO.Soko_No, WORK_SOKO)
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case Else
                                'Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)            '2016.06.23
                                Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 1, SOKO_ID)    '2016.06.23
'                                Exit Function
                                GoTo Abort_Tran
                        End Select
        
                        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
        
                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            MI_QTY = 0
                        Else
                        
                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                SUMI_QTY = 0
                            Else
                                SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                MI_QTY = 0
                            End If
                        End If
                        
'------------ 2005.12.30
                        
                    Case Else
                        
                        
                        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                            MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            SUMI_QTY = 0
                        Else
                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            MI_QTY = 0
                        End If
                End Select
                
        
'                Wk_SOKO = KASO_NYUKA_SOKO
'                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
'                    Wk_SOKO = KASO_SMODOSHI_SOKO
'
'                End If
        
        
        
        
        
                '入荷数で在庫データ更新（＋）
                If Nyuko_Update_Proc(JGYOBU, _
                                    Soko_T(i, j).NAIGAI, _
                                    HIN_GAI, _
                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                    (WORK_SOKO & "01" & "01" & "01"), _
                                    YOIN_TU_NYUKA, _
                                    SUMI_QTY, MI_QTY, _
                                    WS_NO, WS_NO, 5, _
                                    DEN_DT & " 伝№:" & DEN_NO, , , , MENU_NO, , , StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode), StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode), ID_NO2, YOSAN_FROM, YOSAN_TO) Then
'                    Exit Function
                    GoTo Abort_Tran
            
                End If
            
                '前借り数で在庫データ更新（－）
                If WK_E_QTY <> 0 Then
                '在庫データLOCK
                    If Zaiko_Lock_Proc((WORK_SOKO & "01" & "01" & "01"), _
                                        JGYOBU, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        WS_NO, , , 5) Then
'                        Exit Function
                        GoTo Abort_Tran
    
                    End If
        
                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                        MI_QTY = WK_E_QTY
                    Else
                        SUMI_QTY = WK_E_QTY
                    End If
            
            
                    If Syuko_Update_Proc(JGYOBU, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        DEN_DT, _
                                        (WORK_SOKO & "01" & "01" & "01"), _
                                        YOIN_MAE_SOUSAI, _
                                        SUMI_QTY, MI_QTY, 0, _
                                        WS_NO, WS_NO, 5) Then
'                        Exit Function
                        GoTo Abort_Tran
        
                    End If
            
            
            
            
            
            
                End If
                
                
                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    GoTo Abort_Tran
                End If
                
                
                Out_Cnt = Out_Cnt + 1
                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                DoEvents
    
            End If
        
        
        
        
        
'出荷予定変換################################################## 2005/05/16 Add 滋賀物流↓
        Else
            
            If JGYOBU = AIRCON Then
            
                If Not_SHUSI And Trim(HOST_SOKO) <> "R8" Then     '2019/12/13 滋賀DC 収支R8対応
                Else
                    If IO_KBN = "2" Then
                                
                        
                        wkMUKE_CODE = ""
                        
                        If Trim(HOST_SOKO) = "S8" Then
                            wkMUKE_CODE = "S8"
                        ElseIf Trim(HOST_SOKO) = "R8" Then '2019/12/13 滋賀DC 収支R8対応
                            wkMUKE_CODE = "R8"             '2019/12/13 滋賀DC 収支R8対応
                        Else
                            If Trim(HOST_SOKO) = "ST" Then              'ST追加　   2016.03.11
                                wkMUKE_CODE = "ST"                      '           2016.03.11
                            Else                                        '           2016.03.11
                                If Trim(HOST_SOKO) = "SH" Then
                                Else
                                    Select Case Trim(YOSAN_TO)
                                    
                                        Case "Z0014"
                                            wkMUKE_CODE = "LM"
                                        Case "B0070"
                                            If Trim(HOST_SOKO) = "S2" Then
                                                wkMUKE_CODE = "S2"
                                            Else
                                                wkMUKE_CODE = "AC"
                                            End If
                                        Case Else
                                             wkMUKE_CODE = "AC"
                                    End Select
                                End If                                  '           2016.03.11
                            End If
                        
                            If Trim(HOST_SOKO) <> "S2" Then             '≠"S2" and ="B0070"    '2019.01.22
                                If Trim(YOSAN_TO) = "B0070" Then
                                    wkMUKE_CODE = "AC"
                                End If
                            End If
                        End If
            
            
                        If wkMUKE_CODE = "" Then
                        Else
'                            Skip_Flg = False
                                                        '入荷予定重複チェック
'                            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'                            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
'                            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
'
'                            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
'                                    Skip_Flg = True
'                                Case BtErrKeyNotFound
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "入荷予定")
'                                    Exit Function
'                            End Select
            
            
            
            
'
'                            Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
'                            Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
'                            Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
'
'                            sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
'                                    Skip_Flg = True
'                                Case BtErrKeyNotFound
'                                Case Else
'                                    Call File_Error(sts, BtOpGetEqual, "照合用入荷予定")
'                                    Exit Function
'                            End Select
            
            
                            
                            If Not DUP_FLG Then
            
                                                                'トランザクション開始
                                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts <> BtNoErr Then
                                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                    Exit Function
                                End If
                                                                '品目マスタチェック
                                If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                                    GoTo Abort_Tran
                                End If
            
        '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
                                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
'                                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI) '2019/12/13 滋賀DC 収支R8対応
                                Call UniCode_Conv(Y_NYUREC.NAIGAI, "1")
                                Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                        
                        
                                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                                Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                                Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                                Call UniCode_Conv(Y_NYUREC.TANKA, "")
                                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                                Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                                Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
                
                
                                Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
                
                
                
                                                    '入荷リスト出力フラグ   2007.06.12
                                Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
                
                
                
                
                
                
                
                
                
                                '----------------   2010.07.08 ▽
                                Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)                      '原産国名
                                Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, GEN_GENSANKOKU)              '現物表示原産国名
                                Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)      '資材仕入先ﾜｰｸｾﾝﾀｰ
                                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, KANKYO_KBN)                      '環境種類区分
                                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, KANKYO_KBN_ST)                '環境種類区分適用開始
                                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)          '環境種類区分数量
                                Call UniCode_Conv(Y_NYUREC.ID_NO2, ID_NO2)                              'ID_NO
                                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, AITESAKI_CODE)                '相手先ｺｰﾄﾞ
                                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, JYUCHU_YMD)                      '受注年月日
                                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)          '指定納期年月日
                                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "")                          '入庫ﾘｽﾄ出力F
                                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, "")                           '入庫棚番
                                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, "")                           '前借相殺数
                                
                                
                                
                                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
                                
                                
                                
                                
                                
                                '----------------   2010.07.08 △
                
                
                
                
                
                
                
                
                
                
                
                
                                Call UniCode_Conv(Y_NYUREC.FILLER, "")
                                
'                                Do
'                                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                                    Select Case sts
'                                        Case BtNoErr
'                                            Exit Do
'                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
'                                        Case Else
'                                            Call File_Error(sts, BtOpInsert, "入荷予定")
'                                            Exit Function
'                                    End Select
'                                Loop
        '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
            
                                Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                                Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                                Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                                Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
            
                                If Rec_LENG = 138 Then                                      '2016.04.19
                                    If Den_No_Set_Proc(21, JGYOBU, wkStr) Then
                                        GoTo Abort_Tran
                                    Else
                                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, wkStr)
                                    End If
                                Else                                                        '2016.04.19
                                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)           '2016.04.19
                                End If                                                      '2016.04.19
            
            
                                Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                                Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, wkMUKE_CODE)
                                Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                                Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                                Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                                Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                                Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                                Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                                
                                Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                                
                                'If JGYOBU = AIRCON Then             '2008.02.01
                                '    If Left(DEN_NO, 1) = "0" Then
                                '        DEN_NO = Right(DEN_NO, Len(DEN_NO) - 1)
                                '    End If
                                'End If
                                
                                Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                                
                                
                                
                                
                                
                                wkStr = Format(Val(YOTEI_QTY), "0000000")
                                Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                                Call UniCode_Conv(Y_SYUREC.MUKE_CODE, wkMUKE_CODE)
                                Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                                Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                                
                                Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                                Call UniCode_Conv(Y_SYUREC.TANKA, "")
                                
                                
                                Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                                Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                                Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                                Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                                
                                
                                
                                Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)
    
                                Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                                Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                                Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
    
                                Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)
    
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                                Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                                Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                                Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                                Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)
    
                                Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
                                Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
                                Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
                                Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
                                Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
                                Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
                                Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
                                Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
                                Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
                                Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
                                Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
                                Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
                                Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                                Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
    
                                Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                                Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                                Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                                Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                                Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                                Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                                Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                                Call UniCode_Conv(Y_SYUREC.FILLER, "")
            
                                Loop_Cnt = 0
                                
                                Do
                                    sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                            If ans = vbCancel Then
''                                                Exit Function
'                                                GoTo Abort_Tran
'                                            End If
                                        
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                GoTo Abort_Tran
                                            End If
                                        
                                            DoEvents
                                            Sleep (500)
                                        
                                        
                                        Case Else
                                            'Call File_Error(sts, BtOpInsert, "出荷予定", 0)            '2016.06.23
                                            Call File_Error(sts, BtOpInsert, "出荷予定", 1, Y_SYU_ID)   '2016.06.23
'                                            Exit Function
                                            GoTo Abort_Tran
                                    End Select
                                Loop
            
                                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                If sts <> BtNoErr Then
                                    GoTo Abort_Tran
                                End If
            
                                Out_Cnt = Out_Cnt + 1
                                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                                DoEvents
            
                                If SYUKA_LOG_ON Then
                                    Call SYUKA_LOG_OUT_PROC("INS", "入荷から生成")
                                End If
            
'                                If Not Fast_Flg Then
'                                    Close #DUP_SYUKANo
'                                End If
                            End If
                        End If
                    End If
                End If
            End If
'#################################################################################### 2005/05/16 Add ↑
        
'#################################################################################### 2008/02/22 Add ↓
            If JGYOBU = SOJIKI Then
            
                If IO_KBN = "1" Then
                            
                    wkMUKE_CODE = ""
                            
                    If Trim(HOST_SOKO) = "SS" Then
                        wkMUKE_CODE = "00000000"
                    End If
                    If Trim(HOST_SOKO) = "ZZ" Then
                        wkMUKE_CODE = "88888888"
                    End If
        
        
                    If wkMUKE_CODE = "" Then
                    Else
 '                       Skip_Flg = False
                                                    '入荷予定重複チェック
 '                       Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
 '                       Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
 '                       Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
 '
 '                       sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
 '                       Select Case sts
 '                           Case BtNoErr
 '                               Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
 '                               Skip_Flg = True
 '                           Case BtErrKeyNotFound
 '                           Case Else
 '                               Call File_Error(sts, BtOpGetEqual, "入荷予定")
 '                               Exit Function
 '                       End Select
        
        
 '                       Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
 '                       Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
 '                       Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
 '
 '                       sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
 '                       Select Case sts
 '                           Case BtNoErr
 '                               Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
 '                               Skip_Flg = True
 '                           Case BtErrKeyNotFound
 '                           Case Else
 '                               Call File_Error(sts, BtOpGetEqual, "照合用入荷予定")
 '                               Exit Function
 '                       End Select
        
        
        
                        If Not DUP_FLG Then
        
                                                            'トランザクション開始
                            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            If sts <> BtNoErr Then
                                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                Exit Function
                            End If
                                                            '品目マスタチェック
                            If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                                GoTo Abort_Tran
                            End If
        
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
                            Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                            Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                            Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                            Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                    
                    
                            Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                            Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                            Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                            Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                            Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                            Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                            Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                            Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                            Call UniCode_Conv(Y_NYUREC.TANKA, "")
                            Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                            Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                            Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                            Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                            Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                            Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                            Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                            Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                            Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                            Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                            Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                            Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                            Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                            Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                            Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                            Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                            Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                            Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                            Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                            Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                            Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                            Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                            Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                            Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                            Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                            Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                            Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                            Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                            Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                            Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
            
            
                            Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
            
            
            
                                                '入荷リスト出力フラグ   2007.06.12
                            Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
            
                            Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                            Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                            Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                            Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
            
                            Call UniCode_Conv(Y_NYUREC.FILLER, "")
                            
'                            Do
'                                sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                                Select Case sts
'                                    Case BtNoErr
'                                       Exit Do
'　                                  Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                        If ans = vbCancel Then
'                                            Exit Function
'                                        End If
'                                    Case Else
'                                        Call File_Error(sts, BtOpInsert, "入荷予定")
'                                        Exit Function
'                                End Select
'                            Loop
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
        
                            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                            Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                            Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                            
                            If Trim(HOST_SOKO) = "SS" Then
                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                            Else
                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                            End If
        
                            
                            If Rec_LENG = 138 Then                                  '2016.04.19
                                If Den_No_Set_Proc(21, JGYOBU, wkStr) Then
                                    GoTo Abort_Tran
                                Else
                                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, wkStr)
                                End If
                            Else                                                    '2016.04.19
                                Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)       '2016.04.19
                            End If                                                  '2016.04.19
        
                            Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                            Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, wkMUKE_CODE)
                            Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                            Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                            Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                            Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                            Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                            Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                            
                            Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                            'If JGYOBU = AIRCON Then             '2008.02.01
                            '    If Left(DEN_NO, 1) = "0" Then
                            '        DEN_NO = Right(DEN_NO, Len(DEN_NO) - 1)
                            '    End If
                            'End If
                            
                            Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                            
                            
                            
                            
                            
                            wkStr = Format(Val(YOTEI_QTY), "0000000")
                            Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                            Call UniCode_Conv(Y_SYUREC.MUKE_CODE, wkMUKE_CODE)
                            Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                            Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                            
                            Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                            Call UniCode_Conv(Y_SYUREC.TANKA, "")
                            
                            
                            Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                            Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                            Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                            Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                            
                            
                            
                            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                            Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                            Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                            Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                            Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                            If Trim(HOST_SOKO) = "SS" Then
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                            Else
                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                            End If

                            
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                            Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                            Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                            Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                            Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

                            Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
                            Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
                            Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
                            Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
                            Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
                            Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
                            Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
                            Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
                            Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
                            Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
                            Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
                            Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
                            Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                            Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")

                            Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                            Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                            Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                            Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                            Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                            Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                            Call UniCode_Conv(Y_SYUREC.FILLER, "")
        
        
                            Loop_Cnt = 0
        
                            Do
                                sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                        If ans = vbCancel Then
''                                            Exit Function
'                                            GoTo Abort_Tran
'                                        End If
                                    
                                    
                                        Loop_Cnt = Loop_Cnt + 1
                                        If Loop_Cnt > 5 Then
                                            GoTo Abort_Tran
                                        End If
                                    
                                        DoEvents
                                        Sleep (500)
                                    
                                    
                                    
                                    Case Else
                                        'Call File_Error(sts, BtOpInsert, "出荷予定", 0)            '2016.06.23
                                        Call File_Error(sts, BtOpInsert, "出荷予定", 1, Y_SYU_ID)   '2016.06.23
'                                        Exit Function
                                        GoTo Abort_Tran
                                End Select
                            Loop
        
                            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            If sts <> BtNoErr Then
                                GoTo Abort_Tran
                            End If
        
                            Out_Cnt = Out_Cnt + 1
                            lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                            DoEvents
        
                            If SYUKA_LOG_ON Then
                                Call SYUKA_LOG_OUT_PROC("INS", "入荷から生成")
                            End If
        
'                            If Not Fast_Flg Then
'                                Close #DUP_SYUKANo
'                            End If
                        End If
                    End If
                End If
            End If
        
'#################################################################################### 2008/02/22 Add ↑
        
        
        
        
        
        
        
        
        
        
        
'#################################################################################### 2010/07/21 Add ↓
            If JGYOBU = DENKA Then
            
                If IO_KBN = "2" And Trim(YOSAN_FROM) = "SDC" And Trim(AITESAKI_CODE) = "01KA" Then
                            
'                    Skip_Flg = False
                                                '入荷予定重複チェック
'                    Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'                    Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
'                    Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
'
'                    sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
'                            Skip_Flg = True
'                        Case BtErrKeyNotFound
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "入荷予定")
'                            Exit Function
'                    End Select
    
'                    Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU)
'                    Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
'                    Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)
'
'                    sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                            Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP 事業部=" & JGYOBU & "ＴＥＸＴＩＤ＝" & TEXT_NO)
'                            Skip_Flg = True
'                        Case BtErrKeyNotFound
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "照合用入荷予定")
'                            Exit Function
'                    End Select
    
    
                    If Not DUP_FLG Then
    
                                                        'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        '                    Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                        End If
                                                        '品目マスタチェック
                        If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                            GoTo Abort_Tran
                        End If
        
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                        Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                
                
                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
        
        
                        Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
        
        
        
                                            '入荷リスト出力フラグ   2007.06.12
                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
        
        
                        Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                        Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                        Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                        Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
        
        
                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                        
'                        Do
'                            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'                            Select Case sts
'                                Case BtNoErr
'                                    Exit Do
'                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                    If ans = vbCancel Then
'                                        Exit Function
'                                    End If
'                                Case Else
'                                    Call File_Error(sts, BtOpInsert, "入荷予定")
'                                    Exit Function
'                            End Select
'                        Loop
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
        
                        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                        Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                        
                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                        End If
    
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)
    
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                        
                        Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                        Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                        
                        wkStr = Format(Val(YOTEI_QTY), "0000000")
                        Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.TANKA, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                        
                        
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                        End If

                        
                        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

                        Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")

                        Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    
                        Loop_Cnt = 0
    
                        Do
                            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                    If ans = vbCancel Then
''                                        Exit Function
'                                        GoTo Abort_Tran
'                                    End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                Case Else
                                    'Call File_Error(sts, BtOpInsert, "出荷予定", 0)            '2016.06.23
                                    Call File_Error(sts, BtOpInsert, "出荷予定", 1, Y_SYU_ID)   '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
    
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            GoTo Abort_Tran
                        End If
    
                        Out_Cnt = Out_Cnt + 1
                        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                        DoEvents
    
                        If SYUKA_LOG_ON Then
                            Call SYUKA_LOG_OUT_PROC("INS", "入荷から生成")
                        End If
    
'                        If Not Fast_Flg Then
'                            Close #DUP_SYUKANo
'                        End If
                    End If
                End If
            End If
        
'#################################################################################### 2008/02/22 Add ↑
        
        
'#################################################################################### 2018/09/19 Add ↓
            If JGYOBU = OVEN Then
                If (IO_KBN = "2" And Trim(HOST_SOKO) = "01" And Trim(YOSAN_FROM) = "SDC" And Trim(AITESAKI_CODE) = "0107") Or _
                    (IO_KBN = "2" And Trim(HOST_SOKO) = "06" And Trim(YOSAN_FROM) = "SDC" And Trim(AITESAKI_CODE) = "0607") Or _
                    (IO_KBN = "2" And Trim(YOSAN_FROM) = "SDC" And Mid(AITESAKI_CODE, 3, 2) = "KA") Then
                            
    
    
                    If Not DUP_FLG Then
    
                                                        'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        '                    Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                        End If
                                                        '品目マスタチェック
                        If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                            GoTo Abort_Tran
                        End If
        
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                        Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                
                
                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
        
        
                        Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
        
        
        
                                            '入荷リスト出力フラグ   2007.06.12
                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
        
        
                        Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                        Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                        Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                        Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
        
        
                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                        
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
        
                        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                        Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                        
                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                        End If
    
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)
    
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                        
                        Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                        Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                        
                        wkStr = Format(Val(YOTEI_QTY), "0000000")
                        Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.TANKA, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                        
                        
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                        End If

                        
                        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

                        Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")

                        Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    
                        Loop_Cnt = 0
    
                        Do
                            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                    If ans = vbCancel Then
''                                        Exit Function
'                                        GoTo Abort_Tran
'                                    End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                Case Else
                                    'Call File_Error(sts, BtOpInsert, "出荷予定", 0)            '2016.06.23
                                    Call File_Error(sts, BtOpInsert, "出荷予定", 1, Y_SYU_ID)   '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
    
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            GoTo Abort_Tran
                        End If
    
                        Out_Cnt = Out_Cnt + 1
                        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                        DoEvents
    
                        If SYUKA_LOG_ON Then
                            Call SYUKA_LOG_OUT_PROC("INS", "入荷から生成")
                        End If
    
                    End If
                End If
            End If
        
'#################################################################################### 2018/09/19 Add ↑
        
'#################################################################################### 2018/09/20 Add ↓
            If JGYOBU = SHOKUSEN Then
                If IO_KBN = "2" And Trim(YOSAN_TO) = "904" Then
                            
    
    
                    If Not DUP_FLG Then
    
                                                        'トランザクション開始
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        '                    Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                        End If
                                                        '品目マスタチェック
                        If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                            GoTo Abort_Tran
                        End If
        
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                        Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
                
                
                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
        
        
                        Call UniCode_Conv(Y_NYUREC.H_SOKO, " ")
        
        
        
                                            '入荷リスト出力フラグ   2007.06.12
                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, "9")
        
        
                        Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                        Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                        Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                        Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
        
        
                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                        
    '-------------------------------------------------------'入荷データのみ登録する（再取込み時のﾁｪｯｸのため）
        
                        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                        Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                        Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                        
                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_TUK)
                        End If
    
                        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO2)
    
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_GAI)
                        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
                        Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
                        
                        Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_GAI)
                            
                        Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                        
                        wkStr = Format(Val(YOTEI_QTY), "0000000")
                        Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
                        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, AITESAKI_CODE)
                        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, HOST_SOKO)
                        Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, DEN_DT)
                        Call UniCode_Conv(Y_SYUREC.TANKA, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                        Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                        
                        
                        
                        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, DEN_DT)

                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")

                        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, SYUK_NAME)

                        If Trim(HOST_SOKO) = "SS" Then
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                        Else
                            Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_TUK)
                        End If

                        
                        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                        Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                        Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CYOK_KBN)

                        Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
                        Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
                        Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
                        Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
                        Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
                        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
                        Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
                        Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
                        Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
                        Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
                        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                        Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")

                        Call UniCode_Conv(Y_SYUREC.HIN_NAI, HIN_NAI)
                        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
    
                        Loop_Cnt = 0
    
                        Do
                            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                    Beep
'                                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                    If ans = vbCancel Then
''                                        Exit Function
'                                        GoTo Abort_Tran
'                                    End If
                                
                                
                                    Loop_Cnt = Loop_Cnt + 1
                                    If Loop_Cnt > 5 Then
                                        GoTo Abort_Tran
                                    End If
                                
                                    DoEvents
                                    Sleep (500)
                                
                                Case Else
                                    'Call File_Error(sts, BtOpInsert, "出荷予定", 0)            '2016.06.23
                                    Call File_Error(sts, BtOpInsert, "出荷予定", 1, Y_SYU_ID)   '2016.06.23
'                                    Exit Function
                                    GoTo Abort_Tran
                            End Select
                        Loop
    
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            GoTo Abort_Tran
                        End If
    
                        Out_Cnt = Out_Cnt + 1
                        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                        DoEvents
    
                        If SYUKA_LOG_ON Then
                            Call SYUKA_LOG_OUT_PROC("INS", "入荷から生成")
                        End If
    
                    End If
                End If
            End If
        
'#################################################################################### 2018/09/19 Add ↑
        
        
        
        
        
        End If
        
        
        
    
    Loop

    Nyuka_Update_Proc = False
    Exit Function

Abort_Tran:
    
'>>>>>  2015.11.19
    If Fast_Flg Then
        Open (FileName) For Output As DUP_SYUKANo
        Write #DUP_SYUKANo, , , "入庫取込み異常リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
        Write #DUP_SYUKANo, "エラー内容", "伝票日付", "伝票№", "予算元", "予算先", "ﾎｽﾄ倉庫", "品番", "数量", "TEXT_NO"      '2015.11.19
        Fast_Flg = False
    End If


    Write #DUP_SYUKANo, "＜重複＞",
    Write #DUP_SYUKANo, DEN_DT,
    Write #DUP_SYUKANo, DEN_NO,
    Write #DUP_SYUKANo, YOSAN_FROM,
    Write #DUP_SYUKANo, YOSAN_TO,
    Write #DUP_SYUKANo, HOST_SOKO,

    Write #DUP_SYUKANo, HIN_GAI,
    Write #DUP_SYUKANo, YOTEI_QTY,
    Write #DUP_SYUKANo, TEXT_NO
'>>>>>  2015.11.19
    
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


    

End Function
    
Private Function Syuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   「出荷予定データ」更新処理
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim Skip_Flg    As Boolean
Dim sts         As Integer
    
Dim ans         As Integer

Dim c               As String * 128

Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
    
Dim INS_NOW         As String * 14

Dim wkSS            As String
Dim wkMUKE_CODE     As String
Dim wkCHOKU_KBN     As String * 1


Dim wkText          As String
Dim Length      As Integer


Dim JGYOBA              As String * 8       '事業場
Dim DATA_KBN            As String * 1       'データ区分
Dim TORI_KBN            As String * 2       '取引区分
Dim ID_NO               As String * 12      'ID-NO
Dim KAIKEI_JGYOBA       As String * 8       '会計用事業場ｺｰﾄﾞ
Dim SHISAN_JGYOBA       As String * 8       '資産管理事業場ｺｰﾄﾞ
Dim HIN_NO              As String * 20      '品目番号
Dim DEN_NO              As String * 10      '伝票番号
Dim SURYO               As String * 7       '出庫数量
Dim MUKE_CODE           As String * 8       '出庫先
Dim SYUKO_SYUSI         As String * 8       '出庫収支
Dim SHISAN_SYUSI        As String * 8       '資産管理用在庫収支ｺｰﾄﾞ
Dim HOJYO_SYUSI         As String * 8       '補助在庫収支ｺｰﾄﾞ
Dim SYUKO_YMD           As String * 8       '出庫日付
Dim TANKA               As String * 10      '単価
Dim ODER_NO             As String * 12      'オーダー番号
Dim ITEM_NO             As String * 5       'アイテム番号
Dim ODER_NO_R           As String * 5       'オーダー略号
Dim KOSO_KEITAI         As String * 14      '個装形態       10-->14 2011.10.31
Dim SYUKA_YMD           As String * 8       '出荷日
Dim TANABAN1            As String * 10      '棚番１
Dim TANABAN2            As String * 10      '棚番２
Dim TANABAN3            As String * 10      '棚番３
Dim MUKE_NAME           As String * 24      '出庫先名称
Dim CYU_KBN             As String * 1       '注文区分
Dim CYU_KBN_NAME        As String * 40      '注文区分名称
Dim ORIGIN1             As String * 10      '原産国１
Dim ORIGIN2             As String * 10      '原産国２
Dim BIKOU2              As String * 40      '備考２
Dim HAN_KBN             As String * 1       '販売区分
Dim CHOKU_KBN           As String * 1       '直送区分
Dim UNIT_ID_NO          As String * 12      'ﾕﾆｯﾄ修理ID-NO
Dim ZAIKO_HIKIATE       As String * 3       '在庫引当順序
Dim GOKON_KANRI_NO      As String * 8       '合梱管理番号
Dim JYUCHU_ZAN          As String * 7       '受注残数量
Dim KYOKYU_KBN          As String * 1       '供給区分
Dim SHOHIN_SYUSI        As String * 8       '商品化納入先収支
Dim S_SHISAN_SYUSI      As String * 8       '商品化納品資産管理収支ｺｰﾄﾞ
Dim S_HOJYO_SYUSI       As String * 8       '商品化納品補助収支ｺｰﾄﾞ
Dim BIKOU1              As String * 40      '備考１
Dim CHOHA_KBN           As String * 1       '帳端区分
Dim JYU_HIN_NO          As String * 40      '受注品目番号
Dim HIN_NAME            As String * 40      '品名
Dim HIN_CHANGE_KBN      As String * 1       '品番変更区分
Dim MODULE_EXCHANGE     As String * 1       'モジュール交換区分
Dim ZAIKO_SYUSI         As String * 8       '残在庫まとめ在庫収支コード
Dim ZAN_SHISAN_SYUSI    As String * 8       '残在庫まとめ資産管理収支ｺｰﾄﾞ
Dim ZAN_HOJYO_SYUSI     As String * 8       '残在庫まとめ補助収支ｺｰﾄﾞ
Dim NOUKI_YMD           As String * 8       '指定納期
Dim SERVICE_KANRI_NO    As String * 9       'サービス会社管理番号
Dim KISHU_CODE          As String * 3       '機種品目コード
Dim ENVIRONMENT_KBN     As String * 1       '環境規格部品区分
Dim SS_CODE             As String * 8       '直送先コード
Dim KEPIN_KAIJYO        As String * 1       '欠品解消区分


Dim wkSyukaRec      As wkSyukaRec_tag










Dim Upd_com             As Integer          '2008.02.23

Dim wkTemp              As String


Dim WK_Y_QTY            As Long             '2009.04.14
Dim WK_Qty              As Long             '2009.04.14
Dim WK_E_QTY            As Long             '2009.04.14

Dim WORK_SOKO           As String * 2       '2009.04.14

Dim SUMI_QTY            As Long             '2009.04.14
Dim MI_QTY              As Long             '2009.04.14

'2011.01.19
Dim GENSAN_CNT          As Integer
Dim com                 As Integer
Dim GENSANKOKU          As String * 20

Dim Loop_Cnt            As Integer

'2011.01.19



    Syuka_Update_Proc = True



    Fast_Flg = True

    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)


    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")

    Do Until EOF(FileNo)
'        Line Input #FileNo, wkText
        Get #FileNo, , wkSyukaRec
        
        
        
'        If StrConv(wkSyukaRec.CRLF, vbUnicode) <> vbCrLf Then
'            Call NG_File_Make_Proc
'            Exit Do
'        End If
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
    
    
'        Length = 1
'        JGYOBA = Mid(wkText, Length, Len(JGYOBA))                   '事業場
        JGYOBA = StrConv(wkSyukaRec.JGYOBA, vbUnicode)
        
        
'        Length = Length + Len(JGYOBA)
'        DATA_KBN = Mid(wkText, Length, Len(DATA_KBN))               'データ区分
        DATA_KBN = StrConv(wkSyukaRec.DATA_KBN, vbUnicode)
        
        
        
'        Length = Length + Len(DATA_KBN)
'        TORI_KBN = Mid(wkText, Length, Len(TORI_KBN))               '取引区分
        TORI_KBN = StrConv(wkSyukaRec.TORI_KBN, vbUnicode)
    
'        Length = Length + Len(TORI_KBN)
'        ID_NO = Mid(wkText, Length, Len(ID_NO))                     'ID-NO
        ID_NO = StrConv(wkSyukaRec.ID_NO, vbUnicode)
    
'        Length = Length + Len(ID_NO)
'        KAIKEI_JGYOBA = Mid(wkText, Length, Len(KAIKEI_JGYOBA))     '会計用事業場ｺｰﾄﾞ
        KAIKEI_JGYOBA = StrConv(wkSyukaRec.KAIKEI_JGYOBA, vbUnicode)
    
'        Length = Length + Len(KAIKEI_JGYOBA)
'        SHISAN_JGYOBA = Mid(wkText, Length, Len(SHISAN_JGYOBA))     '資産管理事業場ｺｰﾄﾞ
        SHISAN_JGYOBA = StrConv(wkSyukaRec.SHISAN_JGYOBA, vbUnicode)
    
'        Length = Length + Len(SHISAN_JGYOBA)
'        HIN_NO = Mid(wkText, Length, Len(HIN_NO))                   '品目番号
        HIN_NO = StrConv(wkSyukaRec.HIN_NO, vbUnicode)
    
'        Length = Length + Len(HIN_NO)
'        DEN_NO = Mid(wkText, Length, Len(DEN_NO))                   '伝票番号
        DEN_NO = StrConv(wkSyukaRec.DEN_NO, vbUnicode)
    
'        Length = Length + Len(DEN_NO)
'        SURYO = Mid(wkText, Length, Len(SURYO))                     '出庫数量
        SURYO = StrConv(wkSyukaRec.SURYO, vbUnicode)
    
'        Length = Length + Len(SURYO)
'        MUKE_CODE = Mid(wkText, Length, Len(MUKE_CODE))             '出庫先
        MUKE_CODE = StrConv(wkSyukaRec.MUKE_CODE, vbUnicode)
    
'        Length = Length + Len(MUKE_CODE)
'        SYUKO_SYUSI = Mid(wkText, Length, Len(SYUKO_SYUSI))         '出庫収支
        SYUKO_SYUSI = StrConv(wkSyukaRec.SYUKO_SYUSI, vbUnicode)
    
'        Length = Length + Len(SYUKO_SYUSI)
'        SHISAN_SYUSI = Mid(wkText, Length, Len(SHISAN_SYUSI))       '資産管理用在庫収支ｺｰﾄﾞ
        SHISAN_SYUSI = StrConv(wkSyukaRec.SHISAN_SYUSI, vbUnicode)
    
    
'        Length = Length + Len(SHISAN_SYUSI)
'        HOJYO_SYUSI = Mid(wkText, Length, Len(HOJYO_SYUSI))         '補助在庫収支ｺｰﾄﾞ
        HOJYO_SYUSI = StrConv(wkSyukaRec.HOJYO_SYUSI, vbUnicode)
        
'        Length = Length + Len(HOJYO_SYUSI)
'        SYUKO_YMD = Mid(wkText, Length, Len(SYUKO_YMD))             '出庫日付
        SYUKO_YMD = StrConv(wkSyukaRec.SYUKO_YMD, vbUnicode)
    
'        Length = Length + Len(SYUKO_YMD)
'        TANKA = Mid(wkText, Length, Len(TANKA))                     '単価
        TANKA = StrConv(wkSyukaRec.TANKA, vbUnicode)
    
'        Length = Length + Len(TANKA)
'        ODER_NO = Mid(wkText, Length, Len(ODER_NO))                 'オーダー番号
        ODER_NO = StrConv(wkSyukaRec.ODER_NO, vbUnicode)
    
'        Length = Length + Len(ODER_NO)
'        ITEM_NO = Mid(wkText, Length, Len(ITEM_NO))                 'アイテム番号
        ITEM_NO = StrConv(wkSyukaRec.ITEM_NO, vbUnicode)
    
'        Length = Length + Len(ITEM_NO)
'        ODER_NO_R = Mid(wkText, Length, Len(ODER_NO_R))             'オーダー略号
        ODER_NO_R = StrConv(wkSyukaRec.ODER_NO_R, vbUnicode)
    
'        Length = Length + Len(ODER_NO_R)
'        KOSO_KEITAI = Mid(wkText, Length, Len(KOSO_KEITAI))         '個装形態
        KOSO_KEITAI = StrConv(wkSyukaRec.KOSO_KEITAI, vbUnicode)
    
'        Length = Length + Len(KOSO_KEITAI)
'        SYUKA_YMD = Mid(wkText, Length, Len(SYUKA_YMD))             '出荷日
        SYUKA_YMD = StrConv(wkSyukaRec.SYUKA_YMD, vbUnicode)
    
'        Length = Length + Len(SYUKA_YMD)
'        TANABAN1 = Mid(wkText, Length, Len(TANABAN1))               '棚番１
        TANABAN1 = StrConv(wkSyukaRec.TANABAN1, vbUnicode)
    
'        Length = Length + Len(TANABAN1)
'        TANABAN2 = Mid(wkText, Length, Len(TANABAN2))               '棚番２
        TANABAN2 = StrConv(wkSyukaRec.TANABAN2, vbUnicode)
    
'        Length = Length + Len(TANABAN2)
'        TANABAN3 = Mid(wkText, Length, Len(TANABAN3))               '棚番３
        TANABAN3 = StrConv(wkSyukaRec.TANABAN3, vbUnicode)
    
    
    
    
'        Length = Length + Len(TANABAN3)
'        MUKE_NAME = Mid(wkText, Length, Len(MUKE_NAME))             '出庫先名称
        MUKE_NAME = StrConv(wkSyukaRec.MUKE_NAME, vbUnicode)
    
            
    
    
    
    
    
    
'        Length = Length + Len(MUKE_NAME)
'        CYU_KBN = Mid(wkText, Length, Len(CYU_KBN))                 '注文区分
        CYU_KBN = StrConv(wkSyukaRec.CYU_KBN, vbUnicode)
    
    
    
    
    
'        Length = Length + Len(CYU_KBN)
'        CYU_KBN_NAME = Mid(wkText, Length, Len(CYU_KBN_NAME))       '注文区分名称
        CYU_KBN_NAME = StrConv(wkSyukaRec.CYU_KBN_NAME, vbUnicode)
        
        
        
'        Length = Length + Len(CYU_KBN_NAME)
'        ORIGIN1 = Mid(wkText, Length, Len(ORIGIN1))                 '原産国１
        ORIGIN1 = StrConv(wkSyukaRec.ORIGIN1, vbUnicode)
    
'        Length = Length + Len(ORIGIN1)
'        ORIGIN2 = Mid(wkText, Length, Len(ORIGIN2))                 '原産国２
        ORIGIN2 = StrConv(wkSyukaRec.ORIGIN2, vbUnicode)
    
'        Length = Length + Len(ORIGIN2)
'        BIKOU2 = Mid(wkText, Length, Len(BIKOU2))                   '備考２
        BIKOU2 = StrConv(wkSyukaRec.BIKOU2, vbUnicode)
    
'        Length = Length + Len(BIKOU2)
'        HAN_KBN = Mid(wkText, Length, Len(HAN_KBN))                 '販売区分
        HAN_KBN = StrConv(wkSyukaRec.HAN_KBN, vbUnicode)
    
'        Length = Length + Len(HAN_KBN)
'        CHOKU_KBN = Mid(wkText, Length, Len(CHOKU_KBN))             '直送区分
        CHOKU_KBN = StrConv(wkSyukaRec.CHOKU_KBN, vbUnicode)
    
'        Length = Length + Len(CHOKU_KBN)
'        UNIT_ID_NO = Mid(wkText, Length, Len(UNIT_ID_NO))           'ﾕﾆｯﾄ修理ID-NO
        UNIT_ID_NO = StrConv(wkSyukaRec.UNIT_ID_NO, vbUnicode)
    
'        Length = Length + Len(UNIT_ID_NO)
'        ZAIKO_HIKIATE = Mid(wkText, Length, Len(ZAIKO_HIKIATE))     '在庫引当順序
        ZAIKO_HIKIATE = StrConv(wkSyukaRec.ZAIKO_HIKIATE, vbUnicode)
    
'        Length = Length + Len(ZAIKO_HIKIATE)
'        GOKON_KANRI_NO = Mid(wkText, Length, Len(GOKON_KANRI_NO))   '合梱管理番号
        GOKON_KANRI_NO = StrConv(wkSyukaRec.GOKON_KANRI_NO, vbUnicode)
    
'        Length = Length + Len(GOKON_KANRI_NO)
'        JYUCHU_ZAN = Mid(wkText, Length, Len(JYUCHU_ZAN))           '受注残数量
        JYUCHU_ZAN = StrConv(wkSyukaRec.JYUCHU_ZAN, vbUnicode)
    
'        Length = Length + Len(JYUCHU_ZAN)
'        KYOKYU_KBN = Mid(wkText, Length, Len(KYOKYU_KBN))           '供給区分
        KYOKYU_KBN = StrConv(wkSyukaRec.KYOKYU_KBN, vbUnicode)
    
'        Length = Length + Len(KYOKYU_KBN)
'        SHOHIN_SYUSI = Mid(wkText, Length, Len(SHOHIN_SYUSI))       '商品化納入先収支
        SHOHIN_SYUSI = StrConv(wkSyukaRec.SHOHIN_SYUSI, vbUnicode)
    
'        Length = Length + Len(SHOHIN_SYUSI)
'        S_SHISAN_SYUSI = Mid(wkText, Length, Len(S_SHISAN_SYUSI))   '商品化納品資産管理収支ｺｰﾄﾞ
        S_SHISAN_SYUSI = StrConv(wkSyukaRec.S_SHISAN_SYUSI, vbUnicode)
    
'        Length = Length + Len(S_SHISAN_SYUSI)
'        S_HOJYO_SYUSI = Mid(wkText, Length, Len(S_HOJYO_SYUSI))     '商品化納品補助収支ｺｰﾄﾞ
        S_HOJYO_SYUSI = StrConv(wkSyukaRec.S_HOJYO_SYUSI, vbUnicode)
    
'        Length = Length + Len(S_SHISAN_SYUSI)
'        BIKOU1 = Mid(wkText, Length, Len(BIKOU1))                   '備考１
        BIKOU1 = StrConv(wkSyukaRec.BIKOU1, vbUnicode)
    
'        Length = Length + Len(BIKOU1)
'        CHOHA_KBN = Mid(wkText, Length, Len(CHOHA_KBN))             '帳端区分
        CHOHA_KBN = StrConv(wkSyukaRec.CHOHA_KBN, vbUnicode)
    
'        Length = Length + Len(CHOHA_KBN)
'        JYU_HIN_NO = Mid(wkText, Length, Len(JYU_HIN_NO))           '受注品目番号
        JYU_HIN_NO = StrConv(wkSyukaRec.JYU_HIN_NO, vbUnicode)
    
'        Length = Length + Len(JYU_HIN_NO)
'        HIN_NAME = Mid(wkText, Length, Len(HIN_NAME))               '品名
        HIN_NAME = StrConv(wkSyukaRec.HIN_NAME, vbUnicode)
    
'        Length = Length + Len(HIN_NAME)
'        HIN_CHANGE_KBN = Mid(wkText, Length, Len(HIN_CHANGE_KBN))   '品番変更区分
        HIN_CHANGE_KBN = StrConv(wkSyukaRec.HIN_CHANGE_KBN, vbUnicode)
    
'        Length = Length + Len(HIN_CHANGE_KBN)
'        MODULE_EXCHANGE = Mid(wkText, Length, Len(MODULE_EXCHANGE)) 'モジュール交換区分
        MODULE_EXCHANGE = StrConv(wkSyukaRec.MODULE_EXCHANGE, vbUnicode)
    
'        Length = Length + Len(MODULE_EXCHANGE)
'        ZAIKO_SYUSI = Mid(wkText, Length, Len(ZAIKO_SYUSI))         '残在庫まとめ在庫収支コード
        ZAIKO_SYUSI = StrConv(wkSyukaRec.ZAIKO_SYUSI, vbUnicode)
    
'        Length = Length + Len(ZAIKO_SYUSI)
'        ZAN_SHISAN_SYUSI = Mid(wkText, Length, Len(ZAN_SHISAN_SYUSI))   '残在庫まとめ資産管理収支ｺｰﾄﾞ
        ZAN_SHISAN_SYUSI = StrConv(wkSyukaRec.ZAN_SHISAN_SYUSI, vbUnicode)
    
'        Length = Length + Len(ZAN_SHISAN_SYUSI)
'        ZAN_HOJYO_SYUSI = Mid(wkText, Length, Len(ZAN_HOJYO_SYUSI)) '残在庫まとめ補助収支ｺｰﾄﾞ
        ZAN_HOJYO_SYUSI = StrConv(wkSyukaRec.ZAN_HOJYO_SYUSI, vbUnicode)
    
'        Length = Length + Len(ZAN_HOJYO_SYUSI)
'        NOUKI_YMD = Mid(wkText, Length, Len(NOUKI_YMD))             '指定納期
        NOUKI_YMD = StrConv(wkSyukaRec.NOUKI_YMD, vbUnicode)
    
'        Length = Length + Len(NOUKI_YMD)
'        SERVICE_KANRI_NO = Mid(wkText, Length, Len(SERVICE_KANRI_NO))   'サービス会社管理番号
        SERVICE_KANRI_NO = StrConv(wkSyukaRec.SERVICE_KANRI_NO, vbUnicode)
    
'        Length = Length + Len(SERVICE_KANRI_NO)
'        KISHU_CODE = Mid(wkText, Length, Len(KISHU_CODE))           '機種品目コード
        KISHU_CODE = StrConv(wkSyukaRec.KISHU_CODE, vbUnicode)
    
'        Length = Length + Len(KISHU_CODE)
'        ENVIRONMENT_KBN = Mid(wkText, Length, Len(ENVIRONMENT_KBN)) '環境規格部品区分
        ENVIRONMENT_KBN = StrConv(wkSyukaRec.ENVIRONMENT_KBN, vbUnicode)
    
'        Length = Length + Len(ENVIRONMENT_KBN)
'        SS_CODE = Mid(wkText, Length, Len(SS_CODE))                 '直送先コード
        SS_CODE = StrConv(wkSyukaRec.SS_CODE, vbUnicode)
    
'        Length = Length + Len(SS_CODE)
'        KEPIN_KAIJYO = Mid(wkText, Length, Len(KEPIN_KAIJYO))       '欠品解消区分
        KEPIN_KAIJYO = StrConv(wkSyukaRec.KEPIN_KAIJYO, vbUnicode)
        
If ID_NO = "700092591973" Then
     Debug.Print ID_NO
End If
        
        Skip_Flg = True
        For i = 0 To UBound(JGYOBU_T)               '収支区分のチェック
            If JGYOBU = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(SYUKO_SYUSI) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
'-------------------------- PPSC取り込みより
        If Trim(CYU_KBN) = "" Then
            wkCHOKU_KBN = ""
        Else
            wkCHOKU_KBN = "1"
        End If
                                                                                                
                                                        
        If Trim(CYU_KBN) = "" Then
            If Trim(MUKE_CODE) = "A1" Or _
                Trim(MUKE_CODE) = "A2" Or _
                Trim(MUKE_CODE) = "A3" Or _
                Trim(MUKE_CODE) = "A4" Or _
                Trim(MUKE_CODE) = "A5" Or _
                Trim(MUKE_CODE) = "A6" Or _
                Trim(MUKE_CODE) = "A7" Then
                CYU_KBN = "3"
            End If
        
            If MUKE_CODE = "22000440" Or _
                MUKE_CODE = "22000441" Or _
                MUKE_CODE = "22000442" Or _
                MUKE_CODE = "22000443" Or _
                MUKE_CODE = "22000444" Or _
                MUKE_CODE = "22000445" Or _
                MUKE_CODE = "22000446" Then
                CYU_KBN = "2"
            End If
        
        End If
                                                        
        If Trim(CYU_KBN) = "" Then
            CYU_KBN = "3"
        End If
    
'-------------------------- PPSC取り込みより
    
    
    
    
    
    
    
    
        '「00036003」の対応2006.06.03
    
        'If JGYOBU = AIRCON Then                        'エアコンは除外 2006.11.10
'        If JGYOBU = AIRCON Or JGYOBU = OVEN Then        'エアコン、電子レンジは除外 2011.05.16
        
        If JGYOBU = AIRCON Or JGYOBU = OVEN Or JGYOBU = REIZOU Or JGYOBU = SHOKUSEN Then         'エアコン、電子レンジは除外 2011.05.16 冷蔵庫追加 2014.12.17 食洗 2015.03.03
        Else
            If Trim(MUKE_CODE) = "00036003" Then
                Skip_Flg = True
            End If
        End If
    
        If Not Skip_Flg Then
                                '出荷予定重複チェック
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_NO)
    
    
            sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Call LOG_OUT(LOG_F, "Y_SYUKA.DAT DUP 事業部=" & JGYOBU & "伝票ＩＤ＝" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'                    Skip_Flg = True
                
                
                    If Fast_Flg Then
                        Open (FileName) For Output As DUP_SYUKANo
'                        Write #DUP_SYUKANo, , , "出荷重複リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS") '2015.11.19
                        Write #DUP_SYUKANo, , , "出荷取込み異常リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
'                        Write #DUP_SYUKANo, "出荷日", "伝票№", "支払先ｺｰﾄﾞ", "倉庫/ＳＳｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"                  '2015.11.19
                        Write #DUP_SYUKANo, "エラー内容", "出荷日", "伝票№", "出荷先ｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"      '2015.11.19
                        Fast_Flg = False
                    End If
                
                
                    Write #DUP_SYUKANo, "＜重複＞",
                    Write #DUP_SYUKANo, SYUKA_YMD,
                    Write #DUP_SYUKANo, DEN_NO,
                    Write #DUP_SYUKANo, MUKE_CODE,
                    Write #DUP_SYUKANo, MUKE_NAME,
                    Write #DUP_SYUKANo, CYU_KBN,
                    Write #DUP_SYUKANo, CYU_KBN_NAME,
                    Write #DUP_SYUKANo, HIN_NO,
                    Write #DUP_SYUKANo, SURYO,
                    Write #DUP_SYUKANo, ID_NO
                
                    Upd_com = BtOpUpdate
                
                Case BtErrKeyNotFound
                    Upd_com = BtOpInsert
                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "出荷予定", 0)                      '2016.06.23
                    Call File_Error(sts, BtOpGetEqual, "出荷予定", 1, Y_SYU_ID)             '2016.06.23
                    Exit Function
            End Select
    
    
    
    
            
            If Not Skip_Flg Then
                
                                '品目マスタのチェック
'                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Soko_T(i, j).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(HS_OUT_SIJREC.HIN_NO, vbUnicode))
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Skip_Flg = True
'                        Call Log_Out(LOG_F, "伝票ID=" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
'                        Exit Function
'                End Select
                
                
                
                If Not Skip_Flg Then
                                                    'トランザクション開始
                    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                        Exit Function
                    End If
                    
                    
                                                '品目マスタチェック
                    If Item_Check_Proc(Out_Mode, JGYOBU, Soko_T(i, j).NAIGAI, HIN_NO, , HIN_NAME) Then
                        GoTo Abort_Tran
                    End If
                    '2012.12.20
                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "0" And StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "1" Then
                        Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_F)
                    End If
                    '2012.12.20
                                                                    
                    wkMUKE_CODE = MUKE_CODE
                                                                    
                                                                    
                    If Len(Trim(SS_CODE)) = 0 Or _
                        IsNumeric(Trim(SS_CODE)) Then
                    Else
                        SS_CODE = ""
                    End If
                                                                    
                                                                    
'-----------    2005.12.30
                    If JGYOBU = AIRCON Then
                        
                        'MTSｺｰﾄﾞの読み替え
                        If GetIni(App.EXEName, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode), App.EXEName, c) Then
                        Else
                            Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, Trim(c))
                            Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
                        End If
                        
                        
                        
                        
                        'エアコンだった場合向け先に直送先をｾｯﾄ  2004.12.01
                        
                        If Trim(MUKE_CODE) = Trim(SS_CODE) Then
                            SS_CODE = ""
                        Else
                            If Len(Trim(SS_CODE)) <> 0 Then
                                MUKE_CODE = SS_CODE
                                SS_CODE = ""
                            End If
                        End If
                    Else
                        
                        '洗濯機の場合、備考１を直送先にセット 2006.03.25
                        If JGYOBU = SENTAKU And SYUKO_SYUSI = "S2" Then
                            If StrComp(ODER_NO, "FAX", vbTextCompare) Then
                            
                                wkSS = ""
                            
                                For k = 1 To Len(BIKOU1)
                                    If IsNumeric(Mid(BIKOU1, k, 1)) Then
                                        wkSS = wkSS & Mid(BIKOU1, k, 1)
                                    Else
                                        Exit For
                                    End If
                                Next k
                            
                                SS_CODE = wkSS
                            End If
                        
                        End If
                        
                        
                        
                        '他の事業部は現状のまま
                        If Len(Trim(SS_CODE)) = 0 Or _
                            IsNumeric(Trim(SS_CODE)) Then
                        Else
                            SS_CODE = ""
                        End If
                    End If
                        
                    Call UniCode_Conv(K0_MTS.MUKE_CODE, MUKE_CODE)
                    Call UniCode_Conv(K0_MTS.SS_CODE, SS_CODE)
                             
                             
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            
                            
                            If JGYOBU = AIRCON Then
                                'エアコンだった場合向け先に直送先で向け先ﾏｽﾀを新規作成  2004.12.01
                            
                                Call UniCode_Conv(MTSREC.NAIGAI, Soko_T(i, j).NAIGAI)
                                Call UniCode_Conv(MTSREC.DATA_KBN, "")
                                Call UniCode_Conv(MTSREC.MUKE_CODE, MUKE_CODE)
                                Call UniCode_Conv(MTSREC.SS_CODE, "")
                                Call UniCode_Conv(MTSREC.MUKE_NAME, MUKE_NAME)
                                Call UniCode_Conv(MTSREC.SS_NAME, "")
                                Call UniCode_Conv(MTSREC.MUKE_DNAME, MUKE_NAME)
                                Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "")
                                Call UniCode_Conv(MTSREC.FILLER, "")
                                
                                Loop_Cnt = 0
                                
                                
                                Do
                                    sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("他端末でデータ使用中です。<MTS.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
                                        
                                        
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                GoTo Abort_Tran
                                            End If
                                        
                                            DoEvents
                                            Sleep (500)
                                       
                                        
                                        
                                        Case Else
                                            'Call File_Error(sts, BtOpInsert, "向け先管理ﾏｽﾀ" & "key=" & StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode) & "-" & StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode), 0)                      '2016.06.23
                                            Call File_Error(sts, BtOpInsert, "向け先管理ﾏｽﾀ" & "key=" & StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode) & "-" & StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode), 1, MTS_ID)               '2016.06.23

'                                            Exit Function      '2015.11.19
                                            GoTo Abort_Tran     '2015.11.19
                                    End Select
                                Loop
                                                        
                                                        
                                                        
                                                        
                            
                            Else
                                '他の事業部は現状のまま
                                If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
                                    Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
                                    Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
                                Else
                                    Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
                                    Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
                                End If
                            End If
                            
                        
                        Case Else
                            'Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ", 0)                      '2016.06.23
                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ", 1, MTS_ID)               '2016.06.23
'                            Exit Function      '2015.11.19
                            GoTo Abort_Tran     '2015.11.19
                    End Select
                                                                    
                                                                    
                                    
                    If HAN_KBN = "2" Then
                        CYU_KBN = "E"
                    
                    
                    End If
                    '注文区分＝6は２に
                    If JGYOBU = SENTAKU And SYUKO_SYUSI = "S2" Then
                        
                        If CYU_KBN = "6" Then
                            CYU_KBN = "2"
                        
                        
                        End If
                    End If
                    
                    
                    '注文区分対象外は1とする
                    If CYU_KBN = "1" Or _
                        CYU_KBN = "2" Or _
                        CYU_KBN = "3" Or _
                        CYU_KBN = "E" Then
                    Else
            
                        CYU_KBN = "1"
                    End If
                    
                    
                    
                    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                    Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                    
                    
                                        
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN)
                    
                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, ID_NO)
                    
                    Call UniCode_Conv(Y_SYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HIN_NO)
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, MUKE_CODE)
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, SS_CODE)
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, SYUKA_YMD)
                    
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, JGYOBA)
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, DATA_KBN)
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, TORI_KBN)
                    Call UniCode_Conv(Y_SYUREC.ID_NO, ID_NO)
                    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, KAIKEI_JGYOBA)
                    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, SHISAN_JGYOBA)
                    
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, HIN_NO)
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, DEN_NO)
                    Call UniCode_Conv(Y_SYUREC.SURYO, SURYO)
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, MUKE_CODE)
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, SYUKO_SYUSI)
                    
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, SHISAN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, HOJYO_SYUSI)
                    
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, SYUKA_YMD)
                    Call UniCode_Conv(Y_SYUREC.TANKA, TANKA)
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, ODER_NO)
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, ITEM_NO)
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, ODER_NO_R)
                    '20011.10.31
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, Left(KOSO_KEITAI, 10))
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, SYUKA_YMD)
                    
                    
                    If TANA_SPACE Then
                    
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                        
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, TANABAN1)
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, TANABAN2)
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, TANABAN3)
                    End If
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, MUKE_NAME)
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN)
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_NAME)
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, ORIGIN1)
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, ORIGIN2)
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, BIKOU2)
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, HAN_KBN)
                    
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, CHOKU_KBN)
                    
    
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, UNIT_ID_NO)
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, ZAIKO_HIKIATE)
                    
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, GOKON_KANRI_NO)
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, JYUCHU_ZAN)
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, KYOKYU_KBN)
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, SHOHIN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, S_SHISAN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, S_HOJYO_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, BIKOU1)
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, CHOHA_KBN)
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, JYU_HIN_NO)
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, HIN_NAME)
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, HIN_CHANGE_KBN)
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, MODULE_EXCHANGE)
                    
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, ZAIKO_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, ZAN_SHISAN_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, ZAN_HOJYO_SYUSI)
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, NOUKI_YMD)
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, SERVICE_KANRI_NO)
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, KISHU_CODE)
                    
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, ENVIRONMENT_KBN)
                    
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, SS_CODE)
                    
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, KEPIN_KAIJYO)
                    Call UniCode_Conv(Y_SYUREC.HIN_NAI, "")
                    
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    
                    
                    If Upd_com = BtOpInsert Then    '2008.02.23
                        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                        Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
                        Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
                        Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
                        Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")               '2006.07.20
                        
                        
                        Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")              '2006.09.07
                        Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")            '2006.09.07
                        
                        Call UniCode_Conv(Y_SYUREC.H_IO_KBN, "")
                        Call UniCode_Conv(Y_SYUREC.H_SOKO_CODE, "")
                        
                        
                        Call UniCode_Conv(Y_SYUREC.FILLER, "")
                    End If
            
                    Loop_Cnt = 0
    
                    Do
'                        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        sts = BTRV(Upd_com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                If ans = vbCancel Then
'                                    GoTo Abort_Tran
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    GoTo Abort_Tran
                                End If
                            
                                DoEvents
                                Sleep (500)
                            
                            
                            
                            Case BtErrDEAD_LOCK
                                GoTo Abort_Tran
                            Case Else
'                                Call File_Error(sts, BtOpInsert, "出荷予定")
                                'Call File_Error(sts, Upd_com, "出荷予定", 0)                           '2016.06.23
                                Call File_Error(sts, BtOpGetEqual, "出荷予定", 1, Y_SYU_ID)             '2016.06.23
                                GoTo Abort_Tran
                        End Select
                    Loop
    
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
        
        
        
        
                    Out_Cnt = Out_Cnt + 1
                    lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                    DoEvents
        
        
        
        
        
        
        
                    If SYUKA_LOG_ON Then
                        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
                    End If
                
                
                
                    '入荷振替   2009.04.14
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G11" Or Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G22" Then
                        
                        
                        
                        
                        
                        ' "00023410"を追加 2009.06.25 "00021397"を追加　2012.04.06
                        If (Trim(StrConv(Y_SYUREC.JGYOBA, vbUnicode)) = "00023510" Or Trim(StrConv(Y_SYUREC.JGYOBA, vbUnicode)) = "00023410" Or Trim(StrConv(Y_SYUREC.JGYOBA, vbUnicode)) = "00021397") Then
                            If Trim(StrConv(Y_SYUREC.DATA_KBN, vbUnicode)) = "7" Then
                                If Trim(StrConv(Y_SYUREC.TORI_KBN, vbUnicode)) = "19" Then
'                                    If Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "00" Or _
'                                        Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "01" Or _
'                                        Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "07" Or _
'                                        Trim(StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode)) = "08" Then
                                                                    '入荷データ作成
                                        Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                                        Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                                        Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
                                        Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                                        Call UniCode_Conv(Y_NYUREC.TEXT_NO, Right(ID_NO, 9))
                                
                                
                                        Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                                        Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_NO)
                                        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                                        Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(SURYO), "0000000"))
                                        Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                                        Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                                        Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, SYUKA_YMD)
                                        Call UniCode_Conv(Y_NYUREC.TANKA, "")
                                        Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                                        Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                                        Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, SYUKA_YMD)
                                        Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                                        Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                                        Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                                        Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                                        Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                                        Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                                        Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                                        Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                                        Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                                        Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                                        Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                                        Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                                        Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                                        Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                                        Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                                        Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                                        Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                                        Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                                        Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
                                
                                
                                        Last_Proc_F = True              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り
                                
                                
                                        '入荷ﾁｪｯｸﾃﾞｰﾀ更新
                                        Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
                                        Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
                                        Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_NO)
                            
                                        WK_Y_QTY = CLng(SURYO)
                            
                            
                                        Loop_Cnt = 0
                                        
                                        Do
                                            sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
                                                        WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
                                                        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
                                                
                                                        Loop_Cnt = 0
                                                        
                                                        Do
                                                        
                                                            sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                                            Select Case sts
                                                                Case BtNoErr
                                                                    Exit Do
                                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                                    Beep
'                                                                    ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                                                    If ans = vbCancel Then
'                                                                        Exit Function
'                                                                    End If
                                                                
                                                                
                                                                    Loop_Cnt = Loop_Cnt + 1
                                                                    If Loop_Cnt > 5 Then
                                                                        GoTo Abort_Tran
                                                                    End If
                                                                
                                                                    DoEvents
                                                                    Sleep (500)
                                                                
                                                                
                                                                
                                                                Case BtErrDEAD_LOCK
                                                                    'Exit Function          '2015.11.19
                                                                    GoTo Abort_Tran         '2015.11.19
                                                                Case Else
                                                                    'Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)                        '2016.06.23
                                                                    Call File_Error(sts, BtOpUpdate, "入荷ﾁｪｯｸﾃﾞｰﾀ", 1, J_NYU_ID)               '2016.06.23
                                                                    'Exit Function          '2015.11.19
                                                                    GoTo Abort_Tran         '2015.11.19
                                                            End Select
                                                        
                                                        Loop
                                                        WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    Else
                                                        
                                                        Loop_Cnt = 0
                                                        Do
                                                            sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                                            Select Case sts
                                                                Case BtNoErr
                                                                    Exit Do
                                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                                    Beep
'                                                                    ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                                                    If ans = vbCancel Then
'                                                                        Exit Function
'                                                                    End If
                                                                
                                                                
                                                                    Loop_Cnt = Loop_Cnt + 1
                                                                    If Loop_Cnt > 5 Then
                                                                        GoTo Abort_Tran
                                                                    End If
                                                                
                                                                    DoEvents
                                                                    Sleep (500)
                                                                
                                                                
                                                                Case BtErrDEAD_LOCK
                                                                    
                                                                    'Exit Function      '2015.11.19
                                                                    GoTo Abort_Tran     '2015.11.19
                                                            Case Else
                                                                    'Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)                        '2016.06.23
                                                                    Call File_Error(sts, BtOpDelete, "入荷ﾁｪｯｸﾃﾞｰﾀ", 1, J_NYU_ID)               '2016.06.23
                                                                    'Exit Function
                                                                    GoTo Abort_Tran     '2015.11.19
                                                            End Select
                                                        Loop
                                                        WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                                                    End If
                                            
                                                    Exit Do
                                                Case BtErrKeyNotFound
                                                    WK_E_QTY = 0
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                    Beep
'                                                    ans = MsgBox("他端末でデータ使用中です。<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                                    If ans = vbCancel Then
'                                                        Exit Function
'                                                   End If
                                                
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                               
                                                
                                                Case BtErrDEAD_LOCK
                                                    'Exit Function      '2015.11.19
                                                    GoTo Abort_Tran     '2015.11.19
                                                Case Else
                                                    'Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ", 0)                  '2016.06.23
                                                    Call File_Error(sts, BtOpGetEqual, "入荷ﾁｪｯｸﾃﾞｰﾀ", 1, J_NYU_ID)         '2016.06.23
                                                    'Exit Function      '2015.11.19
                                                    GoTo Abort_Tran     '2015.11.19
                                            End Select
                                        Loop
                                                            '先行入荷数（入荷実績数）
                                        Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
                                
                                                            '予算単位元
                                        Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(Y_SYUREC.KEY_MUKE_CODE, vbUnicode))
                                                            '予算単位先
                                        Call UniCode_Conv(Y_NYUREC.YOSAN_TO, "")
                                                            '標準棚番
                                        Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                                StrConv(ITEMREC.ST_DAN, vbUnicode))
                                        Call UniCode_Conv(Y_NYUREC.HIN_NAI, "")
                                                            'H倉庫 2006.10.17
                                        Call UniCode_Conv(Y_NYUREC.H_SOKO, StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode))
                        
                                                            '入荷リスト出力フラグ   2007.06.12
                                        Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, " ")
                        
                                        
                                        
                                
'                If Trim(ORIGIN1) = "" Then
                
'''2011.01.19
'''                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode))
                    
                    
                    
                    
                    Call UniCode_Conv(K0_GENSAN.JGYOBU, JGYOBU)
                    Call UniCode_Conv(K0_GENSAN.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(K0_GENSAN.HIN_GAI, HIN_NO)
                    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")


                    com = BtOpGetGreaterEqual
                    
                    GENSAN_CNT = 0
                    
                    
                    GENSANKOKU = ""
                    
                    Do
                        DoEvents
                    
                        sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                        Select Case sts
                            Case BtNoErr
                            
                                If StrConv(GENSANREC.JGYOBU, vbUnicode) <> StrConv(Y_NYUREC.JGYOBU, vbUnicode) Or _
                                    StrConv(GENSANREC.NAIGAI, vbUnicode) <> StrConv(Y_NYUREC.NAIGAI, vbUnicode) Or _
                                    StrConv(GENSANREC.HIN_GAI, vbUnicode) <> StrConv(Y_NYUREC.HIN_NO, vbUnicode) Then
                                    Exit Do
                                End If
                            
                            
                                GENSAN_CNT = GENSAN_CNT + 1
                                If GENSAN_CNT > 1 Then
                                    GENSANKOKU = ""
                                    Exit Do
                                End If
                            
                            
                                GENSANKOKU = StrConv(GENSANREC.GENSANKOKU, vbUnicode)
                            
                            Case BtErrEOF
                                Exit Do
                            
                            Case Else
                                'Call File_Error(sts, com, "原産国マスタ", 0)               '2016.06.23
                                Call File_Error(sts, com, "原産国マスタ", 1, GENSAN_ID)     '2016.06.23
                                'Exit Function      '2015.11.19
                                GoTo Abort_Tran     '2015.11.19
                        End Select
                    
                        com = BtOpGetNext
                                    
                    Loop
                    
                    
                    
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, "")
                    If GENSAN_CNT = 1 Then
                    
                        Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)
                    End If
                    
'''2011.01.19
                    
                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, StrConv(ITEMREC.TORI_GEN_GENSANKOKU, vbUnicode))
                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, StrConv(ITEMREC.TORI_SHIIRE_WORK_CENTER, vbUnicode))



'               Else
'
'                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, ORIGIN1)
'                    Call UniCode_Conv(Y_NYUREC.GEN_GENSANKOKU, "")
'                    Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, "")
'
'                End If
                
                
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, "")
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, "")
                                        
                Call UniCode_Conv(Y_NYUREC.ID_NO2, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, "")
                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, "")
                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, "")
                                        
                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "0")
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "8")
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "0")
                
                
                
                WORK_SOKO = "90"
                
                Select Case JGYOBU
                    Case AIRCON, SENTAKU
                    
                    
                    Case Else
                        
                        WORK_SOKO = "81"
                        
                        '2009.06.25
                        If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G22" Then
                            WORK_SOKO = "80"
                        End If
                        '2009.06.25
                End Select
                
                
                
                
                
                
                
                
                
                
                
                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, WORK_SOKO & "010101")
                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, Format(WK_E_QTY, "00000000"))
                                        
                                        
                                        
                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2010")
                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)
                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")
                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")
                                        
                                        
                                        
                                        Call UniCode_Conv(Y_NYUREC.FILLER, "")
                                        
                                        
                                        Loop_Cnt = 0
                                        
                                        
                                        Do
                                            sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
                                            Select Case sts
                                                Case BtNoErr
                                                    Exit Do
                                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                                    Beep
'                                                    ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                                                    If ans = vbCancel Then
'                                                        Exit Function
'                                                    End If
                                                
                                                    Loop_Cnt = Loop_Cnt + 1
                                                    If Loop_Cnt > 5 Then
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                    DoEvents
                                                    Sleep (500)
                                               
                                                
                                                Case Else
                                                    
                                                    
                                                    
                    '2010.05.24
                    If Fast_Flg Then
                        Open (FileName) For Output As DUP_SYUKANo
'                        Write #DUP_SYUKANo, , , "出荷重複リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS") '2015.11.19
                        Write #DUP_SYUKANo, , , "出荷取込み異常リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
'                        Write #DUP_SYUKANo, "出荷日", "伝票№", "支払先ｺｰﾄﾞ", "倉庫/ＳＳｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"                  '2015.11.19
                        Write #DUP_SYUKANo, "エラー内容", "出荷日", "伝票№", "出荷先ｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"      '2015.11.19
                        Fast_Flg = False
                    End If
                    
                    Write #DUP_SYUKANo, "[入荷振替データ]",
                    Write #DUP_SYUKANo, SYUKA_YMD,
                    Write #DUP_SYUKANo, DEN_NO,
                    Write #DUP_SYUKANo, MUKE_CODE,
                    Write #DUP_SYUKANo, MUKE_NAME,
                    Write #DUP_SYUKANo, CYU_KBN,
                    Write #DUP_SYUKANo, CYU_KBN_NAME,
                    Write #DUP_SYUKANo, HIN_NO,
                    Write #DUP_SYUKANo, SURYO,
                    Write #DUP_SYUKANo, ID_NO

'                    Write #DUP_SYUKANo, "[入荷振替データ]"

'                    Call File_Error(sts, BtOpInsert, "入荷予定")
'                    Exit Function
                    Call File_Error(sts, BtOpInsert, "入荷予定", 0)
                    GoTo Loop_Proc
                    '2010.05.24
                                                    
                                                    
                                            End Select
                                        Loop
                                    
                        '------------ 2005.12.30
                                        WORK_SOKO = "90"
                                        
                                        Select Case JGYOBU
                                            Case AIRCON, SENTAKU
                                                Call UniCode_Conv(K0_SOKO.Soko_No, WORK_SOKO)
                                                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                                                Select Case sts
                                                    Case BtNoErr
                                                    Case Else
                                                        Call File_Error(sts, BtOpGetEqual, "倉庫マスタ", 0)
                                                        'Exit Function      '2015.11.19
                                                        GoTo Abort_Tran     '2015.11.19
                                                End Select
                                
                                                If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
                                
                                                    SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    MI_QTY = 0
                                                Else
                                                
                                                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                        MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                        SUMI_QTY = 0
                                                    Else
                                                        SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                        MI_QTY = 0
                                                    End If
                                                End If
                                                
                        '------------ 2005.12.30
                                            
                                            
                                            Case Else
                                                
                                                WORK_SOKO = "81"
                                                
                                                '2009.06.25
                                                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "G22" Then
                                                    WORK_SOKO = "80"
                                                End If
                                                '2009.06.25
                                                
                                                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                    MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    SUMI_QTY = 0
                                                Else
                                                    SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                                                    MI_QTY = 0
                                                End If
                                        End Select
                                        
                                
                        '                Wk_SOKO = KASO_NYUKA_SOKO
                        '                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
                        '                    Wk_SOKO = KASO_SMODOSHI_SOKO
                        '
                        '                End If
                                
                                        '入荷数で在庫データ更新（＋）
                                        If Nyuko_Update_Proc(JGYOBU, _
                                                            Soko_T(i, j).NAIGAI, _
                                                            HIN_NO, _
                                                            StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                                            (WORK_SOKO & "01" & "01" & "01"), _
                                                            YOIN_TU_NYUKA, _
                                                            SUMI_QTY, MI_QTY, _
                                                            WS_NO, WS_NO, 5, _
                                                            StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode) & " 伝№:" & DEN_NO, , , , MENU_NO, , KAMOKU_FURIKAE, _
                                                            StrConv(Y_NYUREC.GENSANKOKU, vbUnicode), _
                                                            StrConv(Y_NYUREC.SHIIRE_WORK_CENTER, vbUnicode), _
                                                            StrConv(Y_NYUREC.ID_NO2, vbUnicode), _
                                                            StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode), _
                                                            StrConv(Y_NYUREC.YOSAN_TO, vbUnicode), Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))) Then
                                            'Exit Function      '2015.11.19
                                            GoTo Abort_Tran     '2015.11.19
                                    
                                        End If
                                    
                                        '前借り数で在庫データ更新（－）
                                        If WK_E_QTY <> 0 Then
                                        '在庫データLOCK
                                            If Zaiko_Lock_Proc((WORK_SOKO & "01" & "01" & "01"), _
                                                                JGYOBU, _
                                                                Soko_T(i, j).NAIGAI, _
                                                                HIN_NO, _
                                                                WS_NO, , , 5) Then
                                                'Exit Function      '2015.11.19
                                                GoTo Abort_Tran     '2015.11.19
                            
                                            End If
                                
                                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                MI_QTY = WK_E_QTY
                                            Else
                                                SUMI_QTY = WK_E_QTY
                                            End If
                                    
                                    
                                            If Syuko_Update_Proc(JGYOBU, _
                                                                Soko_T(i, j).NAIGAI, _
                                                                HIN_NO, _
                                                                StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                                                (WORK_SOKO & "01" & "01" & "01"), _
                                                                YOIN_MAE_SOUSAI, _
                                                                SUMI_QTY, MI_QTY, 0, _
                                                                WS_NO, WS_NO, 5) Then
                                                'Exit Function      '2015.11.19
                                                GoTo Abort_Tran     '2015.11.19
                                
                                            End If
                                    
                                    
                                    
                                    
                                    
                                    
                                        End If
 
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
'                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
Loop_Proc:
    
    Loop
    
        
    Close #DUP_SYUKANo
        
    Syuka_Update_Proc = False

    Exit Function

Abort_Tran:
    
    
    
    If Fast_Flg Then
        Open (FileName) For Output As DUP_SYUKANo
'                        Write #DUP_SYUKANo, , , "出荷重複リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS") '2015.11.19
        Write #DUP_SYUKANo, , , "出荷取込み異常リスト", , "作成日：", Format(Now, "YYYY/MM/DD HH:MM:SS")  '2015.11.19
'                        Write #DUP_SYUKANo, "出荷日", "伝票№", "支払先ｺｰﾄﾞ", "倉庫/ＳＳｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"                  '2015.11.19
        Write #DUP_SYUKANo, "エラー内容", "出荷日", "伝票№", "出荷先ｺｰﾄﾞ", "名称", "注文区分", "注文区分名称", "品番", "数量", "伝票ＩＤ"      '2015.11.19
        Fast_Flg = False
    End If


    Write #DUP_SYUKANo, "＜ﾌｧｲﾙ出力異常＞　sts=" & sts,
    Write #DUP_SYUKANo, SYUKA_YMD,
    Write #DUP_SYUKANo, DEN_NO,
    Write #DUP_SYUKANo, MUKE_CODE,
    Write #DUP_SYUKANo, MUKE_NAME,
    Write #DUP_SYUKANo, CYU_KBN,
    Write #DUP_SYUKANo, CYU_KBN_NAME,
    Write #DUP_SYUKANo, HIN_NO,
    Write #DUP_SYUKANo, SURYO,
    Write #DUP_SYUKANo, ID_NO
    
    Close #DUP_SYUKANo          '2015.11.19
    
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
                                            'ヘッダー印刷（「品番変更リスト」）
Private Sub P_Hin_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).CODE Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "＊＊＊　品番変更リスト　＊＊＊";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print
                                        '明細ヘッダ印刷
    Printer.Print "------- 品番（外部）-------";
    Printer.Print Tab(30);
    Printer.Print "------- 品番（内部）-------";
    Printer.Print
    
    Printer.Print Tab(MGN_L);
    Printer.Print "受信データ";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "マスタ";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "受信データ";
    Printer.Print Tab(MGN_L + 44);
    Printer.Print "マスタ";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "伝票日付";
    Printer.Print Tab(MGN_L + 69);
    Printer.Print "入出庫区";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "伝票№";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "入出数";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "倉";
    Printer.Print Tab(MGN_L + 96);
    Printer.Print "注文区";
    Printer.Print Tab(MGN_L + 103);
    Printer.Print "出荷先"
    Printer.Print

    Lcnt = 7 + MGN_U

End Sub
                                            '明細印刷（「品番変更リスト」）
Private Sub P_Hin_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim Emsg As String
Dim Wqty As Long
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99

    For i = 0 To LBox_Hin.ListCount - 1
        
'        Ldata = LBox_Hin.List(i)
'
'                                        'ヘッダーコントロール
        If Lcnt > LMAX Or _
           B_Jgyobu <> Left(Ldata, 1) Then
            Call P_Hin_Head(Lcnt, Left(Ldata, 1))
            B_Jgyobu = Left(Ldata, 1)
        End If
'
'                                        '明細印刷
'        Ldata = Mid(Ldata, 11, Len(Ldata) - 11)                     '事業部，ﾃｷｽﾄ№，国内外　除外'
'
'        Printer.Print Tab(MGN_L);
'        Printer.Print ChrCut(Ldata, 13);                            '受信ﾃﾞｰﾀ品番（外部）
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "1" Or Right(Ldata, 1) = "2" Then      '外部品番変更？
'            Printer.Print Tab(MGN_L + 15);
'            Printer.Print Work;                                     'マスタ品番（外部）
'        End If
'
'        Printer.Print Tab(MGN_L + 30);
'        Printer.Print ChrCut(Ldata, 13);                            '受信ﾃﾞｰﾀ品番（内部）
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "0" Then                               '内部品番変更？
'            Printer.Print Tab(MGN_L + 44);
'            Printer.Print Work;                                     'マスタ品番（内部）
'        End If
'
'        Printer.Print Tab(MGN_L + 58);                              '伝票日付
'        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);
'
'        Printer.Print Tab(MGN_L + 69);                              '入出庫区分
'        wk_IO = ChrCut(Ldata, 1)
'        Select Case wk_IO
'            Case IO_KBN_URI
'                Printer.Print wk_IO & " " & (IO_KBN_0);
'            Case IO_KBN_NYU
'                Printer.Print wk_IO & " " & (IO_KBN_1);
'            Case IO_KBN_SYU
'                Printer.Print wk_IO & " " & (IO_KBN_2);
'            Case IO_KBN_ZAT
'                Printer.Print wk_IO & " " & (IO_KBN_3);
'            Case Else
'                Printer.Print wk_IO;
'        End Select
'
'        Printer.Print Tab(MGN_L + 78);
'        Printer.Print ChrCut(Ldata, 6);                             '伝票№
'
'        Printer.Print Tab(MGN_L + 85);                              '入出庫数
'        Wqty = CLng(ChrCut(Ldata, 6))
'
'
'        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Format(Wqty, "00000000"), Work)
'
'        Printer.Print Work;
'
'        Printer.Print Tab(MGN_L + 93);
'        Printer.Print ChrCut(Ldata, 2);                             '倉庫区分（ﾎｽﾄ）
'
'        Printer.Print Tab(MGN_L + 96);                              '注文区分
'        Select Case Left(Ldata, 1)
'            Case CYU_KBN_TUK
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
'            Case CYU_KBN_SPO
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
'            Case CYU_KBN_HJU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
'            Case CYU_KBN_BOU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
'            Case Else
'                Printer.Print ChrCut(Ldata, 1);
'        End Select
'
'        Printer.Print Tab(MGN_L + 103);
'        Printer.Print ChrCut(Ldata, 5);                             '支給先／出荷先7
'
'        Printer.Print Tab(MGN_L + 110);                             '変更メッセージ
'        Select Case Left(Ldata, 1)
'            Case "0"
'                Printer.Print "内部変更 ﾏｽﾀ品番入替";
'            Case "1"
'                Printer.Print "外部変更 ﾏｽﾀ品番入替";
'            Case "2"
'                Printer.Print "在庫有！外部変更不可";
'        End Select
        
        Printer.Print LBox_Hin.List(i)
        
        Call LOG_OUT(LOG_F, LBox_Hin.List(i))
        
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '用紙の長辺を上にして印刷
    End If

End Sub

Private Sub Form_Activate()

Dim Ret         As String


Dim i           As Integer
Dim FullPath    As String

    Call NG_File_Make_Proc

    '---------------------------------------------  事業部毎メインループ
    For i = 0 To UBound(JGYOBU_T)
        
        In_Cnt = 0
        Out_Cnt = 0

        lblJGYOBU(i).Caption = JGYOBU_T(i).NAME
        lblJGYOBU(i).ForeColor = JGYOBU_T(i).COLOR

        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents

        FileNo = FreeFile
        FileName = HS_IN_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

        On Error GoTo Error_Proc

        Open FileName For Input As #FileNo

        On Error GoTo 0


        If Nyuka_Update_Proc(JGYOBU_T(i).CODE) Then     '入荷予定データ更新処理

            Unload Me

        End If


        Close #FileNo

        '-----------------------------------------------
    
        FileNo = FreeFile
        FileName = HS_OUT_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
        
        On Error GoTo Error_Proc
        
'        Open fileName For Input As #FileNo
        Open FileName For Binary As #FileNo
    
        On Error GoTo 0
    
    
        If Syuka_Update_Proc(JGYOBU_T(i).CODE) Then  '出荷予定データ更新処理

            Unload Me
        End If
    
    
        Close #FileNo
    
    
    
    
    Next i

    If Not Err_FLg Then
        Call NG_File_Kill_Proc
    End If

    Unload Me

Error_Proc:

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim j           As Integer

Dim c           As String * 128
Dim sts         As Integer


Dim sBuffer     As String * 255
Dim com         As String
    
Dim Max_Soko    As Integer
    
Dim GENSAN_WK   As Variant              '2016.12.28
    
    
    If App.PrevInstance Then
 '       Beep
 '       MsgBox "同一プログラム実行中です。"
        End
    End If


    F1020101.Caption = F1020101.Caption & Last_Update_Day


    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                 '出荷重複データ出力ファイル名取り込み
    If GetIni("FILE", "DUP_SYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "出荷重複データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    DUP_SYUKA_DATA = Trim(c)
                               
    If JGYOB_TB_Set() Then      '事業部の獲得
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
    '---------------------------------------------- *
    '    SYS.INI -- > F102010.INI
    '   2015.03.04
    '---------------------------------------------- *
        
                                
                                
                                
                                '倉庫最大数を取り込み
                                
    If GetIni(App.EXEName, "MAX_SOKO", App.EXEName, c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                    
                                
                                
                                
                                '在庫取り込み用テーブル作成
    ReDim Soko_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '倉庫情報取り込み
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '有効倉庫獲得
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "倉庫情報の獲得に失敗しました。処理を中止して下さい。"
                End
            End If
    
            If Trim(c) = "**" Then  '倉庫指定終了
                Exit Do
            End If
    
    
'            ReDim Preserve JSOKO_T(i).JSOKO_T(0 To j)
            Soko_T(i, j).HS_SOKO = Trim(c)
                                '国内外情報獲得
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "国内外情報の獲得に失敗しました。処理を中止して下さい。"
                End
            End If
            
            Soko_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i
                                
                                
                                '入庫データファイル名の獲得
    If GetIni("FILE", "HS_SIJ_IN", "SYS", c) Then
        Beep
        MsgBox "入庫データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    HS_IN_SIJ = Trim(c)
                                
                                '出庫データファイル名の獲得
    If GetIni("FILE", "HS_SIJ_OUT", "SYS", c) Then
        Beep
        MsgBox "出庫データファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    HS_OUT_SIJ = Trim(c)
                                
                                
                                
                                '「通常入荷」要因の獲得
    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
        Beep
        MsgBox "「通常入荷」要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_TU_NYUKA = Trim(c)
                                
                                '「前借相殺」要因の獲得
    If GetIni("YOIN", "YOIN_MAE_SOUSAI", "SYS", c) Then
        Beep
        MsgBox "「前借相殺」要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_MAE_SOUSAI = Trim(c)
                                
                                '仮想入荷倉庫の獲得
    If GetIni("SYSTEM", "KASO_NYUKA", "SYS", c) Then
        Beep
        MsgBox "仮想入荷倉庫の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    KASO_NYUKA_SOKO = Trim(c)
                                '仮想支給戻し倉庫の獲得
    If GetIni("SYSTEM", "KASO_SMODOSHI ", "SYS", c) Then
        Beep
        MsgBox "仮想入荷倉庫の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    KASO_SMODOSHI_SOKO = Trim(c)
                                
                                
                                'その他向け先（国内）の獲得
    If GetIni(App.EXEName, "ETC_MTS_NAI", App.EXEName, c) Then
        Beep
        MsgBox "その他向け先の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    ETC_MTS_NAI = Trim(c)
                                
                                'その他向け先（海外）の獲得
    If GetIni(App.EXEName, "ETC_MTS_GAI", App.EXEName, c) Then
        Beep
        MsgBox "その他向け先の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    ETC_MTS_GAI = Trim(c)
                                
                                
                                
'---------------------------------------------- 'ﾒﾆｭｰ№の獲得    2007.11.06
    If GetIni(App.EXEName, "MENU_NO", App.EXEName, c) Then
        MENU_NO = ""
    Else
        MENU_NO = RTrim(c)
    End If
                                
                                
                                '洗濯機専用
    If GetIni(App.EXEName, "CENTER", App.EXEName, c) Then
        MyCenter = "O"
    Else
        MyCenter = Trim(c)
    End If
                                
                                
                                
'---------------------------------------------- '科目振替の要因 2009.06.26
    KAMOKU_FURIKAE = YOIN_TU_NYUKA
    If GetIni(App.EXEName, "KAMOKU_FURIKAE", App.EXEName, c) Then
    Else
        KAMOKU_FURIKAE = RTrim(c)
    End If
                                
                                
                                'その他直送先の獲得
'    If GetIni(App.EXEName, "ETC_SS_NAI", "SYS", c) Then
'        Beep
'        MsgBox "その他直送先の獲得に失敗しました。処理を中止して下さい。"
'        End
'    End If
'    ETC_SS_NAI = Trim(c)
                                
                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)


'---------------------------------------------- '棚番設定情報の獲得    2009.03.07
    If GetIni(App.EXEName, "TANA_SPACE", App.EXEName, c) Then
        TANA_SPACE = False
    Else
        If Trim(c) = "1" Then
            TANA_SPACE = True
        Else
            TANA_SPACE = False
        End If
    End If



                                '品名による除外 2011.07.04
    NOT_Hin_Name_F = False
    If GetIni(App.EXEName, "NOT_HIN_NAME", App.EXEName, c) Then
    Else
        NOT_Hin_Name = Split(Trim(c), ",", -1)
        NOT_Hin_Name_F = True
    End If
                                '品名による除外 2011.07.04

'---------------------------------------------- '商品化ﾃﾞﾌｫﾙﾄ    2012.12.20
    If GetIni(App.EXEName, "GOODS_F", App.EXEName, c) Then
        GOODS_F = "0"
    Else
        If Trim(c) = "1" Then
            GOODS_F = "1"
        Else
            GOODS_F = "0"
        End If
    End If
'---------------------------------------------- '商品化ﾃﾞﾌｫﾙﾄ    2012.12.20
    
'---------------------------------------------- '更新対象原産国 2016.12.28
    If GetIni(App.EXEName, "GENSAN", App.EXEName, c) Then
        c = "*"
    End If
    GENSAN_WK = Split(Trim(c), ",", -1)

    For i = 0 To UBound(GENSAN_WK)
    
        ReDim Preserve GENSAN_T(0 To i)
        GENSAN_T(i) = GENSAN_WK(i)
    
    
    Next i
'---------------------------------------------- '更新対象原産国 2016.12.28






                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '構成マスタＯＰＥＮ '2005.12.30
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタ（更新用ワーク）ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '要因マスタＯＰＥＮ
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ   2005.12.30
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '原産国マスタＯＰＥＮ   2010.07.08
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                'PNマスタＯＰＥＮ   2010.09.01
    If PN_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    If Country_Open(BtOpenNomal) Then
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
                                '入荷予定ＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '照合用入荷予定ＯＰＥＮ 2007.06.15
    If Y_GLICS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '出荷予定ＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '入荷ﾁｪｯｸﾃﾞｰﾀＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
'発番マスタＯＰＥＮ ################################################################## 2005/05/16 Add ↓
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
'#################################################################################### 2005/05/16 Add ↑
                                
                                
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '印刷フォント設定
    With NormalFont
        .NAME = F1020101.FontName
        .Size = F1020101.FontSize
    End With
    Set Printer.Font = NormalFont

    Last_Proc_F = False         '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有無フラグクリア


    '仕向け先獲得       2005.12.30
    i = -1
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")
    com = BtOpGetGreater
    SHIMUKE_Flg = False
    
    Do
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SHIMUKE_T(0 To i)
            
            
                SHIMUKE_Flg = True
            
                SHIMUKE_T(i).SHIMUKE_CODE = StrConv(P_CODEREC.C_Code, vbUnicode)
                SHIMUKE_T(i).JGYOBU = StrConv(P_CODEREC.OPTION1, vbUnicode)
                SHIMUKE_T(i).NAIGAI = StrConv(P_CODEREC.OPTION2, vbUnicode)
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                Unload Me
        End Select
    
        com = BtOpGetNext
    Loop
        
    
    
    '仕向け先獲得       2005.12.30


    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    DoEvents
    
'    If Last_Proc_F = True Then              '入荷ﾁｪｯｸﾃﾞｰﾀ削除処理　実行有り？
'        Call Last_Proc
'    End If

                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "倉庫マスタ")              '2016.06.23
            Call File_Error(sts, BtOpClose, "倉庫マスタ", 1, SOKO_ID)   '2016.06.23
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "品目マスタ")              '2016.06.23
            Call File_Error(sts, BtOpClose, "品目マスタ", 1, ITEM_ID)   '2016.06.23
        End If
    End If
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "構成マスタ")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "構成マスタ", 1, P_COMPO_ID)    '2016.06.23
        End If
    End If
                                            '品目マスタ（更新用ワーク）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "品目マスタ")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "品目マスタ", ITEM_ID)          '2016.06.23
        End If
    End If
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "向け先管理マスタ")            '2016.06.23
            Call File_Error(sts, BtOpClose, "向け先管理マスタ", 1, MTS_ID)  '2016.06.23
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "要因マスタ")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "要因マスタ", 1, YOIN_ID)         '2016.06.23
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "担当者マスタ")                '2016.06.23
            Call File_Error(sts, BtOpClose, "担当者マスタ", 1, TANTO_ID)      '2016.06.23
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "コードマスタ")                '2016.06.23
            Call File_Error(sts, BtOpClose, "コードマスタ", 1, P_CODE_ID)     '2016.06.23
        End If
    End If
                                            '原産国マスタＣＬＯＳＥ 2010.07.08
    sts = BTRV(BtOpClose, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "原産国マスタ")                '2016.06.23
            Call File_Error(sts, BtOpClose, "原産国マスタ", 1, GENSAN_ID)     '2016.06.23
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "在庫データ")                      '2016.06.23
            Call File_Error(sts, BtOpClose, "在庫データ", 1, ZAIKO_ID)          '2016.06.23
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "在庫移動歴")                      '2016.06.23
            Call File_Error(sts, BtOpClose, "在庫移動歴", 1, IDO_ID)            '2016.06.23
        End If
    End If
                                            '入荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "入荷予定")                        '2016.06.23
            Call File_Error(sts, BtOpClose, "入荷予定", 1, Y_NYU_ID)            '2016.06.23
        End If
    End If
                                            '照合用入荷予定ＣＬＯＳＥ   2007.06.16
    sts = BTRV(BtOpClose, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "照合用入荷予定")                  '2016.06.23
            Call File_Error(sts, BtOpClose, "照合用入荷予定", 1, Y_GLICS_ID)    '2016.06.23
        End If
    End If
                                            '出荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "出荷予定")                        '2016.06.23
            Call File_Error(sts, BtOpClose, "出荷予定", 1, Y_SYU_ID)            '2016.06.23
        End If
    End If
                                            '入荷ﾁｪｯｸﾃﾞｰﾀＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            'Call File_Error(sts, BtOpClose, "入荷ﾁｪｯｸﾃﾞｰﾀ")                    '2016.06.23
            Call File_Error(sts, BtOpClose, "入荷ﾁｪｯｸﾃﾞｰﾀ", J_NYU_ID)           '2016.06.23
        End If
    End If
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020101 = Nothing

    End
End Sub

Private Function Item_Check_Proc(Mode As Integer, JGYOBU As String, NAIGAI As String, HIN_GAI As String, _
                                                                                        Optional HIN_NAI As String = "             ", _
                                                                                        Optional HIN_NAME As String = "                         ", _
                                                                                        Optional GENSANKOKU As String = "                    ", _
                                                                                        Optional GEN_GENSANKOKU As String = "                    ", _
                                                                                        Optional SHIIRE_WORK_CENTER As String = "                    ", _
                                                                                        Optional KANKYO_KBN As String = "   ", _
                                                                                        Optional KANKYO_KBN_ST As String = "        ", _
                                                                                        Optional KANKYO_KBN_SURYO As String = "          ") As Integer
'----------------------------------------------------------------------------
'                   「品目マスタ」チェック＆更新処理
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer

Dim HIN_CHANGE  As Integer

    
    
Dim BEF_GAI     As String * 13
Dim BEF_NAI     As String * 13
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
        
Dim i           As Integer
    
    
Dim sBuffer     As String * 255     '2009.01.21
Dim wkTanto     As String           '2009.01.21
    
Dim PN_M_STS    As Integer          '2010.09.01
    
    
    
Dim Loop_Cnt    As Integer          '2011.01.19
    
    
    Item_Check_Proc = True

    HIN_CHANGE = 0
    

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Loop_Cnt = 0

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                


                    If Mode = In_Mode Then          '対内品番変更のチェック
    '                Else
    
                        If Len(Trim(HIN_NAI)) <> 0 Then
                            If Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) <> Trim(HIN_NAI) Then
                                HIN_CHANGE = NAI_CHANGE
                                BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                '内部品番入れ替え
                                Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
                            
                            
                                '担当者更新追加 2009.11.11
                                    
                                                                                        '更新担当者
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '更新日時
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            
                            
                            
                            
                            
                            
                            End If
                        End If
                    
                    
                    
                                        
                        '---------------    2010.07.08  ▽
                        '原産国入れ替えチェック
                        If Len(Trim(GENSANKOKU)) <> 0 Or Len(Trim(GEN_GENSANKOKU)) <> 0 Or Len(Trim(SHIIRE_WORK_CENTER)) <> 0 Then
    '                        If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) <> Trim(GENSANKOKU) Then
                                '原産国入れ替え
                                
                            
                                If Trim(GENSANKOKU) <> "" Then
                                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                                Else
                                    Debug.Print
                                End If
                                
                                
    '                            If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
    '
    '                                Call UniCode_Conv(ITEMREC.GENSANKOKU, GENSANKOKU)
    '
    '
    '                            End If
                                
                                
                                Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                                Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                            
                            
                                '担当者更新追加 2009.11.11
                                    
                                                                                        '更新担当者
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '更新日時
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                
                                
                            
    '                        End If
                        End If
                        '---------------    2010.07.08  △
                    
                    
                        '---------------    2010.07.27  ▽
                        '環境区分チェック
                        If Len(Trim(KANKYO_KBN)) <> 0 Or Len(Trim(KANKYO_KBN_ST)) <> 0 Or Len(Trim(KANKYO_KBN_SURYO)) <> 0 Then
                            
                            
                            
                            If Val(KANKYO_KBN_SURYO) = 0 Then
                            Else
                            
                                
                                '環境区分入れ替え
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                                    
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                                
                                
                                '担当者更新追加 2009.11.11
                                        
                                                                                        '更新担当者
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '更新日時
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            End If
                                
                            
                        End If
                        '---------------    2010.07.08  △
                    
                    
                    
                    End If
                
                
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                
                com = BtOpInsert
                
                PN_M_STS = PN_M_GET(JGYOBU, HIN_GAI, 0)
                Select Case PN_M_STS
                
                    Case False
                    
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")
                    
                    Case True
                        
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")

                
                End Select
                '2010.09.01
                
                
                
                
                Call Rclr_ITEMREC               '2012.02.11
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '事業部
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '国内外
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '品番（外部）
                                                            '品名
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
    
                    
'2009.01.21                If Mode = In_Mode Then  '新規品番時*をセット2008.10.29
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
                    
'                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'                End If
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    
                
                
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))        '品番（内部）
                Else
                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)         '品番（内部）
                End If
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '備考 ホスト倉庫
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '備考 ホスト棚番
'                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '資材コード
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '補充点
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '月平均出荷数
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          'サンプル数
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '最終入荷日付
    
'                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '排他フラグ
'                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '使用子機ＩＤ
'                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '使用中プログラム
    
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '最終照合日付
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '最終照合時在庫数
'                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '元事事業部
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '印刷備考
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '印刷入り数
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Janコード
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '品番読み替え
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '商品化有無
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '個装箱№
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          'ｸﾞﾘｯｸｽ棚番1
                
                
                                
                
                
'*------------------------------------------ 2005.11.15 追加(業務管理項目) ▽
                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '業務管理　 仕入区分
                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           販売区分
                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           収支単位
                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           組立製品
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           標準粗利売価単価　9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           標準粗利売価設定日
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           標準粗利原価単価  9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           標準粗利原価設定日
                                            
                                            
                                                                        '           仕入先情報
                For i = 0 To 2
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             'ｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '仕入単価
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '単価設定日
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ﾛｯﾄ数
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ﾘｰﾄﾞﾀｲﾑ
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ﾘｰﾄﾞﾀｲﾑ
                
                Next i
                                            
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           前月在庫金額
                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           資材区分
                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ラベル貼付
'*------------------------------------------ 2005.11.15 追加(業務管理項目) △

'*------------------------------------------ 2005.11.15 追加(商品ﾗﾍﾞﾙ項目) ▽
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '商品ﾗﾍﾞﾙ   品名
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           備考
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           会社コード
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           機種(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           機種(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           紙
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           プラスチック
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           価格(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           価格(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           価格(3)
                
                
                                                                '           価格(1)
                If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
                End If
                                                                
                                                                
                                                                '           価格(2)
                If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
                End If
                                                                
                                                                
                                                                '           価格(3)
                If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
                End If
                '2010.09.01
                
                
                
                
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           適用機種ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           枚数ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           適用機種備考
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           作業指示
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           備考３
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           事業部コード
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           入り数
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           棚番(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           棚番(2)
                
                
                
'*------------------------------------------ 2008.08.26 新規追加項目一式 ▽
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '収単／担当者コード
                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '在庫管理対象有無 1:対象 0:対象外
    
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           機種(2)
    
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           前月在庫数量
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           最終出荷数
    
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS在庫(S2) 袋井用
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS在庫(P2) 袋井用
                    
                '2010.09.01
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '個装形態
                Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))
                '2010.09.01
    
    
    

    
'2010.09.01
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               'ﾕﾆｯﾄ部品区分
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '国内供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '海外供給部品区分   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '標準単価   2006.07.28


                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      'ﾕﾆｯﾄ部品区分
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '国内供給部品区分   2006.07.28
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '海外供給部品区分   2006.07.28
                Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '標準単価   2006.07.28
'2010.09.01
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '最終仕入先コード   2007.05.29
                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '最終仕入単価       2007.05.29
    
                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'ﾒｰｶｰｺｰﾄﾞ           2007.06.06
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'ﾒｰｶｰ名称           2007.06.06
    
    
                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '再梱包ﾏｰｸ          2007.11.08
    
                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '才数               2008.02.14
    
                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '形式               2008.02.14
                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '材質               2008.02.14
                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              'ﾀﾞﾝﾎﾞｰﾙ　厚さ      2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（W）   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（D）   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 'ﾀﾞﾝﾎﾞｰﾙｻｲｽﾞ（H）   2008.02.14
        
                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '印刷する／しない   2008.02.14
            
        
                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '商品化　工数       2008.02.14
        
                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '商品化　工数原価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '商品化　工数売価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '商品化　単価設定日 2008.02.14
        
    
                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '商品化　資材原価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '商品化　資材売価   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '商品化　単価設定日 2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '輸送箱　出力ﾌﾗｸﾞ   2008.02.14
    
                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '使用テープ種類     2008.02.14
                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '使用テープ長       2008.02.14
    
                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '棚番マーク         2008.04.02
    
    
                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '請求単価　メモ     2008.04.15
    
                '2010.07.08 ▽
                'Call UniCode_Conv(ITEMREC.GENSANKOKU, "")              '原産国             2008.06.11
                Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")              '原産国
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '原産国
                Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))
                '2010.09.01
                
                
                If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Else
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")
                End If
                '2010.07.08 △
    
                '2010.09.01
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
                    
                    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
                    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
                    Select Case sts
                        Case BtNoErr
                            Debug.Print
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(CountryREC.CountryName2, "")
                        Case Else
                    
                            'Call File_Error(sts, BtOpGetEqual, "COUNTRY")                  '2016.06.23
                            Call File_Error(sts, BtOpGetEqual, "COUNTRY", 1, Country_ID)    '2016.06.23
                            Exit Function
                    
                    End Select
                
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
                
                
                End If
                '2010.09.01
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '外装単価 9(8)V99   2008.06.12
                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC加工単価9(8)   2008.06.12
                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU加工単価9(8)     2008.06.12
    
    
                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '生産ロット         2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '分レート           2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '集合梱包           2008.07.07
    
    
                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '単価設定担当者     2008.07.09
    
                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '仕向け先           2008.07.09

                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '請求区分           2008.07.16

                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            'ラベル貼り枚数     2008.07.19

                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '資材件数     　    2008.08.20追加
                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '同梱件数           2008.08.20追加
         

'*------------------------------------------ 2008.08.26 新規追加項目一式 △
                
                
                                
                
                
                
                '↓2009.02.20
                For i = 0 To 9
                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")

                Next i


                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
                '↑2009.02.20
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.STAT, "1")                    '状態区分           2009.01.21
    

                sBuffer = Space(255)                                    '2009.01.21
                If GetComputerNameA(sBuffer, 255) <> 0 Then
                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                Else
                    wkTanto = "???"
                End If

                
                
                
                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '検品ﾒｯｾｰｼﾞ 2009.08.28
                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '請求Ｆ 2009.08.28
                
                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
            
                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                
                
                
                
                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
                
                
                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
                    
                    If Val(KANKYO_KBN_SURYO) <> 0 Then
                    
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                
                    End If
            
                End If
                
                
                                                                        '追加担当者
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
                Else
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
                End If
                '2010.09.01
                                                                        
                                                                        '追加日時
                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                
                
                
                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
    
                
                
                Call UniCode_Conv(ITEMREC.BIKOU20, "")
                
                '2011.07.05
                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
                
                If NOT_Hin_Name_F Then
                    For i = 0 To UBound(NOT_Hin_Name)
                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
                            Exit For
                        End If
                    Next i
                End If
                '2011.07.05
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '更新担当者
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '更新日時
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                                
                
                
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
           
            
            Case BtErrDEAD_LOCK
                Exit Function
                        
            Case Else
                'Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 0)            '2016.06.23
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ", 1, ITEM_ID)    '2016.06.23
                Exit Function
        End Select
    Loop
    
    Loop_Cnt = 0
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
            
            
            
            Case BtErrDEAD_LOCK
                Exit Function
            Case Else
                'Call File_Error(sts, com, "品目マスタ", 0)
                Call File_Error(sts, com, "品目マスタ", 1, ITEM_ID)
                Exit Function
        End Select
    Loop
        
    If SHIMUKE_Flg Then
        If com = BtOpInsert Then
            '構成マスタの追加       2005.12.30
            For i = 0 To UBound(SHIMUKE_T)
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIMUKE_T(i).JGYOBU And _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) = SHIMUKE_T(i).NAIGAI Then
                                                                            '仕向け先ｺｰﾄﾞ
                    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, SHIMUKE_T(i).SHIMUKE_CODE)
                                                                            '事業部
                    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, SHIMUKE_T(i).JGYOBU)
                                                                            '国内外
                    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, SHIMUKE_T(i).NAIGAI)
                                                                            '品番
                    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            'ﾃﾞｰﾀ区分
                    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
                                                                            '追番
                    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
                                                                            '基本クラス
                    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, "")
                                                                            '備考
                    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "")
                    
                    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
                                                                            '更新担当者
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                            '更新日時
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
'                                If ans = vbCancel Then
'                                    Exit Function
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                            
                                DoEvents
                                Sleep (500)
                            
                            
                            Case BtErrDEAD_LOCK
                                Exit Function
                            Case Else
                                'Call File_Error(sts, BtOpInsert, "構成マスタ", 0)              '2016.06.23
                                Call File_Error(sts, BtOpInsert, "構成マスタ", 1, P_COMPO_ID)   '2016.06.23
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If
        
    If HIN_CHANGE <> 0 Then
        LBox_Hin.AddItem JGYOBU & NAIGAI & StrConv(ITEMREC.HIN_GAI, vbUnicode) & BEF_GAI & StrConv(ITEMREC.HIN_NAI, vbUnicode) & BEF_NAI & NAI_CHANGE
    End If

    Item_Check_Proc = False

End Function

Sub NG_File_Make_Proc()
'----------------------------------------------------------------------------
'                   異常終了ファイル出力処理
'----------------------------------------------------------------------------
Dim stream  As Integer                       'ファイル番号
Dim Buf     As String                           '読み込みバッファ
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                'ログファイル名取り込み
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "異常終了ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    
    stream = FreeFile
    Open NG_FILE For Append As stream
    prog = StrConv(App.EXEName, vbUpperCase)
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog)
    Print #stream, Buf
    Close stream
End Sub
Sub NG_File_Kill_Proc()
'----------------------------------------------------------------------------
'                   異常終了ファイル削除処理    2008.10.07
'----------------------------------------------------------------------------
Dim stream  As Integer                       'ファイル番号
Dim Buf     As String                           '読み込みバッファ
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                'ログファイル名取り込み
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "異常終了ファイル名の獲得に失敗しました。処理を中止して下さい。"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    On Error GoTo Err_Proc
    Kill (NG_FILE)

Err_Proc:

End Sub


Private Function Y_GLICS_PUT_PROC(JGYOBU As String, NAIGAI As String, INS_NOW As String, _
                                                                        TEXT_NO As String, _
                                                                        JGYOBU_Code As String, _
                                                                        CYOK_KBN As String, _
                                                                        DEN_DT As String, _
                                                                        IO_KBN As String, _
                                                                        PM_KBN As String, _
                                                                        DEN_SYU As String, _
                                                                        DEN_NO As String, _
                                                                        CYU_KBN As String, _
                                                                        HIN_GAI As String, _
                                                                        HIN_NAI As String, _
                                                                        HIN_NAME As String, _
                                                                        YOTEI_QTY As String, _
                                                                        YOSAN_FROM As String, _
                                                                        YOSAN_TO As String, _
                                                                        HOST_SOKO As String, _
                                                                        HOST_TANA As String, _
                                                                        SYUK_CODE As String, _
                                                                        SYUK_NAME As String, _
                                                                        GENSANKOKU As String, GEN_GENSANKOKU As String, SHIIRE_WORK_CENTER As String, KANKYO_KBN As String, KANKYO_KBN_ST As String, KANKYO_KBN_SURYO As String, ID_NO2 As String, AITESAKI_CODE As String, JYUCHU_YMD As String, SHITEI_NOUKI_YMD As String, MOTO_TEXT_NO As String) As Integer
'----------------------------------------------------------------------------
'           照合用入荷予定ファイル出力処理
'           2007.06.15
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
        
        
Dim Loop_Cnt    As Integer  '2011.01.19
        
    Y_GLICS_PUT_PROC = True
        
    Call UniCode_Conv(Y_GLICSREC.KAN_KBN, KAN_KBN_FIN)
    Call UniCode_Conv(Y_GLICSREC.DT_SYU, "0")
    Call UniCode_Conv(Y_GLICSREC.JGYOBU, JGYOBU)
    Call UniCode_Conv(Y_GLICSREC.NAIGAI, NAIGAI)
    Call UniCode_Conv(Y_GLICSREC.TEXT_NO, TEXT_NO)


    Call UniCode_Conv(Y_GLICSREC.JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.DATA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.TORI_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KAIKEI_JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_JGYOBA, "")
    
    Call UniCode_Conv(Y_GLICSREC.HIN_NO, HIN_GAI)
    Call UniCode_Conv(Y_GLICSREC.DEN_NO, DEN_NO)
    
    
    '2008.01.10 マイナスの対応
    If YOTEI_QTY >= 0 Then
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
    Else
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "000000"))
    End If
    
    Call UniCode_Conv(Y_GLICSREC.MUKE_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANKA, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ITEM_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO_R, "")
    Call UniCode_Conv(Y_GLICSREC.KOSO_KEITAI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKA_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANABAN1, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN2, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN3, "")
    Call UniCode_Conv(Y_GLICSREC.MUKE_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN, CYU_KBN)
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN1, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN2, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU2, "")
    Call UniCode_Conv(Y_GLICSREC.HAN_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.CHOKU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_GLICSREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_GLICSREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU1, "")
    Call UniCode_Conv(Y_GLICSREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAME, HIN_NAME)
    Call UniCode_Conv(Y_GLICSREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_GLICSREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KI_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SS_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.KEPIN_KAIJYO, "")
    
    
    Call UniCode_Conv(Y_GLICSREC.KAN_DT, Format(Now, "YYYYMMDD"))


                        '先行入荷数（入荷実績数）
    Call UniCode_Conv(Y_GLICSREC.BEF_NYU_QTY, "00000000")

                        '予算単位元
    Call UniCode_Conv(Y_GLICSREC.YOSAN_FROM, YOSAN_FROM)
                        '予算単位先
    Call UniCode_Conv(Y_GLICSREC.YOSAN_TO, YOSAN_TO)
                        '標準棚番
    Call UniCode_Conv(Y_GLICSREC.HTANABAN, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAI, HIN_NAI)
                        'H倉庫 2006.10.17
    Call UniCode_Conv(Y_GLICSREC.H_SOKO, HOST_SOKO)

                        '入荷リスト出力フラグ   2007.06.12
    Call UniCode_Conv(Y_GLICSREC.NYU_LIST_OUT, " ")
                        '直送区分
    Call UniCode_Conv(Y_GLICSREC.CYOK_KBN, CYOK_KBN)
                        '入出庫区分
    Call UniCode_Conv(Y_GLICSREC.IO_KBN, IO_KBN)
                        '赤黒区分
    Call UniCode_Conv(Y_GLICSREC.PM_KBN, PM_KBN)
                        '伝票種別
    Call UniCode_Conv(Y_GLICSREC.DEN_SYU, DEN_SYU)
                        '支給先／出荷先
    Call UniCode_Conv(Y_GLICSREC.SYUK_CODE, SYUK_CODE)
                        '支給先／出荷先名
    Call UniCode_Conv(Y_GLICSREC.SYUK_NAME, SYUK_NAME)
                        '挿入年月日
    Call UniCode_Conv(Y_GLICSREC.INS_NOW, INS_NOW)
    
    
    
    
    
    '----------------   2010.07.08 ▽
    Call UniCode_Conv(Y_GLICSREC.GENSANKOKU, GENSANKOKU)                    '原産国名
    Call UniCode_Conv(Y_GLICSREC.GEN_GENSANKOKU, GEN_GENSANKOKU)            '現物表示原産国名
    Call UniCode_Conv(Y_GLICSREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)    '資材仕入先ﾜｰｸｾﾝﾀｰ
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN, KANKYO_KBN)                    '環境種類区分
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_ST, KANKYO_KBN_ST)              '環境種類区分適用開始
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)        '環境種類区分数量
    Call UniCode_Conv(Y_GLICSREC.ID_NO2, ID_NO2)                            'ID_NO
    Call UniCode_Conv(Y_GLICSREC.AITESAKI_CODE, AITESAKI_CODE)              '相手先ｺｰﾄﾞ
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_YMD, JYUCHU_YMD)                    '受注年月日
    Call UniCode_Conv(Y_GLICSREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)        '指定納期年月日
    Call UniCode_Conv(Y_GLICSREC.LIST_OUT_END_F, "")                        '入庫ﾘｽﾄ出力F
    Call UniCode_Conv(Y_GLICSREC.NYUKO_TANABAN, "")                         '入庫棚番
    Call UniCode_Conv(Y_GLICSREC.MAEGARI_SURYO, "")                         '前借相殺数
    '----------------   2010.07.08 △
    
    
    
    '2011.03.23 発生元プログラム
    Call UniCode_Conv(Y_GLICSREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
    '2011.03.23 元テキスト№
    If Trim(MOTO_TEXT_NO) = "" Then
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, "")
    Else
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
    End If
    
    
    
    
    
    
    
    Call UniCode_Conv(Y_GLICSREC.FILLER, "")


    Loop_Cnt = 0

    Do
        sts = BTRV(BtOpInsert, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。<Y_GLICSKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
            
            
            Case BtErrDEAD_LOCK
                Exit Function
            Case Else
                'Call File_Error(sts, BtOpInsert, "入荷予定", 0)                '2016.06.23
                Call File_Error(sts, BtOpInsert, "入荷予定", 1, Y_GLICS_ID)     '2016.06.23
                Exit Function
        End Select
    Loop

    Y_GLICS_PUT_PROC = False

End Function


Private Function MAEGARI_PROC(JGYOBU As String, HIN_GAI As String, YOTEI_QTY As String) As Integer
'----------------------------------------------------------------------------
'           照合用入荷予定ファイル出力処理
'           2018.11.15
'----------------------------------------------------------------------------
Dim com             As Integer
Dim wkYOTEI_QTY     As Long
Dim sts             As Integer
        
    MAEGARI_PROC = True
    '部品前借処理
    Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_J_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_GAI)
                
                                '前借りﾃﾞｰﾀ読込み
    sts = BTRV(BtOpGetEqual, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    Select Case sts
        Case BtNoErr
            com = BtOpUpdate
        Case BtErrKeyNotFound
            com = BtOpInsert
            
        Case Else
            Call File_Error(sts, BtOpGetEqual, "入荷実績データ")
            Exit Function
    End Select
    
    If com = BtOpInsert Then
                                '新規追加
                                                '事業部
        Call UniCode_Conv(J_NYUREC.JGYOBU, JGYOBU)
                                                '国内外
        Call UniCode_Conv(J_NYUREC.NAIGAI, NAIGAI_NAI)
                                                '品目（外部）
        Call UniCode_Conv(J_NYUREC.HIN_GAI, HIN_GAI)
                                                '実績数量
        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(Val(YOTEI_QTY), "00000000"))
                                                '登録日
        Call UniCode_Conv(J_NYUREC.INS_DATE, Format(Now, "YYYYMMDD"))
        
        
        
        Call UniCode_Conv(J_NYUREC.FILLER, "")
    Else
                                                '実績数量
        wkYOTEI_QTY = Val(YOTEI_QTY) + Val(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
        Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(wkYOTEI_QTY, "00000000"))
    End If
    '*------------------------------------------------------'前借りデータ出力
    sts = BTRV(com, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    Select Case sts
        Case BtNoErr
        
        Case Else
            Call File_Error(sts, com, "入荷実績データ")
            Exit Function
                
    End Select

    MAEGARI_PROC = False


End Function
