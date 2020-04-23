VERSION 5.00
Begin VB.Form PC00050org1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "品目マスタコンバート処理"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   9120
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
   ScaleHeight     =   7230
   ScaleWidth      =   9120
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   375
      Index           =   10
      Left            =   6960
      TabIndex        =   30
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全て"
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   29
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   28
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   27
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   26
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   25
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   24
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   23
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   22
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   21
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   5160
      TabIndex        =   19
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　欠品防止支援ログ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   18
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   17
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　　棚卸しデータ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   16
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   5160
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　在庫集計データ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   14
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   13
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "入荷チェックデータ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   12
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5160
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　在庫移動歴＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　過日分出荷予定＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　　出荷予定＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　在庫データ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　品目マスタ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
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
      Top             =   960
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "データコンバート処理"
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4800
   End
End
Attribute VB_Name = "PC00050org1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Function Update_Proc(Mode As Integer) As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim i               As Integer

Dim DISP_INTERVAL   As Long




    Update_Proc = True


    Select Case Mode
        Case 1              '品目へ
            GoTo ITEM_CONV
        Case 2              '在庫へ
            GoTo ZAIKO_CONV
        Case 3              '出荷予定へ
            GoTo Y_SYU_CONV
        Case 4              '過日分出荷予定へ
            GoTo DEL_SYU_CONV
        Case 5              '在庫移動歴へ
            GoTo IDO_CONV
        Case 6              '入荷ﾁｪｯｸへ
            GoTo J_NYU_CONV
        Case 7              '在庫集計へ
            GoTo SUMZ_CONV
        Case 8              '棚卸しへ
            GoTo STOCK_CONV
        Case 9              '欠品防止ﾛｸﾞへ
            GoTo KEPPINLOG_CONV
    End Select


    
    '---------------------------------------------------------  品目マスタ処理
    
    
ITEM_CONV:
    
    MsgLab(1) = "品目マスタコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                '(旧)品目マスタＯＰＥＮ
    If OLD_ITEM_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo ZAIKO_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)品目マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
                                                
        Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(OLD_ITEMREC.JGYOBU, vbUnicode))               '事業部
        Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(OLD_ITEMREC.NAIGAI, vbUnicode))               '国内外
        Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(OLD_ITEMREC.HIN_GAI, vbUnicode))             '品番(外)
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(OLD_ITEMREC.HIN_NAME, vbUnicode))           '品名
        Call UniCode_Conv(ITEMREC.ST_SET_DT, StrConv(OLD_ITEMREC.ST_SET_DT, vbUnicode))         '標準倉庫設定日付
        Call UniCode_Conv(ITEMREC.ST_SOKO, StrConv(OLD_ITEMREC.ST_SOKO, vbUnicode))             '標準棚番　倉庫
        Call UniCode_Conv(ITEMREC.ST_RETU, StrConv(OLD_ITEMREC.ST_RETU, vbUnicode))             '標準棚番　列
        Call UniCode_Conv(ITEMREC.ST_REN, StrConv(OLD_ITEMREC.ST_REN, vbUnicode))               '標準棚番　連
        Call UniCode_Conv(ITEMREC.ST_DAN, StrConv(OLD_ITEMREC.ST_DAN, vbUnicode))               '標準棚番　段
        Call UniCode_Conv(ITEMREC.BEF_SOKO, StrConv(OLD_ITEMREC.BEF_SOKO, vbUnicode))           '前回棚番　倉庫
        Call UniCode_Conv(ITEMREC.BEF_RETU, StrConv(OLD_ITEMREC.BEF_RETU, vbUnicode))           '前回棚番　列
        Call UniCode_Conv(ITEMREC.BEF_REN, StrConv(OLD_ITEMREC.BEF_REN, vbUnicode))             '前回棚番　連
        Call UniCode_Conv(ITEMREC.BEF_DAN, StrConv(OLD_ITEMREC.BEF_DAN, vbUnicode))             '前回棚番　段
                                                                                                '最終入庫日
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, StrConv(OLD_ITEMREC.LAST_NYU_DT, vbUnicode))
                                                                                                '最終出庫日
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, StrConv(OLD_ITEMREC.LAST_SYU_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OLD_ITEMREC.HIN_NAI, vbUnicode))             '品番(内)
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(OLD_ITEMREC.BIKOU_SOKO, vbUnicode))       '備考　ﾎｽﾄ倉庫
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(OLD_ITEMREC.BIKOU_TANA, vbUnicode))       '備考　ﾎｽﾄ棚番
        Call UniCode_Conv(ITEMREC.HOJYU_P, StrConv(OLD_ITEMREC.HOJYU_P, vbUnicode))             '補充点
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, StrConv(OLD_ITEMREC.AVE_SYUKA, vbUnicode))         '月平均出荷数
                
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, StrConv(OLD_ITEMREC.SAMPLE_QTY, vbUnicode))       'ｻﾝﾌﾟﾙ数
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, StrConv(OLD_ITEMREC.LAST_INP_DT, vbUnicode))     '最終入荷日付
        
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, StrConv(OLD_ITEMREC.LAST_CHK_DT, vbUnicode))     '最終照合日付
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, StrConv(OLD_ITEMREC.LAST_CHK_QTY, vbUnicode))   '最終照合時在庫数
        
        Call UniCode_Conv(ITEMREC.BIKOU, StrConv(OLD_ITEMREC.BIKOU, vbUnicode))                 '印刷備考
        Call UniCode_Conv(ITEMREC.IRI_QTY, StrConv(OLD_ITEMREC.IRI_QTY, vbUnicode))             '印刷入り数
        Call UniCode_Conv(ITEMREC.JAN_CODE, StrConv(OLD_ITEMREC.JAN_CODE, vbUnicode))           'JANコード
        Call UniCode_Conv(ITEMREC.HIN_CHANGE, StrConv(OLD_ITEMREC.HIN_CHANGE, vbUnicode))       '品番読み替え
        Call UniCode_Conv(ITEMREC.GOODS_KBN, StrConv(OLD_ITEMREC.GOODS_KBN, vbUnicode))         '商品化有無
        Call UniCode_Conv(ITEMREC.PACKING_NO, StrConv(OLD_ITEMREC.PACKING_NO, vbUnicode))       '個装箱№
        Call UniCode_Conv(ITEMREC.RANK, StrConv(OLD_ITEMREC.RANK, vbUnicode))                   '現在ﾗﾝｸ
        Call UniCode_Conv(ITEMREC.NEW_RANK, StrConv(OLD_ITEMREC.NEW_RANK, vbUnicode))           '新ﾗﾝｸ
        Call UniCode_Conv(ITEMREC.GLICS1_TANA, StrConv(OLD_ITEMREC.GLICS1_TANA, vbUnicode))     'ｸﾞﾘｯｸｽ棚番1
        Call UniCode_Conv(ITEMREC.GLICS2_TANA, StrConv(OLD_ITEMREC.GLICS2_TANA, vbUnicode))     'ｸﾞﾘｯｸｽ棚番2
        Call UniCode_Conv(ITEMREC.GLICS3_TANA, StrConv(OLD_ITEMREC.GLICS3_TANA, vbUnicode))     'ｸﾞﾘｯｸｽ棚番3
        
        
        
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '業務管理　 仕入区分
        Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")                                 '           販売区分
        Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                                     '           収支単位
        Call UniCode_Conv(ITEMREC.G_KUMITATE, "")                                   '           組立製品
        Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")                                  '           標準粗利売価単価　9(8)V99
        Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")                               '           標準粗利売価設定日
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")                                  '           標準粗利原価単価  9(8)V99
        Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")                               '           標準粗利原価設定日
        
        For i = 0 To 2                                                              '仕入先情報
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")                     '           仕入先コード
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")                    '           単価
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")                 '           単価設定日
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")                      '           単価設定日
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")                '           リードタイム
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")            '           最終発注日
            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")           '           最終発注数
        
        Next i
    
        Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")                              '           前月在庫金額
        Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '           資材区分
        Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")                                  '           ﾗﾍﾞﾙ貼り付け
                
        
        Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                                 '品名E
        Call UniCode_Conv(ITEMREC.L_BIKOU, "")                                      '備考
        Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                                '会社名
        Call UniCode_Conv(ITEMREC.L_KISHU1, "")                                     '機種(1)
        Call UniCode_Conv(ITEMREC.L_KISHU2, "")                                     '機種(2)
        Call UniCode_Conv(ITEMREC.L_KISHU3, "")                                     '機種(3)
        Call UniCode_Conv(ITEMREC.L_PAPER, "")                                      '紙
        Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                                    'ﾌﾟﾗｽﾁｯｸ
        Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                                    '価格(1)
        Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                                    '価格(2)
        Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                                    '価格(3)
        Call UniCode_Conv(ITEMREC.L_LABEL, "")                                      '適用機種ﾗﾍﾞﾙ
        Call UniCode_Conv(ITEMREC.L_MAISU, "")                                      'ﾗﾍﾞﾙ枚数
        Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                                '適用機種備考
        Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                                '作業指示
        Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                                     '備考(3)
        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                                '事業部名
        Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                                    '入り数
        Call UniCode_Conv(ITEMREC.L_TANA1, "")                                      '棚番(1)
        Call UniCode_Conv(ITEMREC.L_TANA2, "")                                      '棚番(2)
        
        
        
        Call UniCode_Conv(ITEMREC.S_TANTO, "")                                      '収単／担当者
        Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)                            '在庫管理対象
        
        Do
            sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(0).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If
    '---------------------------------------------------------  在庫データの処理

ZAIKO_CONV:

    MsgLab(1) = "在庫データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(Count, "#0")
                                        
                                '(旧)在庫データＯＰＥＮ
    If OLD_ZAIKO_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo Y_SYU_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)在庫データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(1).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(ZAIKOREC.Soko_No, StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode))       '倉庫№
        Call UniCode_Conv(ZAIKOREC.Retu, StrConv(OLD_ZAIKOREC.Retu, vbUnicode))             '列
        Call UniCode_Conv(ZAIKOREC.Ren, StrConv(OLD_ZAIKOREC.Ren, vbUnicode))               '連
        Call UniCode_Conv(ZAIKOREC.Dan, StrConv(OLD_ZAIKOREC.Dan, vbUnicode))               '段
                                                
        Call UniCode_Conv(ZAIKOREC.JGYOBU, StrConv(OLD_ZAIKOREC.JGYOBU, vbUnicode))         '事業部ｺｰﾄﾞ
        Call UniCode_Conv(ZAIKOREC.NAIGAI, StrConv(OLD_ZAIKOREC.NAIGAI, vbUnicode))         '国内外
        Call UniCode_Conv(ZAIKOREC.HIN_GAI, StrConv(OLD_ZAIKOREC.HIN_GAI, vbUnicode))       '品番（外部）
        
        Call UniCode_Conv(ZAIKOREC.GOODS_ON, StrConv(OLD_ZAIKOREC.GOODS_ON, vbUnicode))     '商品／未商品
        Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(OLD_ZAIKOREC.NYUKA_DT, vbUnicode))     '入荷日
        Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(OLD_ZAIKOREC.NYUKO_DT, vbUnicode))     '入庫日
        Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(OLD_ZAIKOREC.HIN_NAI, vbUnicode))       '品番(内部)
        
        Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, StrConv(OLD_ZAIKOREC.YUKO_Z_QTY, vbUnicode)) '有効在庫
        
        Call UniCode_Conv(ZAIKOREC.LOCK_F, "")                                              '排他フラグ
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                                              '使用中端末
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                                              '使用中プログラム
        
        Call UniCode_Conv(ZAIKOREC.GOODS_YMD, StrConv(OLD_ZAIKOREC.GOODS_YMD, vbUnicode))   '商品化日付
        
        Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, "")                                         '仕入先ｺｰﾄﾞ
        Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, "")                                        '仕入単価
        Call UniCode_Conv(ZAIKOREC.KEIJYO_YM, "")                                           '計上年月
                
        Call UniCode_Conv(ZAIKOREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "在庫データ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(1).Caption = Format(Count, "#0")
    
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  出荷予定データの処理
Y_SYU_CONV:


    MsgLab(1) = "出荷予定データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(Count, "#0")
                                        
                                        
                                '(旧)出荷予定データＯＰＥＮ
    If OLD_Y_SYU_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo DEL_SYU_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_Y_SYU_POS, OLD_Y_SYUREC, Len(OLD_Y_SYUREC), K0_OLD_Y_SYU, Len(K0_OLD_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)出荷予定データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        
        
        
        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                                                      '使用端末ID
        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                                                      '使用中ﾌﾟﾛｸﾞﾗﾑID
        
        Call UniCode_Conv(Y_SYUREC.KAN_KBN, StrConv(OLD_Y_SYUREC.KAN_KBN, vbUnicode))               '完了区分
        Call UniCode_Conv(Y_SYUREC.DT_SYU, StrConv(OLD_Y_SYUREC.DT_SYU, vbUnicode))                 'ﾃﾞｰﾀ種別
        Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(OLD_Y_SYUREC.JGYOBU, vbUnicode))                 '事業部
        
        Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(OLD_Y_SYUREC.KEY_CYU_KBN, vbUnicode))       '注文区分
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, StrConv(OLD_Y_SYUREC.KEY_ID_NO, vbUnicode))           'ID-NO
        
        Call UniCode_Conv(Y_SYUREC.NAIGAI, StrConv(OLD_Y_SYUREC.NAIGAI, vbUnicode))                 '国内外
        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(OLD_Y_SYUREC.KEY_HIN_NO, vbUnicode))         '品目番号
        
        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(OLD_Y_SYUREC.KEY_MUKE_CODE, vbUnicode))   '得意先ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(OLD_Y_SYUREC.KEY_SS_CODE, vbUnicode))       '直送先ｺｰﾄﾞ
        
        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(OLD_Y_SYUREC.KEY_SYUKA_YMD, vbUnicode))   '出荷日付
        
        Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(OLD_Y_SYUREC.JGYOBA, vbUnicode))                 '事業場
        Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(OLD_Y_SYUREC.DATA_KBN, vbUnicode))             'ﾃﾞｰﾀ区分
        Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(OLD_Y_SYUREC.TORI_KBN, vbUnicode))             '取引区分
        Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(OLD_Y_SYUREC.ID_NO, vbUnicode))                   'ID-NO
        Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(OLD_Y_SYUREC.HIN_NO, vbUnicode))                 '品目番号
        
        Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode))                 '伝票番号
        Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(OLD_Y_SYUREC.SURYO, vbUnicode))                   '数量
        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(OLD_Y_SYUREC.MUKE_CODE, vbUnicode))           '得意先ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(OLD_Y_SYUREC.SYUKO_SYUSI, vbUnicode))       '出庫収支
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(OLD_Y_SYUREC.SYUKA_YMD, vbUnicode))           '出荷日付
        Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(OLD_Y_SYUREC.ODER_NO, vbUnicode))               'ｵｰﾀﾞｰ番号
        Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(OLD_Y_SYUREC.ITEM_NO, vbUnicode))               'ｱｲﾃﾑ番号
        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(OLD_Y_SYUREC.MUKE_NAME, vbUnicode))           '得意先名称
        
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(OLD_Y_SYUREC.CYU_KBN, vbUnicode))               '注文区分
        Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(OLD_Y_SYUREC.CYU_KBN_NAME, vbUnicode))     '注文区分名称
        Call UniCode_Conv(Y_SYUREC.EXPORT_KBN, StrConv(OLD_Y_SYUREC.EXPORT_KBN, vbUnicode))         '輸出出荷検査区分
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_KBN, StrConv(OLD_Y_SYUREC.LABEL_ISSUE_KBN, vbUnicode))   '個装ラベル発行区分
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_UNIT, StrConv(OLD_Y_SYUREC.LABEL_ISSUE_UNIT, vbUnicode)) '個装ラベル発行単位数
        Call UniCode_Conv(Y_SYUREC.LABEL_TANKA_KBN, StrConv(OLD_Y_SYUREC.LABEL_TANKA_KBN, vbUnicode))   '個装ラベル単価表示区分
        Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(OLD_Y_SYUREC.TANKA, vbUnicode))                   '単価
        Call UniCode_Conv(Y_SYUREC.KINGAKU, StrConv(OLD_Y_SYUREC.KINGAKU, vbUnicode))               '金額
        
        Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(OLD_Y_SYUREC.BIKOU2, vbUnicode))                 '備考２
        Call UniCode_Conv(Y_SYUREC.REBATE_KBN, StrConv(OLD_Y_SYUREC.REBATE_KBN, vbUnicode))         'ﾘﾍﾞｰﾄ区分
        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(OLD_Y_SYUREC.CHOHA_KBN, vbUnicode))           '帳端区分
        Call UniCode_Conv(Y_SYUREC.ATAISA_KBN, StrConv(OLD_Y_SYUREC.ATAISA_KBN, vbUnicode))         '値差区分
        Call UniCode_Conv(Y_SYUREC.REP_KISHU, StrConv(OLD_Y_SYUREC.REP_KISHU, vbUnicode))           '代表機種
        Call UniCode_Conv(Y_SYUREC.NS_KANRI_NO, StrConv(OLD_Y_SYUREC.NS_KANRI_NO, vbUnicode))       'NS管理区分
        Call UniCode_Conv(Y_SYUREC.MTS_HIN_CODE, StrConv(OLD_Y_SYUREC.MTS_HIN_CODE, vbUnicode))     'MTS部品ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(OLD_Y_SYUREC.BIKOU1, vbUnicode))                 '備考1
        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(OLD_Y_SYUREC.CHOKU_KBN, vbUnicode))           '直送区分
        Call UniCode_Conv(Y_SYUREC.REBATE_RATE, StrConv(OLD_Y_SYUREC.REBATE_RATE, vbUnicode))       'ﾘﾍﾞｰﾄ率
        Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(OLD_Y_SYUREC.HIN_NAME, vbUnicode))             '品名
        Call UniCode_Conv(Y_SYUREC.JGYOBA_GAI, StrConv(OLD_Y_SYUREC.JGYOBA_GAI, vbUnicode))         '対外事業場
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(OLD_Y_SYUREC.KISHU_CODE, vbUnicode))         '機種ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(OLD_Y_SYUREC.SS_CODE, vbUnicode))               '直送先ｺｰﾄﾞ
        Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(OLD_Y_SYUREC.HIN_NAI, vbUnicode))               '品番(内部)
        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(OLD_Y_SYUREC.HTANABAN, vbUnicode))             'ﾎｽﾄ棚番
        
        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, StrConv(OLD_Y_SYUREC.PRINT_YMD, vbUnicode))           '出庫表印刷日付
        Call UniCode_Conv(Y_SYUREC.KAN_YMD, StrConv(OLD_Y_SYUREC.KAN_YMD, vbUnicode))               '完了日付
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, StrConv(OLD_Y_SYUREC.KENPIN_YMD, vbUnicode))         '検品日付
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, StrConv(OLD_Y_SYUREC.TOK_KBN, vbUnicode))               '特売り区分
        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, StrConv(OLD_Y_SYUREC.JITU_SURYO, vbUnicode))         '出庫実績数量
        Call UniCode_Conv(Y_SYUREC.INS_NOW, StrConv(OLD_Y_SYUREC.INS_NOW, vbUnicode))               '取り込み日時
        
        
        
        Call UniCode_Conv(Y_SYUREC.FILLER, "")
        
        
        Do
            sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "出荷予定データ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(2).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If


    '---------------------------------------------------------  過日分出荷予定データの処理
DEL_SYU_CONV:


    MsgLab(1) = "過日分出荷予定データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(Count, "#0")
                                        
                                '(旧)過日分出荷予定データＯＰＥＮ
    If OLD_DEL_SYU_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo IDO_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_DEL_SYU_POS, OLD_DEL_SYUREC, Len(OLD_DEL_SYUREC), K0_OLD_DEL_SYU, Len(K0_OLD_DEL_SYU), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)過日分出荷予定データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(3).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        
        
        
        Call UniCode_Conv(DEL_SYUREC.WEL_ID, "")                                                      '使用端末ID
        Call UniCode_Conv(DEL_SYUREC.PRG_ID, "")                                                      '使用中ﾌﾟﾛｸﾞﾗﾑID
        
        Call UniCode_Conv(DEL_SYUREC.KAN_KBN, StrConv(OLD_DEL_SYUREC.KAN_KBN, vbUnicode))               '完了区分
        Call UniCode_Conv(DEL_SYUREC.DT_SYU, StrConv(OLD_DEL_SYUREC.DT_SYU, vbUnicode))                 'ﾃﾞｰﾀ種別
        Call UniCode_Conv(DEL_SYUREC.JGYOBU, StrConv(OLD_DEL_SYUREC.JGYOBU, vbUnicode))                 '事業部
        
        Call UniCode_Conv(DEL_SYUREC.KEY_CYU_KBN, StrConv(OLD_DEL_SYUREC.KEY_CYU_KBN, vbUnicode))       '注文区分
        Call UniCode_Conv(DEL_SYUREC.KEY_ID_NO, StrConv(OLD_DEL_SYUREC.KEY_ID_NO, vbUnicode))           'ID-NO
        
        Call UniCode_Conv(DEL_SYUREC.NAIGAI, StrConv(OLD_DEL_SYUREC.NAIGAI, vbUnicode))                 '国内外
        Call UniCode_Conv(DEL_SYUREC.KEY_HIN_NO, StrConv(OLD_DEL_SYUREC.KEY_HIN_NO, vbUnicode))         '品目番号
        
        Call UniCode_Conv(DEL_SYUREC.KEY_MUKE_CODE, StrConv(OLD_DEL_SYUREC.KEY_MUKE_CODE, vbUnicode))   '得意先ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.KEY_SS_CODE, StrConv(OLD_DEL_SYUREC.KEY_SS_CODE, vbUnicode))       '直送先ｺｰﾄﾞ
        
        Call UniCode_Conv(DEL_SYUREC.KEY_SYUKA_YMD, StrConv(OLD_DEL_SYUREC.KEY_SYUKA_YMD, vbUnicode))   '出荷日付
        
        Call UniCode_Conv(DEL_SYUREC.JGYOBA, StrConv(OLD_DEL_SYUREC.JGYOBA, vbUnicode))                 '事業場
        Call UniCode_Conv(DEL_SYUREC.DATA_KBN, StrConv(OLD_DEL_SYUREC.DATA_KBN, vbUnicode))             'ﾃﾞｰﾀ区分
        Call UniCode_Conv(DEL_SYUREC.TORI_KBN, StrConv(OLD_DEL_SYUREC.TORI_KBN, vbUnicode))             '取引区分
        Call UniCode_Conv(DEL_SYUREC.ID_NO, StrConv(OLD_DEL_SYUREC.ID_NO, vbUnicode))                   'ID-NO
        Call UniCode_Conv(DEL_SYUREC.HIN_NO, StrConv(OLD_DEL_SYUREC.HIN_NO, vbUnicode))                 '品目番号
        
        Call UniCode_Conv(DEL_SYUREC.DEN_NO, StrConv(OLD_DEL_SYUREC.DEN_NO, vbUnicode))                 '伝票番号
        Call UniCode_Conv(DEL_SYUREC.SURYO, StrConv(OLD_DEL_SYUREC.SURYO, vbUnicode))                   '数量
        Call UniCode_Conv(DEL_SYUREC.MUKE_CODE, StrConv(OLD_DEL_SYUREC.MUKE_CODE, vbUnicode))           '得意先ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.SYUKO_SYUSI, StrConv(OLD_DEL_SYUREC.SYUKO_SYUSI, vbUnicode))       '出庫収支
        Call UniCode_Conv(DEL_SYUREC.SYUKA_YMD, StrConv(OLD_DEL_SYUREC.SYUKA_YMD, vbUnicode))           '出荷日付
        Call UniCode_Conv(DEL_SYUREC.ODER_NO, StrConv(OLD_DEL_SYUREC.ODER_NO, vbUnicode))               'ｵｰﾀﾞｰ番号
        Call UniCode_Conv(DEL_SYUREC.ITEM_NO, StrConv(OLD_DEL_SYUREC.ITEM_NO, vbUnicode))               'ｱｲﾃﾑ番号
        Call UniCode_Conv(DEL_SYUREC.MUKE_NAME, StrConv(OLD_DEL_SYUREC.MUKE_NAME, vbUnicode))           '得意先名称
        
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN, StrConv(OLD_DEL_SYUREC.CYU_KBN, vbUnicode))               '注文区分
        Call UniCode_Conv(DEL_SYUREC.CYU_KBN_NAME, StrConv(OLD_DEL_SYUREC.CYU_KBN_NAME, vbUnicode))     '注文区分名称
        Call UniCode_Conv(DEL_SYUREC.EXPORT_KBN, StrConv(OLD_DEL_SYUREC.EXPORT_KBN, vbUnicode))         '輸出出荷検査区分
        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_KBN, StrConv(OLD_DEL_SYUREC.LABEL_ISSUE_KBN, vbUnicode))   '個装ラベル発行区分
        Call UniCode_Conv(DEL_SYUREC.LABEL_ISSUE_UNIT, StrConv(OLD_DEL_SYUREC.LABEL_ISSUE_UNIT, vbUnicode)) '個装ラベル発行単位数
        Call UniCode_Conv(DEL_SYUREC.LABEL_TANKA_KBN, StrConv(OLD_DEL_SYUREC.LABEL_TANKA_KBN, vbUnicode))   '個装ラベル単価表示区分
        Call UniCode_Conv(DEL_SYUREC.TANKA, StrConv(OLD_DEL_SYUREC.TANKA, vbUnicode))                   '単価
        Call UniCode_Conv(DEL_SYUREC.KINGAKU, StrConv(OLD_DEL_SYUREC.KINGAKU, vbUnicode))               '金額
        
        Call UniCode_Conv(DEL_SYUREC.BIKOU2, StrConv(OLD_DEL_SYUREC.BIKOU2, vbUnicode))                 '備考２
        Call UniCode_Conv(DEL_SYUREC.REBATE_KBN, StrConv(OLD_DEL_SYUREC.REBATE_KBN, vbUnicode))         'ﾘﾍﾞｰﾄ区分
        Call UniCode_Conv(DEL_SYUREC.CHOHA_KBN, StrConv(OLD_DEL_SYUREC.CHOHA_KBN, vbUnicode))           '帳端区分
        Call UniCode_Conv(DEL_SYUREC.ATAISA_KBN, StrConv(OLD_DEL_SYUREC.ATAISA_KBN, vbUnicode))         '値差区分
        Call UniCode_Conv(DEL_SYUREC.REP_KISHU, StrConv(OLD_DEL_SYUREC.REP_KISHU, vbUnicode))           '代表機種
        Call UniCode_Conv(DEL_SYUREC.NS_KANRI_NO, StrConv(OLD_DEL_SYUREC.NS_KANRI_NO, vbUnicode))       'NS管理区分
        Call UniCode_Conv(DEL_SYUREC.MTS_HIN_CODE, StrConv(OLD_DEL_SYUREC.MTS_HIN_CODE, vbUnicode))     'MTS部品ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.BIKOU1, StrConv(OLD_DEL_SYUREC.BIKOU1, vbUnicode))                 '備考1
        Call UniCode_Conv(DEL_SYUREC.CHOKU_KBN, StrConv(OLD_DEL_SYUREC.CHOKU_KBN, vbUnicode))           '直送区分
        Call UniCode_Conv(DEL_SYUREC.REBATE_RATE, StrConv(OLD_DEL_SYUREC.REBATE_RATE, vbUnicode))       'ﾘﾍﾞｰﾄ率
        Call UniCode_Conv(DEL_SYUREC.HIN_NAME, StrConv(OLD_DEL_SYUREC.HIN_NAME, vbUnicode))             '品名
        Call UniCode_Conv(DEL_SYUREC.JGYOBA_GAI, StrConv(OLD_DEL_SYUREC.JGYOBA_GAI, vbUnicode))         '対外事業場
        Call UniCode_Conv(DEL_SYUREC.KISHU_CODE, StrConv(OLD_DEL_SYUREC.KISHU_CODE, vbUnicode))         '機種ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.SS_CODE, StrConv(OLD_DEL_SYUREC.SS_CODE, vbUnicode))               '直送先ｺｰﾄﾞ
        Call UniCode_Conv(DEL_SYUREC.HIN_NAI, StrConv(OLD_DEL_SYUREC.HIN_NAI, vbUnicode))               '品番(内部)
        Call UniCode_Conv(DEL_SYUREC.HTANABAN, StrConv(OLD_DEL_SYUREC.HTANABAN, vbUnicode))             'ﾎｽﾄ棚番
        
        Call UniCode_Conv(DEL_SYUREC.PRINT_YMD, StrConv(OLD_DEL_SYUREC.PRINT_YMD, vbUnicode))           '出庫表印刷日付
        Call UniCode_Conv(DEL_SYUREC.KAN_YMD, StrConv(OLD_DEL_SYUREC.KAN_YMD, vbUnicode))               '完了日付
        Call UniCode_Conv(DEL_SYUREC.KENPIN_YMD, StrConv(OLD_DEL_SYUREC.KENPIN_YMD, vbUnicode))         '検品日付
        Call UniCode_Conv(DEL_SYUREC.TOK_KBN, StrConv(OLD_DEL_SYUREC.TOK_KBN, vbUnicode))               '特売り区分
        Call UniCode_Conv(DEL_SYUREC.JITU_SURYO, StrConv(OLD_DEL_SYUREC.JITU_SURYO, vbUnicode))         '出庫実績数量
        Call UniCode_Conv(DEL_SYUREC.INS_NOW, StrConv(OLD_DEL_SYUREC.INS_NOW, vbUnicode))               '取り込み日時
        
        
        
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
                    Call File_Error(sts, BtOpInsert, "過日分出荷予定データ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(3).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If



    '---------------------------------------------------------  在庫移動歴データの処理
IDO_CONV:


    MsgLab(1) = "在庫移動歴データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(4).Caption = Format(Count, "#0")
                                '(旧)在庫移動歴データＯＰＥＮ
    If OLD_IDO_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo J_NYU_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)在庫移動歴データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(4).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        
        
        
        
        
        Call UniCode_Conv(IDOREC.JITU_DT, StrConv(OLD_IDOREC.JITU_DT, vbUnicode))           '実績日付
        Call UniCode_Conv(IDOREC.JITU_TM, StrConv(OLD_IDOREC.JITU_TM, vbUnicode))           '実績時刻
        Call UniCode_Conv(IDOREC.JGYOBU, StrConv(OLD_IDOREC.JGYOBU, vbUnicode))             '事業部区分
        Call UniCode_Conv(IDOREC.NAIGAI, StrConv(OLD_IDOREC.NAIGAI, vbUnicode))             '国内外
        Call UniCode_Conv(IDOREC.HIN_GAI, StrConv(OLD_IDOREC.HIN_GAI, vbUnicode))           '品番(外部)
        Call UniCode_Conv(IDOREC.RIRK_ID, StrConv(OLD_IDOREC.RIRK_ID, vbUnicode))           '履歴種別
        Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, StrConv(OLD_IDOREC.SUMI_JITU_QTY, vbUnicode))   '実績数量(商品化済み)
        Call UniCode_Conv(IDOREC.MI_JITU_QTY, StrConv(OLD_IDOREC.MI_JITU_QTY, vbUnicode))   '実績数量(未商品)
        Call UniCode_Conv(IDOREC.FROM_SOKO, StrConv(OLD_IDOREC.FROM_SOKO, vbUnicode))       'From 倉庫№
        Call UniCode_Conv(IDOREC.FROM_RETU, StrConv(OLD_IDOREC.FROM_RETU, vbUnicode))       'From 列
        Call UniCode_Conv(IDOREC.FROM_REN, StrConv(OLD_IDOREC.FROM_REN, vbUnicode))         'From 連
        Call UniCode_Conv(IDOREC.FROM_DAN, StrConv(OLD_IDOREC.FROM_DAN, vbUnicode))         'From 段
        
        Call UniCode_Conv(IDOREC.TO_SOKO, StrConv(OLD_IDOREC.TO_SOKO, vbUnicode))           'To 倉庫№
        Call UniCode_Conv(IDOREC.TO_RETU, StrConv(OLD_IDOREC.TO_RETU, vbUnicode))           'To 列
        Call UniCode_Conv(IDOREC.TO_REN, StrConv(OLD_IDOREC.TO_REN, vbUnicode))             'To 連
        Call UniCode_Conv(IDOREC.TO_DAN, StrConv(OLD_IDOREC.TO_DAN, vbUnicode))             'To 段
        
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(OLD_IDOREC.DEN_DT, vbUnicode))             '伝票日付
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(OLD_IDOREC.DEN_NO, vbUnicode))             '伝票№
        Call UniCode_Conv(IDOREC.PRG_ID, StrConv(OLD_IDOREC.PRG_ID, vbUnicode))             '出力元ﾌﾟﾛｸﾞﾗﾑ
        
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(OLD_IDOREC.HIN_NAI, vbUnicode))           '品番(内部)
        
        Call UniCode_Conv(IDOREC.NYUKA_DT, StrConv(OLD_IDOREC.NYUKA_DT, vbUnicode))         '入荷日付
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(OLD_IDOREC.NYUKO_DT, vbUnicode))         '入庫日付
        
        Call UniCode_Conv(IDOREC.WEL_ID, StrConv(OLD_IDOREC.WEL_ID, vbUnicode))             '対象端末№
        
        Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(OLD_IDOREC.RIRK_NAME, vbUnicode))       '履歴種別名称
        
        Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(OLD_IDOREC.HIN_NAME, vbUnicode))         '品名
        
        Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode))          '品目別在庫数（商品化済み）
        Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))            '品目別在庫数（未商品）
        
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.SUMI_FROM_TANA_Zaiko_Qty, vbUnicode))    'FROM棚別品目別在庫数
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.SUMI_TO_TANA_Zaiko_Qty, vbUnicode))      'TO棚別品目別在庫数
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.MI_FROM_TANA_Zaiko_Qty, vbUnicode))      'FROM棚別品目別在庫数
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, _
                                StrConv(OLD_IDOREC.MI_TO_TANA_Zaiko_Qty, vbUnicode))        'TO棚別品目別在庫数
        
        
        Call UniCode_Conv(IDOREC.TOKU_MARK, StrConv(OLD_IDOREC.TOKU_MARK, vbUnicode))       '特売りﾏｰｸ
        Call UniCode_Conv(IDOREC.MEMO, StrConv(OLD_IDOREC.MEMO, vbUnicode))                 'メモ
        Call UniCode_Conv(IDOREC.TANTO_CODE, StrConv(OLD_IDOREC.TANTO_CODE, vbUnicode))     '担当者ｺｰﾄﾞ
        Call UniCode_Conv(IDOREC.TANTO_NAME, StrConv(OLD_IDOREC.TANTO_NAME, vbUnicode))     '担当者名称
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(OLD_IDOREC.MUKE_CODE, vbUnicode))       '得意先ｺｰﾄﾞ
        Call UniCode_Conv(IDOREC.MUKE_NAME, StrConv(OLD_IDOREC.MUKE_NAME, vbUnicode))       '得意先名称
        Call UniCode_Conv(IDOREC.SS_CODE, StrConv(OLD_IDOREC.SS_CODE, vbUnicode))           '直送先ｺｰﾄﾞ
        Call UniCode_Conv(IDOREC.SS_NAME, StrConv(OLD_IDOREC.SS_NAME, vbUnicode))           '直送先名称
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(OLD_IDOREC.MUKE_DNAME, vbUnicode))     '得意先略称
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, StrConv(OLD_IDOREC.MUKE_CHG_CD, vbUnicode))   '向け先読替えｺｰﾄﾞ
        Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(OLD_IDOREC.SUM_KBN, vbUnicode))           '集計区分
        Call UniCode_Conv(IDOREC.ID_NO, StrConv(OLD_IDOREC.ID_NO, vbUnicode))               'ID-NO
        Call UniCode_Conv(IDOREC.Ins_DateTime, StrConv(OLD_IDOREC.Ins_DateTime, vbUnicode)) '挿入日時
        
        Call UniCode_Conv(IDOREC.SHIIRE_CODE, "")                                           '仕入先ｺｰﾄﾞ
        Call UniCode_Conv(IDOREC.SHIIRE_TANKA, "")                                          '仕入単価
        Call UniCode_Conv(IDOREC.KEIJYO_YM, "")                                             '計上年月
        
        
        Call UniCode_Conv(IDOREC.FILLER, "")
        
        
        
        
        
        
        Do
            sts = BTRV(BtOpInsert, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<IDO.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "在庫移動歴データ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(4).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  入荷チェックデータの処理
J_NYU_CONV:


    MsgLab(1) = "入荷チェックデータコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(5).Caption = Format(Count, "#0")
                                        
                                '(旧)入荷チェックデータＯＰＥＮ
    If OLD_J_NYU_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo SUMZ_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_J_NYU_POS, OLD_J_NYUREC, Len(OLD_J_NYUREC), K0_OLD_J_NYU, Len(K0_OLD_J_NYU), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)入荷チェックデータ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(5).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(J_NYUREC.JGYOBU, StrConv(OLD_J_NYUREC.JGYOBU, vbUnicode))         '事業部
        Call UniCode_Conv(J_NYUREC.NAIGAI, StrConv(OLD_J_NYUREC.NAIGAI, vbUnicode))         '国内外
        Call UniCode_Conv(J_NYUREC.HIN_GAI, StrConv(OLD_J_NYUREC.HIN_GAI, vbUnicode))       '品番(外部)
        Call UniCode_Conv(J_NYUREC.JITU_QTY, StrConv(OLD_J_NYUREC.JITU_QTY, vbUnicode))     '実績数量
        Call UniCode_Conv(J_NYUREC.INS_DATE, Format(Now, "YYYYMMDD"))                       '登録日
        Call UniCode_Conv(J_NYUREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<J_NYU.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "入荷ﾁｪｯｸﾃﾞｰﾀ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(5).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  在庫集計データの処理
SUMZ_CONV:


    MsgLab(1) = "在庫集計データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(6).Caption = Format(Count, "#0")
                                        
    
                                '(旧)在庫集計データＯＰＥＮ
    If OLD_SUMZ_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo STOCK_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
    
    
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_SUMZ_POS, OLD_SUMZREC, Len(OLD_SUMZREC), K0_OLD_SUMZ, Len(K0_OLD_SUMZ), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)在庫集計データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(6).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(SUMZREC.JGYOBU, StrConv(OLD_SUMZREC.JGYOBU, vbUnicode))       '事業部
        Call UniCode_Conv(SUMZREC.NAIGAI, StrConv(OLD_SUMZREC.NAIGAI, vbUnicode))       '国内外
        Call UniCode_Conv(SUMZREC.HIN_GAI, StrConv(OLD_SUMZREC.HIN_GAI, vbUnicode))     '品番(外部)
        
        Call UniCode_Conv(SUMZREC.ST_SOKO, StrConv(OLD_SUMZREC.ST_SOKO, vbUnicode))     '標準棚番 倉庫№
        Call UniCode_Conv(SUMZREC.ST_RETU, StrConv(OLD_SUMZREC.ST_RETU, vbUnicode))     '標準棚番 列
        Call UniCode_Conv(SUMZREC.ST_REN, StrConv(OLD_SUMZREC.ST_REN, vbUnicode))       '標準棚番 連
        Call UniCode_Conv(SUMZREC.ST_DAN, StrConv(OLD_SUMZREC.ST_DAN, vbUnicode))       '標準棚番 段
        
        Call UniCode_Conv(SUMZREC.T_Zai_Qty, StrConv(OLD_SUMZREC.T_Zai_Qty, vbUnicode))     '在庫総数(当日)
        Call UniCode_Conv(SUMZREC.ZEN_Zai_Qty, StrConv(OLD_SUMZREC.ZEN_Zai_Qty, vbUnicode)) '在庫総数(前日)
                        
        Call UniCode_Conv(SUMZREC.SYK_E_QTY, StrConv(OLD_SUMZREC.SYK_E_QTY, vbUnicode))     '出庫済数
        Call UniCode_Conv(SUMZREC.NYUKA_YQTY, StrConv(OLD_SUMZREC.NYUKA_YQTY, vbUnicode))   '入荷予定数
                        
        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, StrConv(OLD_SUMZREC.HS_ZAIQTY, vbUnicode))         'ﾎｽﾄ在庫数(当日)
        Call UniCode_Conv(SUMZREC.ZEN_HS_ZAIQTY, StrConv(OLD_SUMZREC.ZEN_HS_ZAIQTY, vbUnicode)) 'ﾎｽﾄ在庫数(前日)
                        
        
        Call UniCode_Conv(SUMZREC.SAI_QTY, StrConv(OLD_SUMZREC.SAI_QTY, vbUnicode))     '差異数
        Call UniCode_Conv(SUMZREC.SUM_DT, StrConv(OLD_SUMZREC.SUM_DT, vbUnicode))       '集計日付
        
        
        
        
        Call UniCode_Conv(SUMZREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<SUMZAI.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "在庫集計ﾃﾞｰﾀ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(6).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  棚卸しデータの処理

STOCK_CONV:

    MsgLab(1) = "棚卸しデータコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(7).Caption = Format(Count, "#0")
                                        
    
                                '(旧)棚卸しデータＯＰＥＮ
    If OLD_STOCK_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                GoTo KEPPINLOG_CONV
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
    
    
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_STOCK_POS, OLD_STOCKREC, Len(OLD_STOCKREC), K0_OLD_STOCK, Len(K0_OLD_STOCK), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)棚卸しデータ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(7).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(STOCKREC.JGYOBU, StrConv(OLD_STOCKREC.JGYOBU, vbUnicode))     '事業部
        Call UniCode_Conv(STOCKREC.NAIGAI, StrConv(OLD_STOCKREC.NAIGAI, vbUnicode))     '国内外
        Call UniCode_Conv(STOCKREC.HIN_GAI, StrConv(OLD_STOCKREC.HIN_GAI, vbUnicode))   '品番(外部)
        
        Call UniCode_Conv(STOCKREC.ST_LOCATION, StrConv(OLD_STOCKREC.ST_LOCATION, vbUnicode))   '標準入庫倉庫
        Call UniCode_Conv(STOCKREC.HOST_ZAIKO, StrConv(OLD_STOCKREC.HOST_ZAIKO, vbUnicode))     'ﾎｽﾄ理論在庫
        Call UniCode_Conv(STOCKREC.POS_ZAIKO, StrConv(OLD_STOCKREC.POS_ZAIKO, vbUnicode))       'POS在庫
        
        Call UniCode_Conv(STOCKREC.ST_ZAIKO, StrConv(OLD_STOCKREC.ST_ZAIKO, vbUnicode))         '標準棚番在庫
        
        Call UniCode_Conv(STOCKREC.EE1_LOCATION, StrConv(OLD_STOCKREC.EE1_LOCATION, vbUnicode)) '別置き棚番1
        Call UniCode_Conv(STOCKREC.EE1_ZAIKO, StrConv(OLD_STOCKREC.EE1_ZAIKO, vbUnicode))       '別置き棚番1 在庫
        Call UniCode_Conv(STOCKREC.EE2_LOCATION, StrConv(OLD_STOCKREC.EE2_LOCATION, vbUnicode)) '別置き棚番2
        Call UniCode_Conv(STOCKREC.EE2_ZAIKO, StrConv(OLD_STOCKREC.EE2_ZAIKO, vbUnicode))       '別置き棚番2 在庫
        Call UniCode_Conv(STOCKREC.EE3_LOCATION, StrConv(OLD_STOCKREC.EE3_LOCATION, vbUnicode)) '別置き棚番3
        Call UniCode_Conv(STOCKREC.EE3_ZAIKO, StrConv(OLD_STOCKREC.EE3_ZAIKO, vbUnicode))       '別置き棚番3 在庫
        
        Call UniCode_Conv(STOCKREC.ETC_ZAIKO, StrConv(OLD_STOCKREC.ETC_ZAIKO, vbUnicode))       'その他在庫
        
        Call UniCode_Conv(STOCKREC.CHECK_MARK, StrConv(OLD_STOCKREC.CHECK_MARK, vbUnicode))     '照合ﾏｰｸ
        
        Call UniCode_Conv(STOCKREC.PRINT_YMD, StrConv(OLD_STOCKREC.PRINT_YMD, vbUnicode))       '印刷日付
        Call UniCode_Conv(STOCKREC.INPUT_YMD, StrConv(OLD_STOCKREC.INPUT_YMD, vbUnicode))       '入力日付
        
        Call UniCode_Conv(STOCKREC.SAI_QTY, StrConv(OLD_STOCKREC.SAI_QTY, vbUnicode))           '差異数
        
        
        
        
        
        
        
        Call UniCode_Conv(STOCKREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<STOCK.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "棚卸しﾃﾞｰﾀ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(7).Caption = Format(Count, "#0")
    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If

    '---------------------------------------------------------  欠品防止支援ログデータの処理

KEPPINLOG_CONV:
    
    MsgLab(1) = "欠品防止支援ログデータコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(8).Caption = Format(Count, "#0")
                                        
                                '(旧)欠品防止支援ログデータＯＰＥＮ
    If OLD_KEPPINLOG_Open(BtOpenNomal) Then
        If sts = BtErrFileNotFound Then
            If Mode = 0 Then
                MsgBox "コンバート終了"
                Update_Proc = False
            Else
                MsgBox "対象データなし"
                Unload Me
            End If
        End If
        Unload Me
    End If
                                        
                                        
    com = BtOpGetFirst
                                        
                                        
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_KEPPINLOG_POS, OLD_KEPPINLOGREC, Len(OLD_KEPPINLOGREC), K0_OLD_KEPPINLOG, Len(K0_OLD_KEPPINLOG), 0)
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "(旧)欠品防止支援ログデータ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(8).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(KEPPINLOGREC.JGYOBU, StrConv(OLD_KEPPINLOGREC.JGYOBU, vbUnicode))         '事業部
        Call UniCode_Conv(KEPPINLOGREC.NAIGAI, StrConv(OLD_KEPPINLOGREC.NAIGAI, vbUnicode))         '国内外
        Call UniCode_Conv(KEPPINLOGREC.HIN_GAI, StrConv(OLD_KEPPINLOGREC.HIN_GAI, vbUnicode))       '品番(外部)
        
        Call UniCode_Conv(KEPPINLOGREC.CREATE_DT, StrConv(OLD_KEPPINLOGREC.CREATE_DT, vbUnicode))   '作成日付
        
        
        Call UniCode_Conv(KEPPINLOGREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<KEPPINLOG.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "欠品防止ﾛｸﾞﾃﾞｰﾀ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

'---------------------------------------------  終了
    Cnt(8).Caption = Format(Count, "#0")

    If Mode <> 0 Then
        Me.MousePointer = vbDefault
        MsgBox "コンバート終了"
        Update_Proc = False
        Exit Function
    End If

    Me.MousePointer = vbDefault
    MsgBox "コンバート終了"
    Update_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)


Dim ans As Integer
                                
    If Index = 10 Then
        Unload Me
    End If
                                '処理選択
    Beep
    ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If ans = vbYes Then
        
        
        If Index = 9 Then
        
        
            If Update_Proc(0) Then
                Unload Me
            End If
        Else
        
        
            If Update_Proc(Index + 1) Then
                Unload Me
            End If
        End If
    End If
'    Unload Me



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
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '在庫データＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '出荷予定データＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '過日分出荷データＯＰＥＮ
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
                                '在庫移動歴データＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '入荷チェックデータＯＰＥＮ
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '在庫集計データＯＰＥＮ
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '棚卸しデータＯＰＥＮ
    If STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                                '欠品防止支援ログデータＯＰＥＮ
    If KEPPINLOG_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '(旧)品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)品目マスタ")
        End If
    End If
    
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '(旧)在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫データ")
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
    sts = BTRV(BtOpClose, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫データ")
        End If
    End If
                                            '過日分出荷データＣＬＯＳＥ
    sts = BTRV(BtOpClose, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "過日分出荷データ")
        End If
    End If
                                            '(旧)過日分出荷データＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_DEL_SYU_POS, OLD_DEL_SYUREC, Len(OLD_DEL_SYUREC), K0_DEL_SYU, Len(K0_DEL_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)過日分出荷データ")
        End If
    End If
    
    
                                            '在庫移動歴データＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴データ")
        End If
    End If
                                            '(旧)在庫移動歴データＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫移動歴データ")
        End If
    End If
    
    
                                            '入荷チェックデータＣＬＯＳＥ
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷チェックデータ")
        End If
    End If
                                            '(旧)入荷チェックデータＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_J_NYU_POS, OLD_J_NYUREC, Len(OLD_J_NYUREC), K0_OLD_J_NYU, Len(K0_OLD_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)入荷チェックデータ")
        End If
    End If
    
                                            '在庫集計データＣＬＯＳＥ
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫集計データ")
        End If
    End If
                                            '(旧)在庫集計データＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_SUMZ_POS, OLD_SUMZREC, Len(OLD_SUMZREC), K0_OLD_SUMZ, Len(K0_OLD_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫集計データ")
        End If
    End If
    
    
                                            '棚卸しデータＣＬＯＳＥ
    sts = BTRV(BtOpClose, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "棚卸しデータ")
        End If
    End If
                                            '(旧)棚卸しデータＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_STOCK_POS, OLD_STOCKREC, Len(OLD_STOCKREC), K0_OLD_STOCK, Len(K0_OLD_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)棚卸しデータ")
        End If
    End If
    
    
                                            '欠品防止支援ログデータＣＬＯＳＥ
    sts = BTRV(BtOpClose, KEPPINLOG_POS, KEPPINLOGREC, Len(KEPPINLOGREC), K0_KEPPINLOG, Len(K0_KEPPINLOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "欠品防止支援ログデータ")
        End If
    End If
                                            '(旧)欠品防止支援ログデータＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_KEPPINLOG_POS, OLD_KEPPINLOGREC, Len(OLD_KEPPINLOGREC), K0_OLD_KEPPINLOG, Len(K0_OLD_KEPPINLOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)欠品防止支援ログデータ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000501 = Nothing

    End
End Sub

