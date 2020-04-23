VERSION 5.00
Begin VB.Form CONV20041 
   BackColor       =   &H00C0C0C0&
   Caption         =   "データコンバート処理"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
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
   ScaleWidth      =   10095
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "　出荷予定＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "在庫移動歴＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "在庫データ＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "品目マスタ＝"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
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
      Top             =   960
      Width           =   4800
   End
End
Attribute VB_Name = "CONV20041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long


Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

    Update_Proc = True

    GoTo ido_upd
'---------------------------------------------  品目マスタのコンバート
    MsgLab(1) = "品目マスタコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）品目マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(OLD_ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(OLD_ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(OLD_ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(OLD_ITEMREC.HIN_NAME, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_SET_DT, StrConv(OLD_ITEMREC.ST_SET_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_SOKO, StrConv(OLD_ITEMREC.ST_SOKO, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_RETU, StrConv(OLD_ITEMREC.ST_RETU, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_REN, StrConv(OLD_ITEMREC.ST_REN, vbUnicode))
        Call UniCode_Conv(ITEMREC.ST_DAN, StrConv(OLD_ITEMREC.ST_DAN, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_SOKO, StrConv(OLD_ITEMREC.BEF_SOKO, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_RETU, StrConv(OLD_ITEMREC.BEF_RETU, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_REN, StrConv(OLD_ITEMREC.BEF_REN, vbUnicode))
        Call UniCode_Conv(ITEMREC.BEF_DAN, StrConv(OLD_ITEMREC.BEF_DAN, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_NYU_DT, StrConv(OLD_ITEMREC.LAST_NYU_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_SYU_DT, StrConv(OLD_ITEMREC.LAST_SYU_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OLD_ITEMREC.HIN_NAI, vbUnicode))
        Call UniCode_Conv(ITEMREC.BIKOU_SOKO, StrConv(OLD_ITEMREC.BIKOU_SOKO, vbUnicode))
        Call UniCode_Conv(ITEMREC.BIKOU_TANA, StrConv(OLD_ITEMREC.BIKOU_TANA, vbUnicode))
        Call UniCode_Conv(ITEMREC.SIZAI_CD, StrConv(OLD_ITEMREC.SIZAI_CD, vbUnicode))
        Call UniCode_Conv(ITEMREC.HOJYU_P, StrConv(OLD_ITEMREC.HOJYU_P, vbUnicode))
        Call UniCode_Conv(ITEMREC.AVE_SYUKA, StrConv(OLD_ITEMREC.AVE_SYUKA, vbUnicode))
        Call UniCode_Conv(ITEMREC.SAMPLE_QTY, StrConv(OLD_ITEMREC.SAMPLE_QTY, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_INP_DT, StrConv(OLD_ITEMREC.LAST_INP_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.LOCK_F, StrConv(OLD_ITEMREC.LOCK_F, vbUnicode))
        Call UniCode_Conv(ITEMREC.WEL_ID, StrConv(OLD_ITEMREC.WEL_ID, vbUnicode))
        Call UniCode_Conv(ITEMREC.PRG_ID, StrConv(OLD_ITEMREC.PRG_ID, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_CHK_DT, StrConv(OLD_ITEMREC.LAST_CHK_DT, vbUnicode))
        Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, StrConv(OLD_ITEMREC.LAST_CHK_QTY, vbUnicode))
        Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, StrConv(OLD_ITEMREC.MOTO_JIGYOBU, vbUnicode))
        Call UniCode_Conv(ITEMREC.BIKOU, StrConv(OLD_ITEMREC.BIKOU, vbUnicode))
        Call UniCode_Conv(ITEMREC.IRI_QTY, StrConv(OLD_ITEMREC.IRI_QTY, vbUnicode))
        Call UniCode_Conv(ITEMREC.JAN_CODE, "")
        Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")
        Call UniCode_Conv(ITEMREC.GOODS_KBN, "0")
        Call UniCode_Conv(ITEMREC.PACKING_NO, "")
        Call UniCode_Conv(ITEMREC.FILLER, "")
        
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
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "品目マスタ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")
'---------------------------------------------  在庫データのコンバート
zaiko_upd:
    MsgLab(1) = "在庫データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(1).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）在庫データ")
                Exit Function
        End Select
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(1).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(ZAIKOREC.Soko_No, StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode))       '倉庫№
        Call UniCode_Conv(ZAIKOREC.Retu, StrConv(OLD_ZAIKOREC.Retu, vbUnicode))             '棚番　列
        Call UniCode_Conv(ZAIKOREC.Ren, StrConv(OLD_ZAIKOREC.Ren, vbUnicode))               '棚番　連
        Call UniCode_Conv(ZAIKOREC.Dan, StrConv(OLD_ZAIKOREC.Dan, vbUnicode))               '棚番　段
        Call UniCode_Conv(ZAIKOREC.JGYOBU, StrConv(OLD_ZAIKOREC.JGYOBU, vbUnicode))         '事業部
        Call UniCode_Conv(ZAIKOREC.NAIGAI, StrConv(OLD_ZAIKOREC.NAIGAI, vbUnicode))         '国内外
        Call UniCode_Conv(ZAIKOREC.HIN_GAI, StrConv(OLD_ZAIKOREC.HIN_GAI, vbUnicode))       '品目（外部）
        
        
        If StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode) = "92" Or _
            StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode) = "93" Or _
            StrConv(OLD_ZAIKOREC.Soko_No, vbUnicode) = "81" Then
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                                       '商品化／未商品化
        Else
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(OLD_ZAIKOREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(OLD_ZAIKOREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OLD_ZAIKOREC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts = BtNoErr Then
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = "1" Then
                    Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                                   '商品化／未商品化
                Else
                    Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                                   '商品化／未商品化
                End If
            Else
                Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                                   '商品化／未商品化
            End If
        End If
        
        Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(OLD_ZAIKOREC.NYUKA_DT, vbUnicode))     '入荷日
        Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(OLD_ZAIKOREC.NYUKO_DT, vbUnicode))     '入庫日
        Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(OLD_ZAIKOREC.HIN_NAI, vbUnicode))       '品番（内部）
        Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, StrConv(OLD_ZAIKOREC.YUKO_Z_QTY, vbUnicode)) '有効在庫数
        Call UniCode_Conv(ZAIKOREC.LOCK_F, "0")                                             '排他フラグ
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                                              '使用子機ID
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                                              '使用子機ID
        Call UniCode_Conv(ZAIKOREC.GOODS_YMD, "")                                           '商品化日付
        Call UniCode_Conv(ZAIKOREC.FILLER, Format(Now, "YYYYMMDD"))                                             ''FILLER
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
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫データ")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    
    Cnt(1).Caption = Format(Count, "#0")

    GoTo Update_End

'---------------------------------------------  在庫移動歴のコンバート
ido_upd:
    MsgLab(1) = "在庫移動歴コンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(Count, "#0")


    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）在庫移動歴")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(2).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        Call UniCode_Conv(IDOREC.JITU_DT, StrConv(OLD_IDOREC.JITU_DT, vbUnicode))           '実績日付
        Call UniCode_Conv(IDOREC.JITU_TM, StrConv(OLD_IDOREC.JITU_TM, vbUnicode))           '実績時刻
        Call UniCode_Conv(IDOREC.JGYOBU, StrConv(OLD_IDOREC.JGYOBU, vbUnicode))             '事業部区分
        Call UniCode_Conv(IDOREC.NAIGAI, StrConv(OLD_IDOREC.NAIGAI, vbUnicode))             '国内外
        Call UniCode_Conv(IDOREC.HIN_GAI, StrConv(OLD_IDOREC.HIN_GAI, vbUnicode))           '品目（外部）
        Call UniCode_Conv(IDOREC.RIRK_ID, StrConv(OLD_IDOREC.RIRK_ID, vbUnicode))           '履歴種別
        Call UniCode_Conv(IDOREC.SUMI_JITU_QTY, StrConv(OLD_IDOREC.JITU_QTY, vbUnicode))    '実績数量(商品化済み)
        Call UniCode_Conv(IDOREC.MI_JITU_QTY, "00000000")                                   '実績数量(実績数量(未商品))
        Call UniCode_Conv(IDOREC.FROM_SOKO, StrConv(OLD_IDOREC.FROM_SOKO, vbUnicode))       'From 倉庫№
        Call UniCode_Conv(IDOREC.FROM_RETU, StrConv(OLD_IDOREC.FROM_RETU, vbUnicode))       'From 列
        Call UniCode_Conv(IDOREC.FROM_REN, StrConv(OLD_IDOREC.FROM_REN, vbUnicode))         'From 連
        Call UniCode_Conv(IDOREC.FROM_DAN, StrConv(OLD_IDOREC.FROM_DAN, vbUnicode))         'From 段
        Call UniCode_Conv(IDOREC.TO_SOKO, StrConv(OLD_IDOREC.TO_SOKO, vbUnicode))         'TO 倉庫№
        Call UniCode_Conv(IDOREC.TO_RETU, StrConv(OLD_IDOREC.TO_RETU, vbUnicode))         'TO 列
        Call UniCode_Conv(IDOREC.TO_REN, StrConv(OLD_IDOREC.TO_REN, vbUnicode))           'TO 連
        Call UniCode_Conv(IDOREC.TO_DAN, StrConv(OLD_IDOREC.TO_DAN, vbUnicode))             'TO 段
        Call UniCode_Conv(IDOREC.DEN_DT, StrConv(OLD_IDOREC.DEN_DT, vbUnicode))             '伝票日付
        Call UniCode_Conv(IDOREC.DEN_NO, StrConv(OLD_IDOREC.DEN_NO, vbUnicode))             '伝票№
        Call UniCode_Conv(IDOREC.PRG_ID, StrConv(OLD_IDOREC.PRG_ID, vbUnicode))             '出力元プログラム
        Call UniCode_Conv(IDOREC.HIN_NAI, StrConv(OLD_IDOREC.HIN_NAI, vbUnicode))           '品番（内部）
        Call UniCode_Conv(IDOREC.NYUKA_DT, StrConv(OLD_IDOREC.NYUKA_DT, vbUnicode))         '入荷日付
        Call UniCode_Conv(IDOREC.NYUKO_DT, StrConv(OLD_IDOREC.NYUKO_DT, vbUnicode))         '入庫日付
        Call UniCode_Conv(IDOREC.WEL_ID, StrConv(OLD_IDOREC.WEL_ID, vbUnicode))             '対象端末№
        Call UniCode_Conv(IDOREC.RIRK_NAME, StrConv(OLD_IDOREC.RIRK_NAME, vbUnicode))       '履歴種別名称
        Call UniCode_Conv(IDOREC.HIN_NAME, StrConv(OLD_IDOREC.HIN_NAME, vbUnicode))         '品名
        Call UniCode_Conv(IDOREC.SUMI_HIN_Zaiko_Qty, _
                                    StrConv(OLD_IDOREC.HIN_Zaiko_Qty, vbUnicode))           '品目別在庫数（商品化済み）
        Call UniCode_Conv(IDOREC.MI_HIN_Zaiko_Qty, "00000000")                              '品目別在庫数（未商品）
        Call UniCode_Conv(IDOREC.SUMI_FROM_TANA_Zaiko_Qty, _
                                    StrConv(OLD_IDOREC.FROM_TANA_Zaiko_Qty, vbUnicode))     'FROM棚別品目別在庫数（商品化済み）
        Call UniCode_Conv(IDOREC.SUMI_TO_TANA_Zaiko_Qty, _
                                    StrConv(OLD_IDOREC.TO_TANA_Zaiko_Qty, vbUnicode))       'TO棚別品目別在庫数（商品化済み）
        Call UniCode_Conv(IDOREC.MI_FROM_TANA_Zaiko_Qty, "00000000")                        'FROM棚別品目別在庫数（未商品）
        Call UniCode_Conv(IDOREC.MI_TO_TANA_Zaiko_Qty, "00000000")                          'TO棚別品目別在庫数（未商品）
        Call UniCode_Conv(IDOREC.TOKU_MARK, StrConv(OLD_IDOREC.TOKU_MARK, vbUnicode))       '特売りマーク
        Call UniCode_Conv(IDOREC.MEMO, StrConv(OLD_IDOREC.MEMO, vbUnicode))                 'メモ
        Call UniCode_Conv(IDOREC.TANTO_CODE, "")                                            '担当者コード
        Call UniCode_Conv(IDOREC.TANTO_NAME, "")                                            '担当者名称
        Call UniCode_Conv(IDOREC.MUKE_CODE, StrConv(OLD_IDOREC.MUKE_CODE, vbUnicode))       '得意先コード
        Call UniCode_Conv(IDOREC.MUKE_NAME, StrConv(OLD_IDOREC.MUKE_DNAME, vbUnicode))       '得意先名称
        Call UniCode_Conv(IDOREC.SS_CODE, "")                                               '直送先コード
        Call UniCode_Conv(IDOREC.SS_NAME, "")                                               '直送先名称
        Call UniCode_Conv(IDOREC.MUKE_DNAME, StrConv(OLD_IDOREC.MUKE_DNAME, vbUnicode))     '得意先略称
        Call UniCode_Conv(IDOREC.MUKE_CHG_CD, StrConv(OLD_IDOREC.MUKE_CHG_CD, vbUnicode))   '向け先読替えコード
        Call UniCode_Conv(IDOREC.SUM_KBN, StrConv(OLD_IDOREC.SUM_KBN, vbUnicode))           '集計区分
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
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "在庫移動歴")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(2).Caption = Format(Count, "#0")
    GoTo Update_End
'---------------------------------------------  出荷予定のコンバート
syuka_upd:
    
    MsgLab(1) = "出荷予定データコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, OLD_Y_SYU_POS, OLD_Y_SYUREC, Len(OLD_Y_SYUREC), K0_OLD_Y_SYU, Len(K0_OLD_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "（旧）出荷予定データ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(3).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
            
        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")                  '使用端末ＩＤ
        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")                  '使用中プトグラムＩＤ
        If CLng(StrConv(OLD_Y_SYUREC.FIX_QTY, vbUnicode)) >= CLng(StrConv(OLD_Y_SYUREC.YOTEI_QTY, vbUnicode)) Then
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)    '完了区分＝完了
        Else
            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)     '完了区分＝未完
        End If
                                                                'データ種別
        Call UniCode_Conv(Y_SYUREC.DT_SYU, StrConv(OLD_Y_SYUREC.DT_SYU, vbUnicode))
                                                                '事業部コード
        Call UniCode_Conv(Y_SYUREC.JGYOBU, StrConv(OLD_Y_SYUREC.JGYOBU, vbUnicode))
                                                                '注文区分（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(OLD_Y_SYUREC.HS_CYU_KBN, vbUnicode))
                                                                '伝票ＩＤ（ＫＥＹ）（←伝票番号）
        Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, Format(CLng(StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode)), "00000000"))
                                                                '国内外
        Call UniCode_Conv(Y_SYUREC.NAIGAI, StrConv(OLD_Y_SYUREC.NAIGAI, vbUnicode))
                                                                '品目番号（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(OLD_Y_SYUREC.HIN_GAI, vbUnicode))
                                                                
        sts = GetIni(App.EXEName, Trim(StrConv(OLD_Y_SYUREC.MUKE_CODE, vbUnicode)), "SETUP", c)
        
        If sts Then
            MTS_CODE = ETS_MTS & StrConv(OLD_Y_SYUREC.NAIGAI, vbUnicode)
        Else
            MTS_CODE = Trim(c)
        End If
                                                                
        SS_CODE = ""
        
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095299" Then
            MTS_CODE = "20513770"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095412" Then
            MTS_CODE = "75T"
            SS_CODE = "20099826"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095414" Then
            MTS_CODE = "75T"
            SS_CODE = "20099826"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095413" Then
            MTS_CODE = "75T"
            SS_CODE = "20099826"
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095056" Then
            MTS_CODE = "20006876"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095298" Then
            MTS_CODE = "20006876"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "094626" Then
            MTS_CODE = "20513770"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095060" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "006460" Then
            MTS_CODE = "20064371"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095059" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095058" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
        
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095057" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "094627" Then
            MTS_CODE = "20513770"
            SS_CODE = ""
        End If
        
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095061" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095062" Then
            MTS_CODE = "29246"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095486" Then
            MTS_CODE = "20054433"
            SS_CODE = ""
        End If
                                                                
        If StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode) = "095487" Then
            MTS_CODE = "20054433"
            SS_CODE = ""
        End If
                                                                
                                                                
                                                                '得意先コード（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, MTS_CODE)
        Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, SS_CODE)        '直送先コード（ＫＥＹ）
                                                                '出荷日付（ＫＥＹ）
        Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(OLD_Y_SYUREC.DEN_DT, vbUnicode))
        Select Case StrConv(OLD_Y_SYUREC.JGYOBU, vbUnicode)     '事業場
            Case SOJIKI         '掃除機
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023210")
            Case DENKA          '電化調理
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023510")
            Case SUIHAN         '炊飯器
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023410")
            Case SENTAKU        '洗濯機（アイロン）
                Call UniCode_Conv(Y_SYUREC.JGYOBA, "00023100")
        End Select
        Call UniCode_Conv(Y_SYUREC.DATA_KBN, "1")               'データ区分（１：売上）
        Call UniCode_Conv(Y_SYUREC.TORI_KBN, "25")              '取引区分
                                                                '伝票ＩＤ（←伝票番号）
        Call UniCode_Conv(Y_SYUREC.ID_NO, Format(CLng(StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode)), "00000000"))
                                                                '品目番号
        Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(OLD_Y_SYUREC.HIN_GAI, vbUnicode))
                                                                '伝票番号
        Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(OLD_Y_SYUREC.DEN_NO, vbUnicode))
                                                                '出荷数量
        Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(StrConv(OLD_Y_SYUREC.YOTEI_QTY, vbUnicode)), "0000000"))
                                                                '得意先コード
        Call UniCode_Conv(Y_SYUREC.MUKE_CODE, MTS_CODE)
        Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")             '出庫収支
                                                                '出荷日付
        Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(OLD_Y_SYUREC.DEN_DT, vbUnicode))
        Call UniCode_Conv(Y_SYUREC.ODER_NO, "")                 'オーダー番号
        Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")                 'アイテム番号
                                                                '得意先名称
        Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(OLD_Y_SYUREC.SYUK_NAME, vbUnicode))
                                                                '注文区分
        Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(OLD_Y_SYUREC.HS_CYU_KBN, vbUnicode))
        Select Case StrConv(OLD_Y_SYUREC.HS_CYU_KBN, vbUnicode) '注文区分名称
            Case CYU_KBN_TUK        '月切
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_1)
            Case CYU_KBN_SPO        'スポット
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_2)
            Case CYU_KBN_HJU        '補充
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_3)
            Case CYU_KBN_TOK        '特売
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_4)
            Case CYU_KBN_KIN        '緊急
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_T)
            Case CYU_KBN_BOU        '貿易
                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_E)
        End Select
        Call UniCode_Conv(Y_SYUREC.EXPORT_KBN, "")              '輸出出荷検査区分
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_KBN, "")         '個装ラベル発行区分
        Call UniCode_Conv(Y_SYUREC.LABEL_ISSUE_UNIT, "")        '個装ラベル発行単位数
        Call UniCode_Conv(Y_SYUREC.LABEL_TANKA_KBN, "")         '個装ラベル単価表示区分
        Call UniCode_Conv(Y_SYUREC.TANKA, "0000000.00")         '単価
        Call UniCode_Conv(Y_SYUREC.TANKA, "0000000000")         '金額
        Call UniCode_Conv(Y_SYUREC.BIKOU2, "")                  '備考２
        Call UniCode_Conv(Y_SYUREC.REBATE_KBN, "")              'リベート区分
        Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")               '帳端区分
        Call UniCode_Conv(Y_SYUREC.ATAISA_KBN, "")              '値差区分
        Call UniCode_Conv(Y_SYUREC.REP_KISHU, "")               '代表機種
        Call UniCode_Conv(Y_SYUREC.NS_KANRI_NO, "")             'ＮＳ管理番号
        Call UniCode_Conv(Y_SYUREC.MTS_HIN_CODE, "")            'ＭＴＳ部品コード
        Call UniCode_Conv(Y_SYUREC.BIKOU1, "コンバートデータ")   '備考１
        Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, "")               '直送区分
        Call UniCode_Conv(Y_SYUREC.REBATE_RATE, "00.00")        'リベート率
                                                                '品名
        Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(OLD_Y_SYUREC.HIN_NAME, vbUnicode))
        Call UniCode_Conv(Y_SYUREC.JGYOBA_GAI, "")              '対外事業場
        Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")              '機種コード
        Call UniCode_Conv(Y_SYUREC.SS_CODE, SS_CODE)            '直送先コード
                                                                '品番（内部）
        Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(OLD_Y_SYUREC.HIN_NAI, vbUnicode))
                                                                'ホスト棚番
        Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(OLD_Y_SYUREC.HOST_TANA, vbUnicode))
                                                                '出庫表印刷日付
        If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "2" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "3" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "5" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "C" Or _
            StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = "D" Then
            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
        End If
                                                                '完了日付
        Call UniCode_Conv(Y_SYUREC.KAN_YMD, StrConv(OLD_Y_SYUREC.KAN_DT, vbUnicode))
                                                                '検品日付
        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, StrConv(OLD_Y_SYUREC.KENPIN_DT, vbUnicode))
                                                                '特売り区分
        Call UniCode_Conv(Y_SYUREC.TOK_KBN, StrConv(OLD_Y_SYUREC.TOK_KBN, vbUnicode))
                                                                '実績出庫数
        Call UniCode_Conv(Y_SYUREC.JITU_SURYO, Format(CLng(StrConv(OLD_Y_SYUREC.FIX_QTY, vbUnicode)), "0000000"))
                                                                        
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
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "出荷予定")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(3).Caption = Format(Count, "#0")


'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '処理選択
    Beep
    ans = MsgBox("実行しますか？", vbYesNo + vbQuestion, "確認入力")
    If ans = vbYes Then
        If Update_Proc() Then
            Unload Me
        End If
    End If
    MsgBox "終了しました。"
    Unload Me

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
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '（旧）品目マスタＯＰＥＮ
'    If OLD_ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
    
                                '在庫データＯＰＥＮ
'    If ZAIKO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '（旧）在庫データＯＰＥＮ
    
'    If OLD_ZAIKO_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '在庫移動歴データＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '（旧）在庫移動歴データＯＰＥＮ
    If OLD_IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データＯＰＥＮ
'    If Y_SYU_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '（旧）出荷予定データＯＰＥＮ
'    If OLD_Y_SYU_Open(BtOpenNomal) Then
'        Unload Me
'    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '品目マスタＣＬＯＳＥ
 '   sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
 '   If sts Then
 '       If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "品目マスタ")
 '       End If
 '   End If
                                            '(旧)品目マスタCLOSE
  '  sts = BTRV(BtOpClose, OLD_ITEM_POS, OLD_ITEMREC, Len(OLD_ITEMREC), K0_OLD_ITEM, Len(K0_OLD_ITEM), 0)
  '  If sts Then
  '      If sts <> BtErrNoOpen Then
  '          Call File_Error(sts, BtOpClose, "（旧）品目マスタ")
  '      End If
  '  End If
                                            '在庫データＣＬＯＳＥ
   ' sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
   ' If sts Then
   '     If sts <> BtErrNoOpen Then
   '         Call File_Error(sts, BtOpClose, "在庫データ")
   '     End If
   ' End If
                                            '(旧)在庫データＣＬＯＳＥ
   ' sts = BTRV(BtOpClose, OLD_ZAIKO_POS, OLD_ZAIKOREC, Len(OLD_ZAIKOREC), K0_OLD_ZAIKO, Len(K0_OLD_ZAIKO), 0)
   ' If sts Then
   '     If sts <> BtErrNoOpen Then
   '         Call File_Error(sts, BtOpClose, "(旧)在庫データ")
   '     End If
   ' End If
    
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '(旧)在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, OLD_IDO_POS, OLD_IDOREC, Len(OLD_IDOREC), K0_OLD_IDO, Len(K0_OLD_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "(旧)在庫移動歴")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    'sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    'If sts Then
    '    If sts <> BtErrNoOpen Then
    '        Call File_Error(sts, BtOpClose, "出荷予定データ")
    '    End If
    'End If
                                            '(旧)出荷予定データＣＬＯＳＥ
    'sts = BTRV(BtOpClose, OLD_Y_SYU_POS, OLD_Y_SYUREC, Len(OLD_Y_SYUREC), K0_OLD_Y_SYU, Len(K0_OLD_Y_SYU), 0)
    'If sts Then
    '    If sts <> BtErrNoOpen Then
    '        Call File_Error(sts, BtOpClose, "(旧)出荷予定データ")
    '    End If
    'End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV20041 = Nothing

    End
End Sub

