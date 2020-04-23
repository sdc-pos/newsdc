VERSION 5.00
Begin VB.Form F9000701 
   Caption         =   "      "
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7065
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Alignment       =   2  '中央揃え
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   13
      Top             =   720
      Width           =   492
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　　了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "実　　行"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "設定先倉庫"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   2640
      TabIndex        =   12
      Top             =   840
      Width           =   1332
   End
   Begin VB.Label SokoName 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4560
      TabIndex        =   11
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label lblIn_CNT 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   3240
      TabIndex        =   8
      Top             =   3000
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4320
      TabIndex        =   10
      Top             =   3000
      Width           =   372
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "在庫登録件数＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   3000
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   4320
      TabIndex        =   7
      Top             =   2520
      Width           =   372
   End
   Begin VB.Label lblIn_CNT 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "品目登録件数＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   2040
      Width           =   372
   End
   Begin VB.Label lblIn_CNT 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   3240
      TabIndex        =   3
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "取込み件数　＝"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   1692
   End
End
Attribute VB_Name = "F9000701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO As String * 2                 'ﾜｰｸｽﾃｰｼｮﾝ番号

Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer
Dim sts     As Integer

    Select Case Index
        Case 0
            
            Call UniCode_Conv(K0_SOKO.Soko_No, Text(0).Text)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    SokoName.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    SokoName.Caption = ""
                    Beep
                    MsgBox "入力した項目は、エラーです。"
                    Text(0).SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                    Unload Me
            End Select
            
            
            Beep
            ans = MsgBox("在庫移管処理を実行しますか？", vbYesNo, "確認")
            If ans = vbNo Then
                Text(0).SetFocus
                Exit Sub
            End If
            
            
            Call Data_Update_Proc
        Case 1
            Unload Me
    End Select

End Sub

Private Sub Form_Load()

Dim c As String * 128
Dim sts As Integer
Dim sBuffer As String * 255
Dim com     As String

                                
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

                                'ﾜｰｸｽﾃｰｼｮﾝ番号取り込み
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If


                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
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
                                '入荷予定ＯＰＥＮ
    If Y_NYU_Open(BtOpenNomal) Then
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
                                '品目マスタ（更新用ワーク）ＯＰＥＮ
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If

    Text(0).Text = "90"
    
    Show
    Text(0).SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '倉庫マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "倉庫マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '入荷予定ＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "入荷予定")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
                                            '品目マスタ（更新用ワーク）ＣＬＯＳＥ
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '要因マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "要因マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            'Ｂｔｒｉｅｖｅリセット
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F9000701 = Nothing
    
    
    End
End Sub


Private Sub Data_Update_Proc()
 
Dim Fno         As Integer
Dim ZaikoTemp   As String
Dim IN_CNT      As Integer
Dim sts         As Integer
Dim Zaiko_CNT   As Integer
Dim Item_CNT   As Integer



    If OutREC_Open_Proc() Then
        Unload Me
    End If
    
    IN_CNT = 0
    DataNo = 0
    
    Fno = FreeFile
    On Error Resume Next
    Open "c:\zaiko\IN_FILE.CSV" For Input As #Fno
    '在庫データ（ＣＳＶ）読み込み
    Do While EOF(Fno) = False
        DoEvents
        
        Line Input #Fno, ZaikoTemp
        ZaikoData = Split(ZaikoTemp, ",", True, vbTextCompare)
        IN_CNT = IN_CNT + 1
    
            
    
        lblIn_CNT(0) = Format(IN_CNT, "#0")
    
        If Data_Put_Proc() Then
            Unload Me
        End If
    
    Loop

    Close #OutFno
    Close #Fno
    Fno = FreeFile
    Open "c:\zaiko\shiji79.dat" For Binary As #Fno
    Zaiko_CNT = 0
    Item_CNT = 0

    Do
                                    
        DoEvents
                                    
                                    '指示データ読み込み
        Get #Fno, , OutREC
        If Left(StrConv(OutREC.TEXT_NO, vbUnicode), 1) < " " Then    'EOF ?
            Exit Do
        End If
                                        'トランザクション開始
        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
            Unload Me
        End If
        
        
        If Upd_Item(Item_CNT) Then
            GoTo Abort_Tran
        End If
                                        
        If NyukaY_Put(Zaiko_CNT) Then
            GoTo Abort_Tran
        End If
                                        
                                        

                                        'トランザクション終了
        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpEndTransaction, "")
            GoTo Abort_Tran
        End If
    Loop


    Close #Fno

    MsgBox "在庫移管処理が終了しました。"
    Unload Me

Abort_Tran:
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Unload Me

End Sub
                                            '品目マスタ更新
Private Function Upd_Item(IN_CNT As Integer) As Boolean
Dim sts As Integer
Dim ans As Integer
Dim Command As Integer
Dim Work As String


    Upd_Item = True


    Call UniCode_Conv(K0_ITEM.JGYOBU, "7")
    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(OutREC.HIN_GAI, vbUnicode))
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Command = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                Command = BtOpInsert
                Call UniCode_Conv(ITEMREC.JGYOBU, "7")
                Call UniCode_Conv(ITEMREC.NAIGAI, "1")
                Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(OutREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
                Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OutREC.HIN_NAI, vbUnicode))
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")
                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")
                
                Call UniCode_Conv(ITEMREC.LOCK_F, "0")          '排他フラグ
                Call UniCode_Conv(ITEMREC.WEL_ID, "")           '使用中子機ＩＤ
                Call UniCode_Conv(ITEMREC.PRG_ID, "")           '使用中プログラム
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "0000000")
                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")
                Call UniCode_Conv(ITEMREC.BIKOU, "")
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")
                
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")
                Call UniCode_Conv(ITEMREC.RANK, "")
                
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
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
    
    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(ITEMREC.LAST_INP_DT, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(OutREC.HIN_NAI, vbUnicode))
    Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(OutREC.HIN_NAME, vbUnicode))
    
    IN_CNT = IN_CNT + 1
    lblIn_CNT(1) = Format(IN_CNT, "#0")
    
    Do
        sts = BTRV(Command, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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

                                            '入荷予定作成 ＆ 入荷更新
Private Function NyukaY_Put(IN_CNT As Integer) As Boolean

Dim sts     As Integer
Dim Work    As String * 8
Dim ans     As Integer

    NyukaY_Put = True
'在庫数＝０は対象外
    If CLng(StrConv(OutREC.YOTEI_QTY, vbUnicode)) = 0 Then
            NyukaY_Put = False
            Exit Function
    End If

'入荷予定作成
                                '完了区分
'    Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_ON)
'                                'データ種別
'    Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
'                                '予定数量
'    Call UniCode_Conv(Y_NYUREC.YOTEI_QTY, Format(CLng(StrConv(OutREC.YOTEI_QTY, vbUnicode)), "00000000"))
'                                '確定数量
'    Call UniCode_Conv(Y_NYUREC.FIX_QTY, "00000000")
'                                '国内外
'    Call UniCode_Conv(Y_NYUREC.NAIGAI, "1")
'                                '事業部区分
'    Call UniCode_Conv(Y_NYUREC.JGYOBU, StrConv(OutREC.JGYOBU, vbUnicode))
'                                '直送区分
'    Call UniCode_Conv(Y_NYUREC.CYOK_KBN, StrConv(OutREC.CYOK_KBN, vbUnicode))
'                                'テキスト№
'    Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(OutREC.TEXT_NO, vbUnicode))
'                                '伝票日付
'    Call UniCode_Conv(Y_NYUREC.DEN_DT, StrConv(OutREC.DEN_DT, vbUnicode))
'                                '入出庫区分
'    Call UniCode_Conv(Y_NYUREC.IO_KBN, StrConv(OutREC.IO_KBN, vbUnicode))
'                                '赤黒区分
'    Call UniCode_Conv(Y_NYUREC.PM_KBN, StrConv(OutREC.PM_KBN, vbUnicode))
'                                '伝票種別
'    Call UniCode_Conv(Y_NYUREC.DEN_SYU, StrConv(OutREC.DEN_SYU, vbUnicode))
'                                '伝票№
'    Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(OutREC.DEN_NO, vbUnicode))
'                                '注文区分
'    Call UniCode_Conv(Y_NYUREC.CYU_KBN, StrConv(OutREC.CYU_KBN, vbUnicode))
'                                '品番（外部）
'    Call UniCode_Conv(Y_NYUREC.HIN_GAI, StrConv(OutREC.HIN_GAI, vbUnicode))
'                                '品番（内部）
'    Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(OutREC.HIN_NAI, vbUnicode))
'                                '品名
'    Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(OutREC.HIN_NAME, vbUnicode))
'                                '予算単位（元）
'    Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(OutREC.YOSAN_FROM, vbUnicode))
'                                '予算単位（先）
'    Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(OutREC.YOSAN_TO, vbUnicode))
'                                '倉庫区分（ﾎｽﾄ）
'    Call UniCode_Conv(Y_NYUREC.HOST_SOKO, StrConv(OutREC.HOST_SOKO, vbUnicode))
'                                '棚番（ﾎｽﾄ）←　標準入庫棚番（品目ﾏｽﾀ）
'    Work = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
'           StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'           StrConv(ITEMREC.ST_REN, vbUnicode) & _
'           StrConv(ITEMREC.ST_DAN, vbUnicode)
'    Call UniCode_Conv(Y_NYUREC.HOST_TANA, Work)
'                                '支給先／出荷先
'    Call UniCode_Conv(Y_NYUREC.SYUK_CODE, StrConv(OutREC.SYUK_CODE, vbUnicode))
'                                '支給先／出荷先名
'    Call UniCode_Conv(Y_NYUREC.SYUK_NAME, StrConv(OutREC.SYUK_NAME, vbUnicode))
'                                '先行入荷数
'    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
'                                '完了日付
'    Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
'                                'FILLER
'    Call UniCode_Conv(Y_NYUREC.FILLER, "")
'
'入荷予定データ追加（入荷分）
'    Do
'        sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
'        Select Case sts
'            Case BtNoErr
'                Exit Do
'            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("他端末でデータ使用中です。<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
'            Case BtErrDuplicates
'                Exit Do
'            Case Else
'                Call File_Error(sts, BtOpInsert, "入荷予定")
'                Exit Function
'        End Select
'    Loop

'入荷数で在庫データ更新（＋）
    If Nyuko_Update_Proc(StrConv(OutREC.JGYOBU, vbUnicode), _
                            "1", _
                            StrConv(OutREC.HIN_GAI, vbUnicode), _
                            StrConv(OutREC.DEN_DT, vbUnicode), _
                            (Trim(Text(0).Text) & "01" & "01" & "01"), _
                            "10", _
                            0, _
                            CLng(StrConv(OutREC.YOTEI_QTY, vbUnicode)), _
                            WS_NO, _
                            WS_NO, _
                            , _
                            "在庫移管") Then
        Exit Function
    
    End If

    IN_CNT = IN_CNT + 1
    lblIn_CNT(2) = Format(IN_CNT, "#0")
    
    NyukaY_Put = False

End Function



   
Private Sub Text_GotFocus(Index As Integer)
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer
Dim sts As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case 0
            Call UniCode_Conv(K0_SOKO.Soko_No, Text(0).Text)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    SokoName.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    SokoName.Caption = ""
                    Beep
                    MsgBox "入力した項目は、エラーです。"
                    Text(0).SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
                    Unload Me
            End Select
    End Select
        
    Command1(0).SetFocus

End Sub
