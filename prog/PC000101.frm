VERSION 5.00
Begin VB.Form PC000101 
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
      Height          =   495
      Index           =   3
      Left            =   5400
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "品目マスタ(商品)＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   6
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "品目マスタ(ﾗﾍﾞﾙ)＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "品目マスタ(資材)＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
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
Attribute VB_Name = "PC000101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type KAISHA_Tbl_Tag
    C_Code          As String * 2
    C_NAME          As String
    JGYOBU          As String * 1
    NAIGAI          As String * 1
End Type


Private KAISHA_Tbl()    As KAISHA_Tbl_Tag



Private Function Update_Proc(SHORI_MODE) As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long


Dim FileNo          As Long
Dim fileName        As String


Dim ITEM_REC        As Variant
Dim RecordBuf       As String
Dim wk              As String

Dim c               As String * 128

Dim i               As Integer
Dim j               As Integer

Dim Err_Flg         As Integer

    Update_Proc = True


    Select Case SHORI_MODE
        Case 0

        FileNo = FreeFile
        
                                    'ログファイル名取り込み
        If GetIni("FILE", "SHIZAI_TXT", "CONV2006", c) Then
            Beep
            MsgBox "[SHIZAI_TXT]の獲得に失敗しました。処理を中止して下さい。"
            Unload Me
        End If
        fileName = RTrim(c)
        
            
        Open fileName For Input As FileNo
        
        
        
    '-----------------------------------------------------------------------------  業務管理－－＞ＰＯＳ
        
        
        MsgLab(1) = "資材　製品マスタコンバート処理中！！"
        Me.MousePointer = vbHourglass
        Count = 0
        DISP_INTERVAL = 0
        Cnt(0).Caption = Format(Count, "#0")
                                            
                                            
                                            
                                            
        Do Until EOF(FileNo)
            
            DoEvents
            
            Line Input #FileNo, RecordBuf
            
            ITEM_REC = Split(RecordBuf, vbTab, -1)
            
            
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(0).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
            
            
            For i = 0 To UBound(ITEM_REC)
                For j = 0 To Len(ITEM_REC(i))
    
                    If Mid(CStr(ITEM_REC(i)), j + 1, 1) = """" Then
                        Mid(ITEM_REC(i), j + 1, 1) = " "
                    End If
    
                Next j
    
                ITEM_REC(i) = Trim(ITEM_REC(i))
            Next i
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)               '事業部(=資材)
            Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)           '国内外(=国内)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(ITEM_REC(1)))   '資材品番
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                    Upd_com = BtOpUpdate
                
                
                Case BtErrKeyNotFound
                    
                    Upd_com = BtOpInsert
                
                Case Else
                    
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
            
                    
            If Upd_com = BtOpInsert Then
                    
                Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)       '事業部=資材
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)   '国内外
                Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(ITEM_REC(1)))       '品目ｺｰﾄﾞ
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                    '標準棚番設定日付
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")                      '標準入庫　倉庫
                Call UniCode_Conv(ITEMREC.ST_RETU, "")                      '標準入庫　列
                Call UniCode_Conv(ITEMREC.ST_REN, "")                       '標準入庫　連
                Call UniCode_Conv(ITEMREC.ST_DAN, "")                       '標準入庫　段
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                     '前回入庫　倉庫
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")                     '前回入庫　列
                Call UniCode_Conv(ITEMREC.BEF_REN, "")                      '前回入庫　連
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")                      '前回入庫　段
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                  '最終入庫日
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                  '最終出庫日
                Call UniCode_Conv(ITEMREC.HIN_NAI, "")                      '品番（内）
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                   'ﾎｽﾄ倉庫
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                   'ﾎｽﾄ棚番
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")              '補充点
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")            '月平均出荷数
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  'ｻﾝﾌﾟﾙ数
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  'ｻﾝﾌﾟﾙ数
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                  '最終入荷日付
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                  '最終照合日付
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")         '照合時在庫数
                Call UniCode_Conv(ITEMREC.BIKOU, "")                        '印刷備考
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")                      '印刷入り数
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")                     'JANｺｰﾄﾞ
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                   '品番読み替えｺｰﾄﾞ
                Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                   '商品化有無
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")                   '個装箱№
                Call UniCode_Conv(ITEMREC.RANK, "")                         '現在ﾗﾝｸ
                Call UniCode_Conv(ITEMREC.NEW_RANK, "")                     '新ﾗﾝｸ
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                  'ｸﾞﾘｯｸｽ棚番1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                  'ｸﾞﾘｯｸｽ棚番2
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                  'ｸﾞﾘｯｸｽ棚番3
            
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                 '品名E
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                      '備考
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                '会社名
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                     '機種(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                     '機種(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                     '機種(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                      '紙
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                    'ﾌﾟﾗｽﾁｯｸ
                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                    '価格(1)
                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                    '価格(2)
                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                    '価格(3)
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                      '適用機種ﾗﾍﾞﾙ
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                      'ﾗﾍﾞﾙ枚数
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                '適用機種備考
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                '作業指示
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                     '備考(3)
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                '事業部名
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                    '入り数
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                      '棚番(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                      '棚番(2)
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                      '収単／担当者
                
                Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_OFF)            '在庫管理対象
                
                Call UniCode_Conv(ITEMREC.FILLER, "")                       'Filler
        
            End If
            
            If IsNumeric(ITEM_REC(33)) Then                                 '危険在庫
                Call UniCode_Conv(ITEMREC.HOJYU_P, Format(CDbl(ITEM_REC(33)), "00000000000"))
            Else
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000000")
            End If
            
            
            Call UniCode_Conv(ITEMREC.HIN_NAME, CStr(ITEM_REC(2)))          '品名
            Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, CStr(ITEM_REC(3)))      '仕入区分
            Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, CStr(ITEM_REC(4)))      '販売区分
            Call UniCode_Conv(ITEMREC.G_SYUSHI, CStr(ITEM_REC(5)))          '収支単位
            If CStr(ITEM_REC(6)) = "" Then                                  '組立製品
                Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_OFF)
            Else
                Call UniCode_Conv(ITEMREC.G_KUMITATE, P_ASSEMBLY_ON)
            End If
            If IsNumeric(CStr(ITEM_REC(7))) Then
                                                                            '売価単価
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, Format(CDbl(ITEM_REC(7)), "00000000.00"))
                                                                            '売価設定日
                
                If IsDate(ITEM_REC(8)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(CStr(ITEM_REC(8)), "YYYYMMDD"))
                Else
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, Format(Now, "YYYYMMDD"))
                End If
            Else
                                                                            '売価単価
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
                                                                            '売価設定日
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")
            End If
            
            If IsNumeric(ITEM_REC(9)) Then
                                                                            '原価
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, Format(CDbl(ITEM_REC(9)), "00000000.00"))
                                                                            '原価設定日
                If IsDate(ITEM_REC(10)) Then
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(CStr(ITEM_REC(10)), "YYYYMMDD"))
                Else
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, Format(Now, "YYYYMMDD"))
                End If
            Else
                                                                            '原価
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")
                                                                            '原価設定日
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")
            End If
            
            
            j = -1
            For i = 13 To 19 Step 3
            
                j = j + 1
                
                If j = 0 Then
                
                    If IsNumeric(ITEM_REC(34)) Then
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, CLng(ITEM_REC(34)))                  'ﾛｯﾄ数
                    Else
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                  'ﾛｯﾄ数
                    End If
                    
                    Select Case Trim(CStr(ITEM_REC(35)))
                        Case "D"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "007")            'ﾘｰﾄﾞﾀｲﾑ
                    
                        Case "F"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "010")            'ﾘｰﾄﾞﾀｲﾑ
                        Case "K"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "014")            'ﾘｰﾄﾞﾀｲﾑ
                    
                        Case "L"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "007")            'ﾘｰﾄﾞﾀｲﾑ
                    
                    
                        Case "P"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "010")            'ﾘｰﾄﾞﾀｲﾑ
                    
                        Case "Q"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "010")            'ﾘｰﾄﾞﾀｲﾑ
                    
                        Case "S"
                                        
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "000")            'ﾘｰﾄﾞﾀｲﾑ
                    
                        Case Else
                            Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")            'ﾘｰﾄﾞﾀｲﾑ
                    End Select
                
                Else
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                  'ﾛｯﾄ数
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")            'ﾘｰﾄﾞﾀｲﾑ
                End If
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, "")        '前回注文日
                Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, "")       '前回注文数
                
                
                
                If ITEM_REC(i) = "" Then
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, "")             '仕入先
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")            '単価
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")         '単価設定日
                Else
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, CStr(ITEM_REC(i)))      '仕入先
                    If IsNumeric(ITEM_REC(i + 1)) Then                                      '単価
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, _
                                                Format(CDbl(ITEM_REC(i + 1)), "00000000.00"))
                    Else
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")
                    End If
                    If IsDate(ITEM_REC(i + 2)) Then                                         '単価設定日
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, _
                                                Format((ITEM_REC(i + 2)), "YYYYMMDD"))
                    Else
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")
                    End If
            
                End If
            
            
            
            Next i
            
            
            If IsNumeric(CStr(ITEM_REC(24))) Then                                         '前月在庫金額
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, Format(ITEM_REC(24), "00000000000"))
            Else
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")
            End If
            
            Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, CStr(ITEM_REC(35)))             '資材区分
            
            Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_OFF)                   'ﾗﾍﾞﾙ貼り付け
            
            
            
            Call UniCode_Conv(ITEMREC.UPD_TANTO, "CONV")                    '更新担当者
                                                                            '更新日時
            
            
            Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
            
            
            Do
                sts = BTRV(Upd_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
                        
                        Call File_Error(sts, Upd_com, "品目マスタ")
                        Exit Function
                End Select
            Loop
            
        
        Loop
    '---------------------------------------------  終了
    
        Cnt(0).Caption = Format(Count, "#0")
        
        Close #FileNo

    Case 1

'-----------------------------------------------------------------------------  商品ラベル－－＞ＰＯＳ


        'コードマスタより会社名／事業部名のセット
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN07_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, "")
            
        com = BtOpGetGreater
    
        i = -1
        Erase KAISHA_Tbl
    
        Do
            
            sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                
                    If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN07_CD Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    
                    Exit Do
                
                Case Else
                    
                    Call File_Error(sts, com, "コードマスタ")
                    Exit Function
            End Select
        
            i = i + 1
            ReDim Preserve KAISHA_Tbl(i)
        
            KAISHA_Tbl(i).C_Code = StrConv(P_CODEREC.C_Code, vbUnicode)
            KAISHA_Tbl(i).C_NAME = Trim(StrConv(P_CODEREC.C_NAME, vbUnicode))
            KAISHA_Tbl(i).JGYOBU = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
            KAISHA_Tbl(i).NAIGAI = Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        
        
        
        
            com = BtOpGetNext
        
        
        Loop
    
    
    
    
        FileNo = FreeFile
        
                                    'ログファイル名取り込み
        If GetIni("FILE", "LABEL_TXT", "CONV2006", c) Then
            Beep
            MsgBox "[LABEL_TXT]の獲得に失敗しました。処理を中止して下さい。"
            Unload Me
        End If
        fileName = RTrim(c)
        
            
        Open fileName For Input As FileNo
        
        
        
        
        
        MsgLab(1) = "商品ラベル　機種マスタコンバート処理中！！"
        Me.MousePointer = vbHourglass
        Count = 0
        DISP_INTERVAL = 0
        Cnt(1).Caption = Format(Count, "#0")
                                            
                                            
                                            
                                            
        Do Until EOF(FileNo)
            
            DoEvents
            
            RecordBuf = ""
            
            Do
                wk = Input(1, FileNo)
                If wk = "!" Then
                    wk = Input(2, FileNo)
                    Exit Do
                End If
                RecordBuf = RecordBuf & wk
            Loop
            
            ITEM_REC = Split(RecordBuf, vbTab, -1)
            
            
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(1).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
            
                    
            For i = 0 To UBound(ITEM_REC)
                For j = 0 To Len(ITEM_REC(i))
    
                    If Mid(CStr(ITEM_REC(i)), j + 1, 1) = """" Then
                        Mid(ITEM_REC(i), j + 1, 1) = " "
                    End If
    
                Next j
    
                ITEM_REC(i) = Trim(ITEM_REC(i))
            Next i
                    
            For i = 0 To UBound(KAISHA_Tbl)
            
            
                If Trim(KAISHA_Tbl(i).C_NAME) = Trim(CStr(ITEM_REC(8))) Then
                    Exit For
                End If
            
            Next i
            
            
            If i > UBound(KAISHA_Tbl) Then
                'エラー発生
                Call Log_Out(LOG_F, CStr(ITEM_REC(0)) & "-" & CStr(ITEM_REC(8)))
                
            Else
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, KAISHA_Tbl(i).JGYOBU)     '事業部
                Call UniCode_Conv(K0_ITEM.NAIGAI, KAISHA_Tbl(i).NAIGAI)     '国内外
                Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(ITEM_REC(0)))       '品番
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Upd_com = BtOpUpdate
                    
                    
                    Case BtErrKeyNotFound
                        
                        Upd_com = BtOpInsert
                    
                    Case Else
                        
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
            
            
                If Upd_com = BtOpInsert Then
                
                    Call UniCode_Conv(ITEMREC.JGYOBU, KAISHA_Tbl(i).JGYOBU)                     '事業部
                    Call UniCode_Conv(ITEMREC.NAIGAI, KAISHA_Tbl(i).NAIGAI)                     '国内外
                    Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(ITEM_REC(0)))                       '品目ｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.HIN_NAME, CStr(ITEM_REC(1)))                      '品目名称
                    
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                                    '標準棚番設定日付
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")                                      '標準入庫　倉庫
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")                                      '標準入庫　列
                    Call UniCode_Conv(ITEMREC.ST_REN, "")                                       '標準入庫　連
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")                                       '標準入庫　段
                    Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                                     '前回入庫　倉庫
                    Call UniCode_Conv(ITEMREC.BEF_RETU, "")                                     '前回入庫　列
                    Call UniCode_Conv(ITEMREC.BEF_REN, "")                                      '前回入庫　連
                    Call UniCode_Conv(ITEMREC.BEF_DAN, "")                                      '前回入庫　段
                    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                                  '最終入庫日
                    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                                  '最終出庫日
                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")                                      '品番（内）
                    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                                   'ﾎｽﾄ倉庫
                    Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                                   'ﾎｽﾄ棚番
                    Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")                              '補充点
                    Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")                            '月平均出荷数
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  'ｻﾝﾌﾟﾙ数
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  'ｻﾝﾌﾟﾙ数
                    Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                                  '最終入荷日付
                    Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                                  '最終照合日付
                    Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")                         '照合時在庫数
                    Call UniCode_Conv(ITEMREC.BIKOU, "")                                        '印刷備考
                    Call UniCode_Conv(ITEMREC.IRI_QTY, "")                                      '印刷入り数
                    Call UniCode_Conv(ITEMREC.JAN_CODE, "")                                     'JANｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                                   '品番読み替えｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                                   '商品化有無
                    Call UniCode_Conv(ITEMREC.PACKING_NO, "")                                   '個装箱№
                    Call UniCode_Conv(ITEMREC.RANK, "")                                         '現在ﾗﾝｸ
                    Call UniCode_Conv(ITEMREC.NEW_RANK, "")                                     '新ﾗﾝｸ
                    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番1
                    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番2
                    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番3
                
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '業務管理　 仕入区分
                    Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")                                 '           販売区分
                    Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                                     '           収支単位
                    Call UniCode_Conv(ITEMREC.G_KUMITATE, "")                                   '           組立製品
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")                                  '           標準粗利売価単価　9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")                               '           標準粗利売価設定日
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")                                  '           標準粗利原価単価  9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")                               '           標準粗利原価設定日
                    
                    For j = 0 To 2                                                              '仕入先情報
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, "")                     '           仕入先コード
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")                    '           単価
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")                 '           単価設定日
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                      '           単価設定日
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")                '           リードタイム
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, "")            '           最終発注日
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, "")           '           最終発注数
                    
                    Next j
                
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")                              '           前月在庫金額
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '           資材区分
                    Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")                                  '           ﾗﾍﾞﾙ貼り付け
                    Call UniCode_Conv(ITEMREC.S_TANTO, "")                                      '収単／担当者
                    Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)                            '在庫管理対象
                    
                    Call UniCode_Conv(ITEMREC.FILLER, "")                                       'Filler
        
                
                End If
                    
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, CStr(ITEM_REC(2)))                  '品名Ｅ
                Call UniCode_Conv(ITEMREC.L_BIKOU, CStr(ITEM_REC(5)))                       '備考
                    
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, KAISHA_Tbl(i).C_Code)              '会社名
                
                Call UniCode_Conv(ITEMREC.L_KISHU1, CStr(ITEM_REC(7)))                      '機種１
                Call UniCode_Conv(ITEMREC.L_KISHU2, CStr(ITEM_REC(9)))                      '機種２
                '--入れ替え
'                Call UniCode_Conv(ITEMREC.L_KISHU3, CStr(ITEM_REC(10)))                     '機種３
                                
'                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, CStr(ITEM_REC(18)))                '適用機種備考
                
                Call UniCode_Conv(ITEMREC.L_KISHU3, CStr(ITEM_REC(18)))                     '機種３
                                
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, CStr(ITEM_REC(10)))                '適用機種備考
                '--入れ替え
                
                
                
                If Trim(CStr(ITEM_REC(11))) = "TRUE" Then                                   '紙
                    Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_PAPER, L_PAPER_OFF)
                End If
                
                If Trim(CStr(ITEM_REC(12))) = "TRUE" Then                                   'ﾌﾟﾗｽﾁｯｸ
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, L_PLASTIC_OFF)
                End If
                If IsNumeric(ITEM_REC(13)) Then                                             '価格(1)
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(CDbl(ITEM_REC(13)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "")
                End If
                If IsNumeric(ITEM_REC(14)) Then                                             '価格(2)
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(CDbl(ITEM_REC(14)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "")
                End If
                If IsNumeric(ITEM_REC(15)) Then                                             '価格(3)
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(CDbl(ITEM_REC(15)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "")
                End If
                    
                If Trim(CStr(ITEM_REC(16))) = "TRUE" Then                                   '適用機種ﾗﾍﾞﾙ
                    Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_LABEL, L_LABEL_OFF)
                End If
                
                If Trim(CStr(ITEM_REC(17))) = "TRUE" Then                                   '枚数ﾗﾍﾞﾙ
                    Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_ON)
                Else
                    Call UniCode_Conv(ITEMREC.L_MAISU, L_MAISU_OFF)
                End If
                    
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, CStr(ITEM_REC(19)))                '作業指示
                Call UniCode_Conv(ITEMREC.L_BIKOU3, CStr(ITEM_REC(20)))                     '備考３
                    
                    
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")
                For i = 0 To UBound(KAISHA_Tbl)                                             '事業部ｺｰﾄﾞ
                
                
                    If KAISHA_Tbl(i).C_NAME = Trim(ITEM_REC(22)) Then
                        Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, KAISHA_Tbl(i).C_Code)
                        Exit For
                    End If
                
                Next i
                    
                If IsNumeric(ITEM_REC(23)) Then                                             '入り数
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, Format(CDbl(ITEM_REC(23)), "00000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")
                End If
            
            
                Call UniCode_Conv(ITEMREC.L_TANA1, CStr(ITEM_REC(24)))                      '棚番1
                Call UniCode_Conv(ITEMREC.L_TANA2, CStr(ITEM_REC(25)))                      '棚番2
                Call UniCode_Conv(ITEMREC.JAN_CODE, CStr(ITEM_REC(26)))                     'JANｺｰﾄﾞ
            
            
                Do
                    sts = BTRV(Upd_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
                            
                            Call File_Error(sts, Upd_com, "品目マスタ")
                            Exit Function
                    End Select
                Loop
            End If
        
        Loop
    '---------------------------------------------  終了
    
        Cnt(1).Caption = Format(Count, "#0")
        
        Close #FileNo
    Case 2


'-----------------------------------------------------------------------------  商品化－－＞ＰＯＳ
                                    'ログファイル名取り込み
        If GetIni("FILE", "COMPO_TXT", "CONV2006", c) Then
            Beep
            MsgBox "[COMPO_TXT]の獲得に失敗しました。処理を中止して下さい。"
            Unload Me
        End If
        FileNo = FreeFile
        
        
        fileName = RTrim(c)
        
            
        Open fileName For Input As FileNo
        
        
        
        
        MsgLab(1) = "商品化　製品マスタコンバート処理中！！"
        Me.MousePointer = vbHourglass
        Count = 0
        DISP_INTERVAL = 0
        Cnt(2).Caption = Format(Count, "#0")
                                            
                                            
                                            
                                            
        Do Until EOF(FileNo)
            
            DoEvents
            
            Line Input #FileNo, RecordBuf
            
            ITEM_REC = Split(RecordBuf, vbTab, -1)
            
            
            Count = Count + 1
            DISP_INTERVAL = DISP_INTERVAL + 1
            If DISP_INTERVAL = 100 Then
                Cnt(2).Caption = Format(Count, "#0")
                DISP_INTERVAL = 0
            End If
            
            
            For i = 0 To UBound(ITEM_REC)
                For j = 0 To Len(ITEM_REC(i))
    
                    If Mid(CStr(ITEM_REC(i)), j + 1, 1) = """" Then
                        Mid(ITEM_REC(i), j + 1, 1) = " "
                    End If
    
                Next j
    
                ITEM_REC(i) = Trim(ITEM_REC(i))
            Next i
                                                                                                
            'コードマスタ読み込み
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, CStr(ITEM_REC(1)))
            
            Err_Flg = False
            
            
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound
                    Err_Flg = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                    Exit Function
            End Select
            
            If Not Err_Flg Then
                        
            
            
            
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))      '事業部(=資材)
                Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))      '国内外(=国内)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(ITEM_REC(0)))   '品番
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    
                        Upd_com = BtOpUpdate
                    
                    
                    Case BtErrKeyNotFound
                        
                        Upd_com = BtOpInsert
                    
                    Case Else
                        
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
                
                        
                If Upd_com = BtOpInsert Then
                        
                    Call UniCode_Conv(ITEMREC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))       '事業部=資材
                    Call UniCode_Conv(ITEMREC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))   '国内外
                    Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(ITEM_REC(0)))       '品目ｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                    '標準棚番設定日付
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")                      '標準入庫　倉庫
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")                      '標準入庫　列
                    Call UniCode_Conv(ITEMREC.ST_REN, "")                       '標準入庫　連
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")                       '標準入庫　段
                    Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                     '前回入庫　倉庫
                    Call UniCode_Conv(ITEMREC.BEF_RETU, "")                     '前回入庫　列
                    Call UniCode_Conv(ITEMREC.BEF_REN, "")                      '前回入庫　連
                    Call UniCode_Conv(ITEMREC.BEF_DAN, "")                      '前回入庫　段
                    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                  '最終入庫日
                    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                  '最終出庫日
                    Call UniCode_Conv(ITEMREC.HIN_NAI, CStr(ITEM_REC(0)))       '品番（内）
                    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                   'ﾎｽﾄ倉庫
                    Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                   'ﾎｽﾄ棚番
                    Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")              '補充点
                    Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")            '月平均出荷数
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  'ｻﾝﾌﾟﾙ数
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                  'ｻﾝﾌﾟﾙ数
                    Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                  '最終入荷日付
                    Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                  '最終照合日付
                    Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")         '照合時在庫数
                    Call UniCode_Conv(ITEMREC.BIKOU, "")                        '印刷備考
                    Call UniCode_Conv(ITEMREC.IRI_QTY, "")                      '印刷入り数
                    Call UniCode_Conv(ITEMREC.JAN_CODE, "")                     'JANｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                   '品番読み替えｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, "0")                   '商品化有無
                    Call UniCode_Conv(ITEMREC.PACKING_NO, "")                   '個装箱№
                    Call UniCode_Conv(ITEMREC.RANK, "")                         '現在ﾗﾝｸ
                    Call UniCode_Conv(ITEMREC.NEW_RANK, "")                     '新ﾗﾝｸ
                    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                  'ｸﾞﾘｯｸｽ棚番1
                    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                  'ｸﾞﾘｯｸｽ棚番2
                    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                  'ｸﾞﾘｯｸｽ棚番3
                
                    Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")                 '品名E
                    Call UniCode_Conv(ITEMREC.L_BIKOU, "")                      '備考
                    Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")                '会社名
                    Call UniCode_Conv(ITEMREC.L_KISHU1, "")                     '機種(1)
                    Call UniCode_Conv(ITEMREC.L_KISHU2, "")                     '機種(2)
                    Call UniCode_Conv(ITEMREC.L_KISHU3, "")                     '機種(3)
                    Call UniCode_Conv(ITEMREC.L_PAPER, "")                      '紙
                    Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                    'ﾌﾟﾗｽﾁｯｸ
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                    '価格(1)
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                    '価格(2)
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                    '価格(3)
                    Call UniCode_Conv(ITEMREC.L_LABEL, "")                      '適用機種ﾗﾍﾞﾙ
                    Call UniCode_Conv(ITEMREC.L_MAISU, "")                      'ﾗﾍﾞﾙ枚数
                    Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")                '適用機種備考
                    Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")                '作業指示
                    Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                     '備考(3)
                    Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")                '事業部名
                    Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                    '入り数
                    Call UniCode_Conv(ITEMREC.L_TANA1, "")                      '棚番(1)
                    Call UniCode_Conv(ITEMREC.L_TANA2, "")                      '棚番(2)
                    
                    Call UniCode_Conv(ITEMREC.S_TANTO, "")                      '収単／担当者
                    
                    Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)            '在庫管理対象
                    
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")                      '品目名称
                        
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")                                    '標準棚番設定日付
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")                                      '標準入庫　倉庫
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")                                      '標準入庫　列
                    Call UniCode_Conv(ITEMREC.ST_REN, "")                                       '標準入庫　連
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")                                       '標準入庫　段
                    Call UniCode_Conv(ITEMREC.BEF_SOKO, "")                                     '前回入庫　倉庫
                    Call UniCode_Conv(ITEMREC.BEF_RETU, "")                                     '前回入庫　列
                    Call UniCode_Conv(ITEMREC.BEF_REN, "")                                      '前回入庫　連
                    Call UniCode_Conv(ITEMREC.BEF_DAN, "")                                      '前回入庫　段
                    Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")                                  '最終入庫日
                    Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")                                  '最終出庫日
                    Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")                                   'ﾎｽﾄ倉庫
                    Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")                                   'ﾎｽﾄ棚番
                    Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")                              '補充点
                    Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")                            '月平均出荷数
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  'ｻﾝﾌﾟﾙ数
                    Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "0")                                  'ｻﾝﾌﾟﾙ数
                    Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")                                  '最終入荷日付
                    Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")                                  '最終照合日付
                    Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "00000000")                         '照合時在庫数
                    Call UniCode_Conv(ITEMREC.BIKOU, "")                                        '印刷備考
                    Call UniCode_Conv(ITEMREC.IRI_QTY, "")                                      '印刷入り数
                    Call UniCode_Conv(ITEMREC.JAN_CODE, "")                                     'JANｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")                                   '品番読み替えｺｰﾄﾞ
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, "1")                                   '商品化有無
                    Call UniCode_Conv(ITEMREC.PACKING_NO, "")                                   '個装箱№
                    Call UniCode_Conv(ITEMREC.RANK, "")                                         '現在ﾗﾝｸ
                    Call UniCode_Conv(ITEMREC.NEW_RANK, "")                                     '新ﾗﾝｸ
                    Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番1
                    Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番2
                    Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")                                  'ｸﾞﾘｯｸｽ棚番3
                
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '業務管理　 仕入区分
                    Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")                                 '           販売区分
                    Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                                     '           収支単位
                    Call UniCode_Conv(ITEMREC.G_KUMITATE, "")                                   '           組立製品
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")                                  '           標準粗利売価単価　9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")                               '           標準粗利売価設定日
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")                                  '           標準粗利原価単価  9(8)V99
                    Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")                               '           標準粗利原価設定日
                    
                    For j = 0 To 2                                                              '仕入先情報
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).CODE, "")                     '           仕入先コード
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA, "")                    '           単価
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).TANKA_DT, "")                 '           単価設定日
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LOT, "")                      '           単価設定日
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LEAD_TIME, "")                '           リードタイム
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_DT, "")            '           最終発注日
                        Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(j).LAST_ORDER_QTY, "")           '           最終発注数
                    
                    Next j
                
                    Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")                              '           前月在庫金額
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")                                 '           資材区分
                    Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")                                  '           ﾗﾍﾞﾙ貼り付け
                    Call UniCode_Conv(ITEMREC.S_TANTO, "")                                      '収単／担当者
                    Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)                            '在庫管理対象
                    
                    Call UniCode_Conv(ITEMREC.FILLER, "")                                       'Filler
                    
                    
            
                
                
                    Do
                        sts = BTRV(Upd_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
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
                                
                                Call File_Error(sts, Upd_com, "品目マスタ")
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            End If
        
        Loop
    '---------------------------------------------  終了
    
        Cnt(2).Caption = Format(Count, "#0")
        
        Close #FileNo
    
    End Select

    MsgBox "コンバート終了"

End Function

Private Sub Command1_Click(index As Integer)

Dim ans     As Integer
Dim Mesg    As String

    Select Case index
    
        Case 0
            Mesg = "「資材」"
        
        Case 1
            Mesg = "「ﾗﾍﾞﾙ」"
        Case 2
            Mesg = "「商品」"
        Case 3
            Unload Me
    End Select


    ans = MsgBox(Mesg & "ｺﾝﾊﾞｰﾄ処理を実行しますか？", vbYesNo, "確認入力")
    If ans = vbYes Then
    
        If Update_Proc(index) Then
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
                                
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
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
    
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000101 = Nothing

    End
End Sub

