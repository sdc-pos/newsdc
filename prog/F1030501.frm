VERSION 5.00
Begin VB.Form F1030501 
   BackColor       =   &H00C0C0C0&
   Caption         =   "出庫表印刷(向け先別)"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2430
   ClientWidth     =   7320
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "出庫表印刷処理"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1030501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const LMAX% = 46                    '頁内最大行数
Private Const MGN_L% = 5                    '左余白（桁数：１から）
Private Const MGN_U% = 1                    '上余白（行数：１から）

Dim Pdate As String                         '印刷開始日付（ﾍｯﾀﾞｰ用）
Dim Ptime As String                         '印刷開始時刻（ﾍｯﾀﾞｰ用）
'Dim PRT_CAN As Boolean                      '印刷途中キャンセル要求


Dim NormalFont As New StdFont               '印刷フォント
Dim Code39Font As New StdFont               '印刷フォント

Private KASO_NYUKA_SOKO As String * 2       '仮想　入荷倉庫番号
Private KASO_SYOHN_SOKO As String * 2       '仮想　商品化倉庫番号
Private KASO_NAI_SOKO As String * 2         '仮想　内職倉庫番号


Private Type Select_Tbl_tag                 '自動印刷用テーブル
    JGYOBU          As String * 1
    MUKE_CODE()     As String * 10
    CYU_KBN         As String * 1
    TITLE           As String
End Type

Dim Select_Tbl()    As Select_Tbl_tag

Dim Yuko_Day        As Integer

Dim Start_YMD       As String * 8
Dim End_YMD         As String * 8

Private Function Y_Syu_Get(com As Integer, Cnt As Integer) As Integer

Dim sts As Integer
Dim OP  As Integer
Dim ans As Integer
Dim i   As Integer

    
    If com = BtOpGetGreaterEqual Then
                                        '最初のＫＥＹセット
        Call UniCode_Conv(K6_Y_SYU.JGYOBU, Select_Tbl(Cnt).JGYOBU)
        
        If Select_Tbl(Cnt).CYU_KBN = "*" Then
            Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, "")
        Else
            Call UniCode_Conv(K6_Y_SYU.KEY_CYU_KBN, Select_Tbl(Cnt).CYU_KBN)
        End If
        Call UniCode_Conv(K6_Y_SYU.HTANABAN, "")
        Call UniCode_Conv(K6_Y_SYU.NAIGAI, "")
        Call UniCode_Conv(K6_Y_SYU.KEY_HIN_NO, "")
    End If
    
    OP = com + BtSNoWait
    
    Do
        
        Do
            sts = BTRV(OP, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Select_Tbl(Cnt).JGYOBU Or _
                        (Select_Tbl(Cnt).CYU_KBN <> "*" And _
                        StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Select_Tbl(Cnt).CYU_KBN) Then
                                                        '事業部，注文区分ブレーク
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "出荷予定ファイル")
                            Y_Syu_Get = sts
                            Exit Function
                        End If
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    
                    End If
                                                        'データ未完了＆未印刷？
                    If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = KAN_KBN_UN And _
                        Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then

                                                        '向け先全件指定ならＯＫ
                            For i = 0 To UBound(Select_Tbl(Cnt).MUKE_CODE)
                                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = Trim(Select_Tbl(Cnt).MUKE_CODE(i)) Then
                                    Exit For
                                End If
                            Next i
                            If i > UBound(Select_Tbl(Cnt).MUKE_CODE) Then
                                OP = BtOpGetNext + BtSNoWait
                                Exit Do
                            End If
                                        'データＯＫ
                            Y_Syu_Get = BtNoErr
                            Exit Function
                        
                        

                    End If

                    OP = BtOpGetNext + BtSNoWait
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    End If
                Case BtErrEOF
                    Y_Syu_Get = sts
                    Exit Function
                Case Else
                    Call File_Error(sts, OP + BtSNoWait, "出荷予定ファイル")
                    Y_Syu_Get = sts
                    Exit Function
            End Select
        Loop
    Loop
End Function

Private Function Print_Proc(Cnt As Integer) As Integer

Dim Lcnt            As Integer
Dim SAVE_SOKO_No    As String * 2
Dim PRI_HIN_GAI     As String * 13
Dim Betu_LOCATION   As String * 8

Dim com             As Integer
Dim sts             As Integer
Dim ans             As Integer
    
Dim ZAIKO_QTY       As Long
Dim TEMP_QTY        As Long

Dim RetBuf          As String
    
Dim SUMI_QTY        As Long
Dim MI_QTY        As Long
    
    Print_Proc = True

    
'    PRT_CAN = False
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time

    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
                                            '出荷予定データ読み込み
        sts = Y_Syu_Get(com, Cnt)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Exit Function
        End Select

        If Lcnt = 99 Then
            SAVE_SOKO_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
        Else
                                            '倉庫のブレーク
            If SAVE_SOKO_No <> Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) Then
                Lcnt = LMAX + 1
                SAVE_SOKO_No = Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2)
            End If
        End If

        If Lcnt > LMAX Then                 'ヘッダーコントロール
            If Head_Proc(Lcnt, Cnt) Then
                Exit Function
            End If
            PRI_HIN_GAI = ""
        End If
                                            
        If StrConv(Y_SYUREC.HIN_NO, vbUnicode) <> PRI_HIN_GAI Then
            PRI_HIN_GAI = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                                            '明細印刷
            Printer.Print Tab(MGN_L);
                                            '標準棚番
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);

            Printer.Print Tab(MGN_L + 10);
                                            '品番(外)
            Printer.Print Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13);

            Printer.Print Tab(MGN_L + 24);
                                            '標準棚　在庫数
            If Len(Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode))) = 0 Then
                SUMI_QTY = 0
                MI_QTY = 0
            Else
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    StrConv(Y_SYUREC.HTANABAN, vbUnicode)) Then
                    Exit Function
                End If
            End If
            
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;

                                            '別置棚検索
            If Tana_Kensaku(Betu_LOCATION) Then
                Print_Proc = True
                Exit Function
            End If
            
            SUMI_QTY = 0
            MI_QTY = 0
            
            If Len(Trim(Betu_LOCATION)) = 0 Then
            Else
                                            '別置棚　在庫数
                Printer.Print Tab(MGN_L + 35);
                Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                & Mid(Betu_LOCATION, 3, 2) & "-" _
                                & Mid(Betu_LOCATION, 5, 2) & "-" _
                                & Right(Betu_LOCATION, 2);
                
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        Last_JGYOBU, _
                                        StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                        StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                        Betu_LOCATION) Then
                    Exit Function
                End If
            End If
            
            Printer.Print Tab(MGN_L + 46);
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '商品化＆内職在庫数
            Printer.Print Tab(MGN_L + 55);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            TEMP_QTY = SUMI_QTY + MI_QTY
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NAI_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            ZAIKO_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '入荷倉庫在庫
            Printer.Print Tab(MGN_L + 64);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
                        
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
        End If
                                            '伝票№
        Printer.Print Tab(MGN_L + 77);
        Printer.Print Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6);
                                            '向け先ｺｰﾄﾞ
        Printer.Print Tab(MGN_L + 86);
        Printer.Print StrConv(Y_SYUREC.MUKE_CODE, vbUnicode);
                                            '向け先名称
        Printer.Print Tab(MGN_L + 95);
        Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
        Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
                Printer.Print StrConv(MTSREC.MUKE_DNAME, vbUnicode);
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                Exit Function
        End Select


        Printer.Print Tab(MGN_L + 105);
        TEMP_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)))
        RetBuf = Format(TEMP_QTY, "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;

        Printer.Print Tab(MGN_L + 115);
                                                '印刷フォント設定（Ｃｏｄｅ３９）
        Set Printer.Font = Code39Font
                            'バーコード(*伝票ID*)
        Printer.Print "*" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "*";
                                                '印刷フォント設定（通常）
        Set Printer.Font = NormalFont
        
        Printer.Print
        Printer.Print
        
        Lcnt = Lcnt + 3



                                                '印刷日付設定更新
        Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
            
        Do
        
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Print_Proc = SYS_CANCEL
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "出荷予定")
                    Print_Proc = SYS_ERR
                    Exit Function
                    
            End Select
        
        
        Loop

        com = BtOpGetNext
        
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If




    Print_Proc = False

End Function
                                    
Private Function Head_Proc(Lcnt As Integer, Cnt As Integer) As Integer
Dim i As Integer
Dim sts As Integer

    Head_Proc = True

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);               '97.10.14
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).Code Then
            Printer.Print "（" & RTrim(JGYOBU_T(i).NAME) & "）";
            Exit For
        End If
    Next i
    
'    Printer.Print Tab(MGN_L + 20); "前作業 ";
'                                        '印刷フォント設定
'    Set Printer.Font = Code39Font
'    Printer.Print "*LAST*";
'    Set Printer.Font = NormalFont
    
    Printer.Print Tab(MGN_L + 41);
    
    Printer.Print Select_Tbl(Cnt).TITLE & "出庫表";
    
    
    Printer.Print Tab(MGN_L + 91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print                                      '97.10.14

    Printer.Print Tab(MGN_L + 5);
    Printer.Print "倉庫：";
    Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2);
    Printer.Print Tab(MGN_L + 15);
    Call UniCode_Conv(K0_SOKO.Soko_No, Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            Printer.Print RTrim(StrConv(SOKOREC.SOKO_NAME, vbUnicode));
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "倉庫マスタ")
            Exit Function
    End Select
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "標準棚番";
    Printer.Print Tab(MGN_L + 10);
    Printer.Print "品番（外部）";
    Printer.Print Tab(MGN_L + 23);
    Printer.Print "標準棚在庫";
    Printer.Print Tab(MGN_L + 35);
    Printer.Print "別置棚番";
    Printer.Print Tab(MGN_L + 47);
    Printer.Print "別置在庫";
    Printer.Print Tab(MGN_L + 56);
    Printer.Print "商品化室";
    Printer.Print Tab(MGN_L + 65);
    Printer.Print "入荷倉庫";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "伝票№";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "出 荷 先";
    Printer.Print Tab(MGN_L + 105);
    Printer.Print "出荷数";
    Printer.Print

    Printer.Print

    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYUREC.NAIGAI, vbUnicode) Or _
                    StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(Y_SYUREC.HIN_NO, vbUnicode) Then
                    Exit Do
                End If
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2) Then
                                                'システム倉庫の判定
                    Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
                                Betu_LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                Exit Do
                        
                            End If
                        Case BtErrKeyNotFound
                                                '考えられないので読み飛ばし
                        Case Else
                            Call File_Error(sts, BtOpGetGreater, "倉庫マスタ")
                            Exit Function
                    End Select
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "在庫データ")
                Exit Function
        End Select
            
            
    Loop
    
    Tana_Kensaku = False

End Function

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim c           As String * 128
Dim i           As Integer
Dim j           As Integer
Dim Get_Data    As String * 10
Dim Work_Date   As String * 8
     
     
     
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

                                
                                '入荷仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NUKA_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NYUKA_SOKO = RTrim(c)
                                '商品化仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_SYOHN_SOKO = RTrim(c)
                                '内職仮想倉庫取り込み
    If GetIni(App.EXEName, "KASO_NAI_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NAI_SOKO = RTrim(c)
                                '印刷情報取込み
    i = -1
    Do
                                '実行パラメータ取込み
        If GetIni(App.EXEName, "JGYO" & Format(i + 2, "00"), "SYS", c) Then
            Beep
            MsgBox "出庫表用印刷パラメータの取込みに失敗しました。処理を中止して下さい。"
            End
        End If
        
        
        
        If Trim(c) = "END" Then
            Exit Do
        End If
        
        
        
        i = i + 1
        ReDim Preserve Select_Tbl(i)
        Select_Tbl(i).JGYOBU = Trim(c)
                                
                                '実行パラメータ取込み
        If GetIni(App.EXEName, "MUKE" & Format(i + 1, "00"), "SYS", c) Then
            Beep
            MsgBox "出庫表用印刷パラメータの取込みに失敗しました。処理を中止して下さい。"
            End
        End If
        
        
        For j = 0 To 99
            Call Data_Select(Trim(c), j + 1, 99, Get_Data)
            If Len(Trim(Get_Data)) = 0 Then
                Exit For
            End If
        
            ReDim Preserve Select_Tbl(i).MUKE_CODE(j)
    
            Select_Tbl(i).MUKE_CODE(j) = Get_Data
        
        Next j
                                
                                
        If GetIni(App.EXEName, "CYU" & Format(i + 1, "00"), "SYS", c) Then
            Beep
            MsgBox "出庫表用印刷パラメータの取込みに失敗しました。処理を中止して下さい。"
            End
        End If
                                
        Select_Tbl(i).CYU_KBN = Trim(c)
                                
        If GetIni(App.EXEName, "TITLE" & Format(i + 1, "00"), "SYS", c) Then
            Beep
            MsgBox "出庫表用印刷パラメータの取込みに失敗しました。処理を中止して下さい。"
            End
        End If
                                
        Select_Tbl(i).TITLE = Trim(c)
                                
                                
    Loop
                                
    If i = (-1) Then            '印刷指示なし
        End
    End If
                                
                                '倉庫マスタＯＰＥＮ
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
'    If ITEM_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '作業管理マスタＯＰＥＮ
    If SAGYO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '在庫データファイルＯＰＥＮ
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定ファイルＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '印刷フォント設定
    With NormalFont
        .NAME = F1030501.FontName
        .Size = 10
    End With
                                '印刷フォント設定（バーコード）
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
                                
        
    If GetIni(App.EXEName, "YUKO_DAY", "SYS", c) Then
        Yuko_Day = 0
    Else
        Yuko_Day = CInt(Trim(c))
    End If


    
    If Yuko_Day = 0 Then
        Start_YMD = Format(Now, "YYYYMMDD")
        End_YMD = Format(Now, "YYYYMMDD")

    Else

        Start_YMD = Format(DateAdd("d", Yuko_Day, Date), "YYYYMMDD")
        End_YMD = "99991231"

    End If
    

    For i = 0 To UBound(Select_Tbl)
        If Print_Proc(i) Then
            Unload Me
        End If
    Next i

    Unload Me
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
'    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "品目マスタ")
'        End If
'    End If
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '作業管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, SAGYO_POS, SAGYOREC, Len(SAGYOREC), K0_SAGYO, Len(K0_SAGYO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "作業管理マスタ")
        End If
    End If
                                            '在庫データＣＬＯＳＥ
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫データ")
        End If
    End If
                                            '出荷予定データＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データ")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030501 = Nothing

    End
End Sub

