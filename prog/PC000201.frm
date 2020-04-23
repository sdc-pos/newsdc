VERSION 5.00
Begin VB.Form PC000201 
   BackColor       =   &H00C0C0C0&
   Caption         =   "構成マスタコンバート処理"
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
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5640
      Width           =   8055
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   4920
      Width           =   8055
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Cnt 
      Alignment       =   1  '右揃え
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "構成マスタ＝"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
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
Attribute VB_Name = "PC000201"
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


Dim FileNo          As Long
Dim fileName        As String


Dim COMPO_REC       As Variant
Dim RecordBuf       As String

Dim c               As String * 128

Dim i               As Integer

Dim SEQNO           As Integer

Dim Err_FLg         As Boolean



    Update_Proc = True

    FileNo = FreeFile
    
                                'ログファイル名取り込み
    If GetIni("FILE", "COMPO_TXT", "CONV2006", c) Then
        Beep
        MsgBox "[COMPO_TXT]の獲得に失敗しました。処理を中止して下さい。"
        Unload Me
    End If
    fileName = RTrim(c)
    
        
    Open fileName For Input As FileNo
    
    
    
    
    
    MsgLab(1) = "構成マスタコンバート処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
                                        
                                        
    Do Until EOF(FileNo)
        
        DoEvents
        
        Line Input #FileNo, RecordBuf
        
        
        Err_FLg = False
        
        COMPO_REC = Split(RecordBuf, vbTab, -1)
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        
        '---------------------------------------------------------- ﾍｯﾀﾞｰﾚｺｰﾄﾞ
        Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))           '仕向け先ｺｰﾄﾞ
                                                                                            
        'コードマスタ読み込み
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, CStr(COMPO_REC(1)))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                                                                                    '事業部
                Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                    '国内外
                Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
            
            Case BtErrKeyNotFound
                Err_FLg = True
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                Exit Function
        End Select
        
        
        If Not Err_FLg Then
            Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(0)))
                            
        
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound
                    Err_FLg = True
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
            End Select
        
            If Not Err_FLg Then
        
        
                Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, CStr(COMPO_REC(0)))    '品番
                                                                                            
                Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)               'データ区分（ﾍｯﾀﾞｰ）
                Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")                   '追番
                                                                                            
                Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, CStr(COMPO_REC(3))) '基本クラス
                Call UniCode_Conv(P_COMPO_O_REC.BIKOU, Trim(CStr(COMPO_REC(68))) & "/" & Trim(CStr(COMPO_REC(69))))    '備考
Text3.Text = Trim(CStr(COMPO_REC(68)))
Text4.Text = Trim(CStr(COMPO_REC(69)))
                                                                                            
                Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")                     'Filler
                
                
                Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, "CONV")              '更新担当者
                                                                                '更新日時
                Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
                Do
                    sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        
                        Case BtErrDuplicates
                            Call Log_Out(LOG_F, "DUP HEAD " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)))
                            
                            Exit Do
                        
                        Case Else
                            
                            
                            
                            Call File_Error(sts, BtOpInsert, "構成マスタ")
                            Exit Function
                    End Select
                Loop
        
            End If
        End If
            
If Trim(COMPO_REC(0)) = "AMC00P-EW09" Then
    Debug.Print
End If
            
        If Err_FLg Then
            'エラー発生
            Call Log_Out(LOG_F, "HEAD " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)))
    
        Else
        '---------------------------------------------------------- 個装資材ﾚｺｰﾄﾞ
            SEQNO = 0
    
            For i = 7 To 13
            
                If COMPO_REC(i) = "" Then
                Else
                                                                                                
                    '-----------------------------------------------    最初はｶﾚﾝﾄの事業部/国内外で
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
        
                    Err_FLg = False
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            '国内外を反転する
                            If Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)) = NAIGAI_NAI Then
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                            Else
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            End If
                        
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                Case BtErrKeyNotFound
                                
                                    '資材として
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
'                                            Err_FLg = True
                                        
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(COMPO_REC(i)))

                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Exit Function
                                    End Select
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
        
                    If Not Err_FLg Then
        
                        Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))   '仕向け先ｺｰﾄﾞ
                                                                                    '事業部
                        Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                    '国内外
                        Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                        
                        Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, CStr(COMPO_REC(0)))        '品番
                        
                        
                        Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_KOSOU)                                  'データ区分（個装資材）
                        
                        SEQNO = SEQNO + 10
                        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                        '追番
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                                    '子　種別
                        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))      '子　事業部
                        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))      '子　国内外
                        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '子　品番
                                
                        If IsNumeric(COMPO_REC(i + 7)) Then                                 '子　員数
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(COMPO_REC(i + 7)), "000.00"))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, "001.00")
                        End If
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                                       '子　備考
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "CONV")                                  '更新担当者
                                                                                                            '更新日時
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                            
                                                                                            
                                                                                            
                        Do
                            sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                
                                Case BtErrDuplicates
                                    Call Log_Out(LOG_F, "DUP 個装資材 " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                                    
                                    Exit Do
                                
                                Case Else
                                    
                                    
                                    
                                    Call File_Error(sts, BtOpInsert, "構成マスタ")
                                    Exit Function
                            End Select
                        Loop
                                                                                            
                                                                                            
                    Else
                        'エラー発生
                        Call Log_Out(LOG_F, "個装資材 " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                    End If
                                                                                    
                End If
            Next i
    
        '---------------------------------------------------------- 外装資材ﾚｺｰﾄﾞ
            SEQNO = 0
    
            For i = 21 To 23
            
                If COMPO_REC(i) = "" Then
                Else
                                                                                                
                    '-----------------------------------------------    最初はｶﾚﾝﾄの事業部/国内外で
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
        
                    Err_FLg = False
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            '国内外を反転する
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                            
                            If Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)) = NAIGAI_NAI Then
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                            Else
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            End If
                        
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                Case BtErrKeyNotFound
                                
                                    '資材として
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
                                
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
'                                            Err_FLg = True
                                        
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(COMPO_REC(i)))
                                        
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Exit Function
                                    End Select
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
                    
                    If Not Err_FLg Then
                    
                        Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))   '仕向け先ｺｰﾄﾞ
                                                                                            '事業部
                        Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                            '国内外
                        Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                        
                        Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, CStr(COMPO_REC(0)))        '品番
        
            
                            
Text1.Text = CStr(COMPO_REC(0))
Text1.Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                            
                            
                        Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_GAISOU)                                 'データ区分（外装資材）
                        SEQNO = SEQNO + 10
                        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                        '追番
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                                    '子　種別
                        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))      '子　事業部
                        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))      '子　国内外
                        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '子　品番
                                
                        If IsNumeric(COMPO_REC(i + 3)) Then                                                 '子　員数
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(COMPO_REC(i + 3)), "000.00"))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, "001.00")
                        End If
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                                       '子　備考
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "CONV")                                  '更新担当者
                                                                                                            '更新日時
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                            
                        Do
                            sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                
                                Case BtErrDuplicates
                                    
                                    Call Log_Out(LOG_F, "DUP 外装資材 " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                                    Exit Do
                                
                                Case Else
                                    
                                    
                                    
                                    Call File_Error(sts, BtOpInsert, "構成マスタ")
                                    Exit Function
                            End Select
                        Loop
                                                                                            
                    Else
                        'エラー発生
                        Call Log_Out(LOG_F, "外装資材 " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                    End If
                                                                                    
                End If
            Next i
    
        '---------------------------------------------------------- 同梱／構成部品ﾚｺｰﾄﾞ
            SEQNO = 0
    
            For i = 27 To 46
            
                If COMPO_REC(i) = "" Then
                Else
                                                                                                
                    '-----------------------------------------------    最初はｶﾚﾝﾄの事業部/国内外で
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, CStr(COMPO_REC(i)))
                    
                    Err_FLg = False
        
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                        
                            '国内外を反転する
                            If Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)) = NAIGAI_NAI Then
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                            Else
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_GAI)
                            End If
                        
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                Case BtErrKeyNotFound
                                
                                    '資材として
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
'                                            Err_FLg = True
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, CStr(COMPO_REC(i)))
                                        
                                        
                                        Case Else
                                            
                                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                            Exit Function
                                    End Select
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Function
                            End Select
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
        
                    If Not Err_FLg Then
                        
                        
                        Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, CStr(COMPO_REC(1)))   '仕向け先ｺｰﾄﾞ
                        
                                                                                            '事業部
                        Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Trim(StrConv(P_CODEREC.OPTION1, vbUnicode)))
                                                                                            '国内外
                        Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Trim(StrConv(P_CODEREC.OPTION2, vbUnicode)))
                        
                        Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, CStr(COMPO_REC(0)))        '品番
                        
                        
                        Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, P_DOUKON)                                 'データ区分（同梱／構成）
                        SEQNO = SEQNO + 10
                        Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                        '追番
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, "")                                    '子　種別
                        Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))      '子　事業部
                        Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))      '子　国内外
                        Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))    '子　品番
                                
                        If IsNumeric(COMPO_REC(i + 20)) Then                                                '子　員数
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(CDbl(COMPO_REC(i + 20)), "000.00"))
                        Else
                            Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, "001.00")
                        End If
                        
                        Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, "")                                       '子　備考
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, "CONV")                                  '更新担当者
                                                                                                            '更新日時
                        Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                                                                            
                        Do
                            sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("他端末でデータ使用中です。<P_COMPOREC.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                
                                Case BtErrDuplicates
                                    Call Log_Out(LOG_F, "DUP 同梱／構成 " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                                    
                                    Exit Do
                                
                                Case Else
                                    
                                    
                                    
                                    Call File_Error(sts, BtOpInsert, "構成マスタ")
                                    Exit Function
                            End Select
                        Loop
                    Else
                        'エラー発生
                        Call Log_Out(LOG_F, "同梱／構成 " & CStr(COMPO_REC(1)) & "-" & CStr(COMPO_REC(0)) & "-" & CStr(COMPO_REC(i)))
                    End If
                                                                                    
                End If
            Next i
    
    
    
        End If
    
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")


    MsgBox "ｺﾝﾊﾞｰﾄ終了！！"

    Close #FileNo



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
                                '構成マスタＯＰＥＮ
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                'コードマスタＯＰＥＮ
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '構成マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "構成マスタ")
        End If
    End If
                                            'コードマスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "コードマスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000201 = Nothing

    End
End Sub

