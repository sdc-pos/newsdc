VERSION 5.00
Begin VB.Form F1200501 
   BackColor       =   &H80000005&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "月平均出荷数"
   ClientHeight    =   7230
   ClientLeft      =   2310
   ClientTop       =   2610
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10095
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2544
      TabIndex        =   1
      Top             =   2160
      Width           =   204
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "月平均出荷数集計処理中"
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
      Width           =   5280
   End
End
Attribute VB_Name = "F1200501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''2011.07.01Dim Tuki            As Integer

''2011.07.01Private Type Syuka_Tbl_Tag
''2011.07.01    YM              As String * 6
''2011.07.01End Type

''2011.007.01Dim SYUKA_tbl()     As Syuka_Tbl_Tag

''Private Start_YMD   As String * 8
''Private End_YMD     As String * 8

'---------------------  2011.07.01
Private Start_YMD1      As String * 8
Private End_YMD1        As String * 8
Private TUKI1           As Integer


Private Start_YMD2      As String * 8
Private End_YMD2        As String * 8
Private TUKI2           As Integer

Private Start_YMD3      As String * 8
Private End_YMD3        As String * 8
Private TUKI3           As Integer

Private Start_2Year_YMD As String * 8
Private End_2Year_YMD   As String * 8
Private TUKI_2Year      As Integer

'---------------------  2011.07.01



Private NOT_MTS     As String

Private YOIN_TBL    As Variant

'2008.10.31
Private SHIZAI_YOIN_TBL     As Variant
'2008.10.31
Private SHIZAI_YOIN_F       As Boolean

Private CYU_KBN_TBL     As Variant
Private CYU_KBN_F       As Boolean



Private HAIKI_CODE  As String

'Private Const Last_Update_Day$ = "月平均出荷数集計処理([F120050] 2012.02.20 09:00) 生産計画項目追加"
Private Const Last_Update_Day$ = "月平均出荷数集計処理([F120050] 2019.12.02 17:00) 生産計画項目が空白でも動作する様に修正"


Private Function Update_Proc() As Integer
'-----------------------------------------------------------
'                   集計処理
'-----------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim Upd_com         As Integer

Dim ans             As Integer
Dim i               As Integer

Dim SKIP_FLG        As Boolean
Dim ING_YM          As String

Dim Max_START_YMD   As String * 8
Dim Max_END_YMD     As String * 8


    Update_Proc = True

    Label1(0).Caption = "「初期化中」"
    
    
    com = BtOpGetFirst
    
    Do
        DoEvents
        Do
            sts = BTRV(com + BtSNoWait, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<AVE_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com, "平均出荷数量")
                    Exit Function
            End Select
        Loop
            
        If sts = BtErrEOF Then
            Exit Do
        End If
    
    
        Do
            sts = BTRV(BtOpDelete, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("他端末でデータ使用中です。<AVE_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpDelete, "平均出荷数量")
                    Exit Function
            End Select
        Loop
    
        com = BtOpGetNext
    
    Loop
'---------------------------------------------------------  資材分▽  2011.07.01
    If SHIZAI_YOIN_F Then
        
        Max_START_YMD = "zzzzzzzz"
        If Start_YMD1 < Max_START_YMD Then
            Max_START_YMD = Start_YMD1
        End If
        
        If Start_2Year_YMD < Max_START_YMD Then
            Max_START_YMD = Start_2Year_YMD
        End If
        
        
        
        Max_END_YMD = ""
        If End_YMD1 > Max_END_YMD Then
            Max_END_YMD = End_YMD1
        End If
        If End_2Year_YMD > Max_END_YMD Then
            Max_END_YMD = End_2Year_YMD
        End If
        
        
        
        
        
        Label1(0).Caption = "「資材分処理中」"
        DoEvents
        
        Call UniCode_Conv(K0_IDO.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_IDO.JITU_DT, Max_START_YMD)
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
    
        com = BtOpGetGreater
        Do
        
        
            DoEvents
        
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            
            
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> SHIZAI Then
                        Exit Do
                    End If
                
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Max_END_YMD Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "在庫移動歴")
                    Exit Function
            End Select
                                                   
            SKIP_FLG = True
            
            For i = 0 To UBound(SHIZAI_YOIN_TBL)
                If StrConv(IDOREC.RIRK_ID, vbUnicode) = SHIZAI_YOIN_TBL(i) Then
                    SKIP_FLG = False
                    Exit For
                End If
            Next i
    
            If Not SKIP_FLG Then
                                        '月平均出荷数チェック
                Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                    Select Case sts
                        Case BtNoErr
                            Upd_com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_com = BtOpInsert
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<AVE_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "平均出荷数量")
                            Exit Function
                    End Select
            
            
                Loop
                
                If Upd_com = BtOpInsert Then
            
                    Call UniCode_Conv(AVE_SYUKAREC.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(AVE_SYUKAREC.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(AVE_SYUKAREC.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
                                              '標準棚番取り込みの為、品目ＲＥＡＤ
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(IDOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(IDOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(IDOREC.HIN_GAI, vbUnicode))
            
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                            Call UniCode_Conv(AVE_SYUKAREC.ST_LOCATION, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                            Call UniCode_Conv(AVE_SYUKAREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
                        
                            Call UniCode_Conv(AVE_SYUKAREC.NAI_BUHIN, StrConv(ITEMREC.NAI_BUHIN, vbUnicode))
                        
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(AVE_SYUKAREC.ST_LOCATION, "")
                            Call UniCode_Conv(AVE_SYUKAREC.HIN_NAME, "")
                            Call UniCode_Conv(AVE_SYUKAREC.NAI_BUHIN, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                            Exit Function
                    End Select
            
                    Call UniCode_Conv(AVE_SYUKAREC.UPDATE_YMD, Format(Now, "YYYYMMDD"))
                    Call UniCode_Conv(AVE_SYUKAREC.ZEN3_SYUKA, "00000000")
                    Call UniCode_Conv(AVE_SYUKAREC.ZEN2_SYUKA, "00000000")
                    Call UniCode_Conv(AVE_SYUKAREC.ZEN1_SYUKA, "00000000")
                    Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "000000.0")
                    Call UniCode_Conv(AVE_SYUKAREC.Two_Year_SYUKA, "00000000")
                
                
                
                    Call UniCode_Conv(AVE_SYUKAREC.ZEN3_YM, "")
                    Call UniCode_Conv(AVE_SYUKAREC.ZEN2_YM, "")
                    Call UniCode_Conv(AVE_SYUKAREC.ZEN1_YM, "")



                    Call UniCode_Conv(AVE_SYUKAREC.TOTAL_CNT, "00000000")           '総出荷件数
                    Call UniCode_Conv(AVE_SYUKAREC.TOTAL_AVE_CNT, "000000.0")       '平均総出荷件数
                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY1, "00000000")        '生産計画出荷数(1)
                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_CNT1, "00000000")        '生産計画出荷件数(1)
                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, "000000.0")    '平均生産計画出荷数(1)
                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_CNT1, "000000.0")    '平均生産計画出荷件数(1)

                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY2, "00000000")        '生産計画出荷数(2)
                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_CNT2, "00000000")        '生産計画出荷件数(2)
                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_QTY2, "000000.0")    '平均生産計画出荷数(2)
                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_CNT2, "000000.0")    '平均生産計画出荷件数(2)

                    Call UniCode_Conv(AVE_SYUKAREC.FILLER, "")


                
                
                End If
    
                    
                '月平均出荷数
                If StrConv(IDOREC.JITU_DT, vbUnicode) >= Start_YMD1 And StrConv(IDOREC.JITU_DT, vbUnicode) <= End_YMD1 Then
                    Call UniCode_Conv(AVE_SYUKAREC.ZEN1_SYUKA, _
                                        Format(Val(StrConv(AVE_SYUKAREC.ZEN1_SYUKA, vbUnicode)) + _
                                        (Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + _
                                        Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))), "00000000"))
        
        
                    Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, Format(Round(Val(StrConv(AVE_SYUKAREC.ZEN1_SYUKA, vbUnicode)) / TUKI1, 1), "000000.0"))
        
        
'                    Call UniCode_Conv(AVE_SYUKAREC.TOTAL_CNT, Format(CLng(StrConv(AVE_SYUKAREC.TOTAL_CNT, vbUnicode)) + 1, "00000000"))
'                    Call UniCode_Conv(AVE_SYUKAREC.TOTAL_AVE_CNT, Format(Round(CLng(StrConv(AVE_SYUKAREC.TOTAL_AVE_CNT, vbUnicode)) / TUKI1, 1), "000000.0"))
        
                End If
                
                
                '過去2年間実績
                Call UniCode_Conv(AVE_SYUKAREC.Two_Year_SYUKA, _
                                    Format(Val(StrConv(AVE_SYUKAREC.Two_Year_SYUKA, vbUnicode)) + _
                                    (Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + _
                                    Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))), "00000000"))
    
    
                
                
                Do
                    sts = BTRV(Upd_com, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("他端末でデータ使用中です。<AVE_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                            If ans = vbCancel Then
                                Exit Function
                            End If
            
                        Case Else
                            Call File_Error(sts, Upd_com, "月平均出荷数")
                            Exit Function
                    End Select
                Loop
            End If
            
            com = BtOpGetNext
        
        Loop
    End If
'---------------------------------------------------------  資材分△　2011.07.01
    
    
    
    
    
    
'---------------------------------------------------------  部品分▽  2011.07.01
    Label1(0).Caption = "「出荷伝票分処理中」"



    Max_START_YMD = "zzzzzzzz"
    If Start_YMD1 < Max_START_YMD Then
        Max_START_YMD = Start_YMD1
    End If
    
    If Start_YMD2 < Max_START_YMD Then
        Max_START_YMD = Start_YMD2
    End If
    
    If Start_YMD3 < Max_START_YMD Then
        Max_START_YMD = Start_YMD3
    End If
    
    If Start_2Year_YMD < Max_START_YMD Then
        Max_START_YMD = Start_2Year_YMD
    End If
        
        
        
    Max_END_YMD = ""
    If End_YMD1 > Max_END_YMD Then
        Max_END_YMD = End_YMD1
    End If
    
    If End_YMD2 > Max_END_YMD Then
        Max_END_YMD = End_YMD2
    End If
    
    If End_YMD3 > Max_END_YMD Then
        Max_END_YMD = End_YMD3
    End If
        
        
    If End_2Year_YMD > Max_START_YMD Then
        Max_END_YMD = End_2Year_YMD
    End If

    Call UniCode_Conv(K1_DEL_SYU.KEY_SYUKA_YMD, Max_START_YMD)
    com = BtOpGetGreaterEqual

    Do
        
        
        DoEvents
    
        sts = BTRV(com, DEL_SYU_POS, DEL_SYUREC, Len(DEL_SYUREC), K1_DEL_SYU, Len(K1_DEL_SYU), 1)
        
        
        Select Case sts
            Case BtNoErr
            
                If StrConv(DEL_SYUREC.KEY_SYUKA_YMD, vbUnicode) > Max_END_YMD Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "削除済み出荷予定")
                Exit Function
        End Select
If Trim(StrConv(DEL_SYUREC.HIN_NO, vbUnicode)) = "AD-3756-LZ" Then
    If CLng(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)) = "2011052" Then
Debug.Print
    End If
End If

        SKIP_FLG = False
        
        '全体から除外する条件
        If Trim(StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode)) = NOT_MTS Then
            SKIP_FLG = True
        End If

        If Trim(StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode)) = HAIKI_CODE Then
                SKIP_FLG = True
        End If


        If CYU_KBN_F Then
            For i = 0 To UBound(CYU_KBN_TBL)
            
                If StrConv(DEL_SYUREC.CYU_KBN, vbUnicode) = CYU_KBN_TBL(i) Then
                    SKIP_FLG = True
                    Exit For
                End If
            Next i
        End If
        '全体から除外する条件

        If Not SKIP_FLG Then
If Trim(StrConv(DEL_SYUREC.HIN_NO, vbUnicode)) = "AD-0BPC2K10" Then

    Debug.Print "OUT= " & StrConv(DEL_SYUREC.ID_NO, vbUnicode) & "=" & StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)
End If
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, StrConv(DEL_SYUREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, StrConv(DEL_SYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(DEL_SYUREC.HIN_NO, vbUnicode))
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                Select Case sts
                    Case BtNoErr
                        Upd_com = BtOpUpdate
                        Exit Do
                    Case BtErrKeyNotFound
                        Upd_com = BtOpInsert
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<AVE_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "平均出荷数量")
                        Exit Function
                End Select
        
        
            Loop
                
            If Upd_com = BtOpInsert Then
        
                Call UniCode_Conv(AVE_SYUKAREC.JGYOBU, StrConv(DEL_SYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(AVE_SYUKAREC.NAIGAI, StrConv(DEL_SYUREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(AVE_SYUKAREC.HIN_GAI, StrConv(DEL_SYUREC.HIN_NO, vbUnicode))
                                          '標準棚番取り込みの為、品目ＲＥＡＤ
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(DEL_SYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(DEL_SYUREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(DEL_SYUREC.HIN_NO, vbUnicode))
        
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Call UniCode_Conv(AVE_SYUKAREC.ST_LOCATION, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                                    StrConv(ITEMREC.ST_DAN, vbUnicode))
                        Call UniCode_Conv(AVE_SYUKAREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    
                        Call UniCode_Conv(AVE_SYUKAREC.NAI_BUHIN, StrConv(ITEMREC.NAI_BUHIN, vbUnicode))
                    
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(AVE_SYUKAREC.ST_LOCATION, "")
                        Call UniCode_Conv(AVE_SYUKAREC.HIN_NAME, "")
                        Call UniCode_Conv(AVE_SYUKAREC.NAI_BUHIN, "")
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
        
                Call UniCode_Conv(AVE_SYUKAREC.UPDATE_YMD, Format(Now, "YYYYMMDD"))
                Call UniCode_Conv(AVE_SYUKAREC.ZEN3_SYUKA, "00000000")
                Call UniCode_Conv(AVE_SYUKAREC.ZEN2_SYUKA, "00000000")
                Call UniCode_Conv(AVE_SYUKAREC.ZEN1_SYUKA, "00000000")
                Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "000000.0")
                Call UniCode_Conv(AVE_SYUKAREC.Two_Year_SYUKA, "00000000")
            
            
            
                Call UniCode_Conv(AVE_SYUKAREC.ZEN3_YM, "")
                Call UniCode_Conv(AVE_SYUKAREC.ZEN2_YM, "")
                Call UniCode_Conv(AVE_SYUKAREC.ZEN1_YM, "")



                Call UniCode_Conv(AVE_SYUKAREC.TOTAL_CNT, "00000000")           '総出荷件数
                Call UniCode_Conv(AVE_SYUKAREC.TOTAL_AVE_CNT, "000000.0")       '平均総出荷件数
                Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY1, "00000000")        '生産計画出荷数(1)
                Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_CNT1, "00000000")        '生産計画出荷件数(1)
                Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, "000000.0")    '平均生産計画出荷数(1)
                Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_CNT1, "000000.0")    '平均生産計画出荷件数(1)

                Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY2, "00000000")        '生産計画出荷数(2)
                Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_CNT2, "00000000")        '生産計画出荷件数(2)
                Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_QTY2, "000000.0")    '平均生産計画出荷数(2)
                Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_CNT2, "000000.0")    '平均生産計画出荷件数(2)
                
                Call UniCode_Conv(AVE_SYUKAREC.FILLER, "")


            
            
            End If
                
            '2年分集計
            Call UniCode_Conv(AVE_SYUKAREC.Two_Year_SYUKA, _
                                Format(Val(StrConv(AVE_SYUKAREC.Two_Year_SYUKA, vbUnicode)) + _
                                Val(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))
    
    
    
            '月平均出荷数
            If StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) >= Start_YMD1 And StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) <= End_YMD1 Then
                Call UniCode_Conv(AVE_SYUKAREC.ZEN1_SYUKA, _
                                    Format(Val(StrConv(AVE_SYUKAREC.ZEN1_SYUKA, vbUnicode)) + _
                                    Val(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))
    
    
                Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, Format(Round(Val(StrConv(AVE_SYUKAREC.ZEN1_SYUKA, vbUnicode)) / TUKI1, 1), "000000.0"))
    
    
                Call UniCode_Conv(AVE_SYUKAREC.TOTAL_CNT, Format(Val(StrConv(AVE_SYUKAREC.TOTAL_CNT, vbUnicode)) + 1, "00000000"))
                Call UniCode_Conv(AVE_SYUKAREC.TOTAL_AVE_CNT, Format(Round(Val(StrConv(AVE_SYUKAREC.TOTAL_CNT, vbUnicode)) / TUKI1, 1), "000000.0"))
    
            End If
            '---------------------------------  生産計画分
            Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode))
            Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(DEL_SYUREC.SS_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(MTSREC.DATA_KBN, "ZZ")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Exit Function
            End Select
    
    
If Trim(StrConv(AVE_SYUKAREC.HIN_GAI, vbUnicode)) = "ABA83-159" Then
'    Call LOG_OUT(LOG_F, StrConv(DEL_SYUREC.ID_NO, vbUnicode) & " " & StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) & " " & StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode))
Debug.Print StrConv(DEL_SYUREC.MUKE_CODE, vbUnicode)
Debug.Print StrConv(MTSREC.SYUKA_KBN, vbUnicode)
End If
    
    
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.02.18
'            If StrConv(MTSREC.SYUKA_KBN, vbUnicode) <> "ZZ" Then
            If StrConv(MTSREC.SYUKA_KBN, vbUnicode) <> "ZZ" And StrConv(DEL_SYUREC.CYU_KBN, vbUnicode) <> "E" Then
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> 2012.02.18
                '生産計画(1)
                If StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) >= Start_YMD2 And StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) <= End_YMD2 Then
                
                
                
                '2019/12/2 空白でも実行時エラーで停止しない様に修正
                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY1, _
                                        Format(GetLng(AVE_SYUKAREC.S_SYUKA_QTY1) + _
                                        GetLng(DEL_SYUREC.JITU_SURYO), "00000000"))
                                        
'                   Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY1, _         2019/12/2 コメントアウト
'                                       Format(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_QTY1, vbUnicode)) + _
'                                       CLng(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))
        
                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, Format(Round(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_QTY1, vbUnicode)) / TUKI2, 1), "000000.0"))
        
        
                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_CNT1, Format(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_CNT1, vbUnicode)) + 1, "00000000"))
                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_CNT1, Format(Round(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_CNT1, vbUnicode)) / TUKI2, 1), "000000.0"))

    
    
                End If
                '生産計画(2)
                If StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) >= Start_YMD3 And StrConv(DEL_SYUREC.SYUKA_YMD, vbUnicode) <= End_YMD3 Then
                
                '2019/12/2 空白でも実行時エラーで停止しない様に修正
                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY2, _
                                        Format(GetLng(AVE_SYUKAREC.S_SYUKA_QTY2) + _
                                        GetLng(DEL_SYUREC.JITU_SURYO), "00000000"))
                                        
'                   Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_QTY2, _         2019/12/2 コメントアウト
'                                       Format(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_QTY2, vbUnicode)) + _
'                                       CLng(StrConv(DEL_SYUREC.JITU_SURYO, vbUnicode)), "00000000"))

                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_QTY2, Format(Round(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_QTY2, vbUnicode)) / TUKI3, 1), "000000.0"))
        
        
                    Call UniCode_Conv(AVE_SYUKAREC.S_SYUKA_CNT2, Format(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_CNT2, vbUnicode)) + 1, "00000000"))
                    Call UniCode_Conv(AVE_SYUKAREC.S_AVE_SYUKA_CNT2, Format(Round(CLng(StrConv(AVE_SYUKAREC.S_SYUKA_CNT2, vbUnicode)) / TUKI3, 1), "000000.0"))
    
    
                End If
            Else
                Debug.Print
            End If
            
            

            
            Do
                sts = BTRV(Upd_com, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。<AVE_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Exit Function
                        End If
        
                    Case Else
                        Call File_Error(sts, Upd_com, "月平均出荷数")
                        Exit Function
                End Select
            Loop
            
        End If

        com = BtOpGetNext
    
    Loop
'---------------------------------------------------------  部品分△  2011.07.01
    
    
    
    
    
    
    
    If WriteIni(App.EXEName, "ZENKAI_YMD", App.EXEName, Format(Now, "YYYY/MM/DD HH:MM:SS")) Then
        Beep
        MsgBox ("INIファイルの書き込みに失敗しました。" & App.EXEName & " ZENKAI_YMD")
        Exit Function
    End If

    Update_Proc = False

End Function

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer
    
Dim wkS_DATE  As String
Dim wkE_DATE  As String
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
   
    Show
                                
    If GetIni(App.EXEName, "ZENKAI_YMD", App.EXEName, c) Then
    Else
        If Len(Trim(c)) > 7 Then
            If Left(Trim(c), 7) = Left(Format(Now, "YYYY/MM/DD"), 7) Then
                '当月集計済み
                Unload Me
            End If
        End If
    End If
                                
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
                                
                                
                                '月平均出荷数取り込み
''2011.07.01    If GetIni(App.EXEName, "TUKI", App.EXEName, c) Then
''2011.07.01        Tuki = 3
''2011.07.01    Else
''2011.07.01        Tuki = CInt(RTrim(c))       '月数の設定は１～３まで
''2011.07.01    End If
                                
                                
                                
                                
                                
'------------------------------------   2011.07.01  平均期間の獲得
    If GetIni(App.EXEName, "TUKI1", App.EXEName, c) Then
        TUKI1 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI1 = 3
        Else
            TUKI1 = Val(RTrim(c))
        End If
    End If

    wkE_DATE = Left(Format(DateAdd("m", -1, Date), "YYYY/MM/DD"), 8) & "31"
    For i = 1 To 31
        If IsDate(wkE_DATE) Then
            Exit For
        End If
        wkE_DATE = Left(wkE_DATE, 8) & Format(Val(Right(wkE_DATE, 2)) - i, "00")
    Next i
    wkS_DATE = Left(Format(DateAdd("m", -TUKI1 + 1, wkE_DATE), "YYYY/MM/DD"), 8) & "01"

    Start_YMD1 = Format(wkS_DATE, "YYYYMMDD")
    End_YMD1 = Format(wkE_DATE, "YYYYMMDD")



    If GetIni(App.EXEName, "TUKI2", App.EXEName, c) Then
        TUKI2 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI2 = 3
        Else
            TUKI2 = Val(RTrim(c))
        End If
    End If

    wkE_DATE = Left(Format(DateAdd("m", -1, Date), "YYYY/MM/DD"), 8) & "31"
    For i = 1 To 31
        If IsDate(wkE_DATE) Then
            Exit For
        End If
        wkE_DATE = Left(wkE_DATE, 8) & Format(Val(Right(wkE_DATE, 2)) - i, "00")
    Next i
    wkS_DATE = Left(Format(DateAdd("m", -TUKI2 + 1, wkE_DATE), "YYYY/MM/DD"), 8) & "01"

    Start_YMD2 = Format(wkS_DATE, "YYYYMMDD")
    End_YMD2 = Format(wkE_DATE, "YYYYMMDD")


    If GetIni(App.EXEName, "TUKI3", App.EXEName, c) Then
        TUKI3 = 12
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI3 = 12
        Else
            TUKI3 = Val(RTrim(c))
        End If
    End If

    wkE_DATE = Left(Format(DateAdd("m", -1, Date), "YYYY/MM/DD"), 8) & "31"
    For i = 1 To 31
        If IsDate(wkE_DATE) Then
            Exit For
        End If
        wkE_DATE = Left(wkE_DATE, 8) & Format(Val(Right(wkE_DATE, 2)) - i, "00")
    Next i
    wkS_DATE = Left(Format(DateAdd("m", -TUKI3 + 1, wkE_DATE), "YYYY/MM/DD"), 8) & "01"

    Start_YMD3 = Format(wkS_DATE, "YYYYMMDD")
    End_YMD3 = Format(wkE_DATE, "YYYYMMDD")





    wkE_DATE = Left(Format(DateAdd("m", -1, Date), "YYYY/MM/DD"), 8) & "31"
    For i = 1 To 31
        If IsDate(wkE_DATE) Then
            Exit For
        End If
        wkE_DATE = Left(wkE_DATE, 8) & Format(Val(Right(wkE_DATE, 2)) - i, "00")
    Next i
    wkS_DATE = Left(Format(DateAdd("m", -24 + 1, wkE_DATE), "YYYY/MM/DD"), 8) & "01"

    Start_2Year_YMD = Format(wkS_DATE, "YYYYMMDD")
    End_2Year_YMD = Format(wkE_DATE, "YYYYMMDD")
    TUKI_2Year = 24

'------------------------------------   2011.07.01
                                
                                
                                '除外MTS
    If GetIni(App.EXEName, "NOT_MTS", App.EXEName, c) Then
        NOT_MTS = "********"
    Else
        NOT_MTS = RTrim(c)
    End If
                                
    If GetIni(App.EXEName, "YOIN", App.EXEName, c) Then
        c = " "
    End If
    YOIN_TBL = Split(Trim(c), ",", -1)
                                
                                
                                
    '2008.10.31
    If GetIni(App.EXEName, "SHIZAI_YOIN", App.EXEName, c) Then
        SHIZAI_YOIN_F = False
    Else
        SHIZAI_YOIN_F = True
        SHIZAI_YOIN_TBL = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                '対象ｾﾝﾀｰ   2008.04.18
    If GetIni(App.EXEName, "MUKE_CODE", App.EXEName, c) Then
        HAIKI_CODE = "********"
    Else
        HAIKI_CODE = RTrim(c)
    End If
                                
                                
                                
                                '注文区分   2010.04.14
    If GetIni(App.EXEName, "CYU_KBN", App.EXEName, c) Then
        CYU_KBN_F = False
    Else
        CYU_KBN_F = True
        CYU_KBN_TBL = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '向け先管理マスタＯＰＥＮ   2011.07.01
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '在庫移動歴ＯＰＥＮ
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '月平均出荷数ＯＰＥＮ
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '削除済み出荷予定ＯＰＥＮ   2011.07.01
    If DEL_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '集計用テーブルの設定
''2011.07.01    ReDim SYUKA_tbl(0 To (Tuki - 1))
    
''2011.07.01    For i = 1 To Tuki
''2011.07.01        SYUKA_tbl(i - 1).YM = Left(Format(DateAdd("m", i * -1, Left(Format(Now, "YYYY/MM/DD"), 7) & "/01"), "YYYYMMDD"), 6)
''2011.07.01    Next i
                                '２年間集計用
''2011.07.01    For i = 31 To 28 Step -1
''2011.07.01        Start_YMD = SYUKA_tbl(0).YM & Format(i, "00")
''2011.07.01        If IsDate(Left(Start_YMD, 4) & "/" & Mid(Start_YMD, 5, 2) & "/" & Right(Start_YMD, 2)) Then
''2011.07.01            Exit For
''2011.07.01        End If
''2011.07.01    Next i
    
''2011.07.01    End_YMD = Left(Format(DateAdd("m", -24, Left(Start_YMD, 4) & "/" & Mid(Start_YMD, 5, 2) & "/" & Right(Start_YMD, 2)), "YYYYMMDD"), 6) & "01"

    Show
    
    Me.Caption = Last_Update_Day '2019/12/2追加
    
                    '処理選択
    If Update_Proc() Then
        Unload Me
    End If
    
    Unload Me


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
                                            '在庫移動歴ＣＬＯＳＥ
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "在庫移動歴")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
'    sts = Shell("..\exe\f110010.exe", vbNormalFocus)
'    If sts = 0 Then
'        MsgBox "[F110010]スキャナ制御の起動に失敗しました｡ "
'        Call Log_Out(LOG_F, "[F110010]スキャナ制御の起動に失敗しました｡")
'    End If
    
    
    Set F1200501 = Nothing

    End
End Sub

