Attribute VB_Name = "ODR_COMN"
Option Explicit
'********************************************************************
'*
'*              ＯＤＲ用　共通変数&Ｓｕｂ
'*
'********************************************************************

Public ODR_Return       As Integer      '確認画面終了状態

Public GW_SHIMEBI       As String       '繰越日付

Public GW_TOUGETU       As String       '締め日から得た当月（yyyymm)

Public GW_MAX_YYMM      As String       '当月（yyyymm)からの最大使用月


Public GW_PC_NM As String               '実行端末名


Public GW_SIMUKE        As String       '仕向け先
Public GW_JIGYOBU       As String       '事業部
Public GW_NAIGAI        As String       '国内外
Public GW_TANTO         As String       '担当者

Public GW_USE_YM        As String       '使用月 yyyymm

Public GW_HINGAI        As String       '対象品番

Public GW_HINGAI_KO     As String       '品番　（子品番）
Public GW_JIGYOBU_KO    As String       '事業部（子品番）
Public GW_NAIGAI_KO     As String       '国内外（子品番）


Public Type SE_JGYOBU_TBL

    SHIMUKE    As String * 2
    JGYOBU          As String * 1
    NAIGAI          As String * 1

End Type


Public SE_JGYOBU_T()       As SE_JGYOBU_TBL

Function SET_JGYOBU_T()
'
'           仕向け先をﾃｰﾌﾞﾙにｾｯﾄ
'           P_CODEファイルは、呼び基でOpen/Close
'

Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer
Dim i           As Integer
    
    SET_JGYOBU_T = True
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    i = -1
    Do
        DoEvents
        
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SE_JGYOBU_T(0 To i)
                SE_JGYOBU_T(i).SHIMUKE = Trim(StrConv(P_CODEREC.C_Code, vbUnicode))
                SE_JGYOBU_T(i).JGYOBU = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
                SE_JGYOBU_T(i).NAIGAI = Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
                        
                        
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "コードマスタ")
                'Unload Me
                Exit Function
        End Select
    
        com = BtOpGetNext
    
    Loop
    

    SET_JGYOBU_T = False

End Function
