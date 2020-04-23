Attribute VB_Name = "MainPr00030"
Option Explicit

Public GLB_SYUSHI_F     As String



Sub Main()
    
    
    
    GLB_SYUSHI_F = Trim(Command)

'    GLB_SYUSHI_F = "01"

    PR000301.Show
End Sub



Public Function tmpP_STOCK_MAKE_Proc() As Integer
'----------------------------------------------------------------------------
'           印刷用一時ファイルを作成する
'----------------------------------------------------------------------------
Dim sts             As Integer

Dim com             As Integer
Dim Save_Jgyobu     As String
Dim Save_Naigai     As String
Dim Save_Hin_Gai    As String

Dim ZEN_ZAIKO       As Long

Dim ZEN_ZAIKO_KIN   As Long



    tmpP_STOCK_MAKE_Proc = True
    
    
    com = BtOpGetFirst
       
    Do
    
        DoEvents
    
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K1_P_STOCK, Len(K1_P_STOCK), 1)
            
If Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) = "D705" Then
    Debug.Print
End If
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材棚卸しﾃﾞｰﾀ")
                Exit Function
        End Select
    
        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" And _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" Then
            ZEN_ZAIKO = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode))
        
        
            If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)) Then
                ZEN_ZAIKO_KIN = CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode))
            Else
                ZEN_ZAIKO_KIN = 0
            End If
        Else
        
        
        
        
        
        
        
        
        
            If tmpP_STOCK_OUTPUT_Proc(ZEN_ZAIKO) Then
                Exit Function
            End If
        
        
        End If
        
        
        
        com = BtOpGetGreater
    
    
    
    Loop
    
    
    
    
    '前月残ｾｯﾄ
    com = BtOpGetFirst
    
    Do
    
        DoEvents
    
        sts = BTRV(com, P_STOCK_POS, P_STOCK_REC, Len(P_STOCK_REC), K1_P_STOCK, Len(K1_P_STOCK), 1)
            
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材棚卸しﾃﾞｰﾀ")
                Exit Function
        End Select
    
        
If Trim(StrConv(P_STOCK_REC.HIN_GAI, vbUnicode)) = "AK01" Then
    Debug.Print
End If
        
        If Trim(StrConv(P_STOCK_REC.CODE, vbUnicode)) = "" And _
            Trim(StrConv(P_STOCK_REC.TANKA, vbUnicode)) = "" Then
        
        
        
            Call UniCode_Conv(K2_tmpP_STOCK.G_SYUSHI, StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode))
            Call UniCode_Conv(K2_tmpP_STOCK.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K2_tmpP_STOCK.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K2_tmpP_STOCK.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
        
        
            Call UniCode_Conv(K2_tmpP_STOCK.INPUT_DATE, "")
            Call UniCode_Conv(K2_tmpP_STOCK.CODE, "")
            Call UniCode_Conv(K2_tmpP_STOCK.TANKA, "")
            
            
            sts = BTRV(BtOpGetGreaterEqual, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K2_tmpP_STOCK, Len(K2_tmpP_STOCK), 2)
                
            Select Case sts
                Case BtNoErr
                
                
                
                
                
                    If StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode) = StrConv(tmpP_STOCK_REC.G_SYUSHI, vbUnicode) And _
                        StrConv(P_STOCK_REC.JGYOBU, vbUnicode) = StrConv(tmpP_STOCK_REC.JGYOBU, vbUnicode) And _
                        StrConv(P_STOCK_REC.NAIGAI, vbUnicode) = StrConv(tmpP_STOCK_REC.NAIGAI, vbUnicode) And _
                        StrConv(P_STOCK_REC.HIN_GAI, vbUnicode) = StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode) Then
                
                
                        If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
                            Call UniCode_Conv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, Format(CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)), "0000000"))
                        End If
                            
                            
                            
                        If IsNumeric(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)) Then
                            Call UniCode_Conv(tmpP_STOCK_REC.ZEN_ZAIKO_KIN, Format(CLng(StrConv(P_STOCK_REC.ZEN_ZAIKO_KIN, vbUnicode)), "0000000"))
                        End If
                            
                            
                            
                    
                        sts = BTRV(BtOpUpdate, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K2_tmpP_STOCK, Len(K2_tmpP_STOCK), 2)
                    
                        If sts Then
                            Call File_Error(sts, BtOpUpdate, "資材棚卸しﾃﾞｰﾀ")
                            Exit Function
                        End If
                
                    End If
                
                Case BtErrEOF
                
                Case Else
                    Call File_Error(sts, BtOpGetGreater, "資材棚卸しﾃﾞｰﾀ")
                    Exit Function
            End Select
        
        
        End If
        
        
        
        com = BtOpGetNext
    
    
    
    Loop
    
    
    
    
    tmpP_STOCK_MAKE_Proc = False
    
End Function

Private Function tmpP_STOCK_OUTPUT_Proc(ZEN_ZAIKO As Long) As Integer
'----------------------------------------------------------------------------
'           印刷用一時ファイルを編集出力する
'----------------------------------------------------------------------------
Dim sts As Integer
    
    tmpP_STOCK_OUTPUT_Proc = True
    '事業部
    Call UniCode_Conv(tmpP_STOCK_REC.JGYOBU, StrConv(P_STOCK_REC.JGYOBU, vbUnicode))
    '国内外
    Call UniCode_Conv(tmpP_STOCK_REC.NAIGAI, StrConv(P_STOCK_REC.NAIGAI, vbUnicode))
    '品番
    Call UniCode_Conv(tmpP_STOCK_REC.HIN_GAI, StrConv(P_STOCK_REC.HIN_GAI, vbUnicode))
    '仕入先
    Call UniCode_Conv(tmpP_STOCK_REC.CODE, StrConv(P_STOCK_REC.CODE, vbUnicode))
    '単価
    Call UniCode_Conv(tmpP_STOCK_REC.TANKA, StrConv(P_STOCK_REC.TANKA, vbUnicode))
    '日付
    Call UniCode_Conv(tmpP_STOCK_REC.INPUT_DATE, StrConv(P_STOCK_REC.INPUT_DATE, vbUnicode))
    '在庫元
    Call UniCode_Conv(tmpP_STOCK_REC.G_SYUSHI, StrConv(P_STOCK_REC.G_SYUSHI, vbUnicode))
    '前月末在庫
'    If ZEN_ZAIKO = -9999999 Then
'        Call UniCode_Conv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")
'    Else
'        Call UniCode_Conv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, Format(ZEN_ZAIKO, "00000000"))
'
'        ZEN_ZAIKO = -9999999
'    End If
    
    Call UniCode_Conv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, "00000000")
    
    Call UniCode_Conv(tmpP_STOCK_REC.ZEN_ZAIKO_KIN, "00000000")
    
    
    '当月入庫数
    Call UniCode_Conv(tmpP_STOCK_REC.NYUKO_QTY, StrConv(P_STOCK_REC.NYUKO_QTY, vbUnicode))
    '当月出庫数
    Call UniCode_Conv(tmpP_STOCK_REC.SYUKO_QTY, StrConv(P_STOCK_REC.SYUKO_QTY, vbUnicode))
    '当月在庫数
    Call UniCode_Conv(tmpP_STOCK_REC.ZAIKO_QTY, StrConv(P_STOCK_REC.ZAIKO_QTY, vbUnicode))
    '仕入単価
    Call UniCode_Conv(tmpP_STOCK_REC.TANKA, StrConv(P_STOCK_REC.TANKA, vbUnicode))
    '仕入先
    Call UniCode_Conv(tmpP_STOCK_REC.CODE, StrConv(P_STOCK_REC.CODE, vbUnicode))
    '最終出荷日
    Call UniCode_Conv(tmpP_STOCK_REC.LAST_SYUKA_DT, StrConv(P_STOCK_REC.LAST_SYUKA_DT, vbUnicode))
    '最終出庫数
    Call UniCode_Conv(tmpP_STOCK_REC.LAST_SYUKA_QTY, StrConv(P_STOCK_REC.LAST_SYUKA_QTY, vbUnicode))
    '前借残
    Call UniCode_Conv(tmpP_STOCK_REC.MAEGARI_QTY, StrConv(P_STOCK_REC.MAEGARI_QTY, vbUnicode))
    
    '元在庫数
    Call UniCode_Conv(tmpP_STOCK_REC.MOTO_ZAIKO_QTY, StrConv(P_STOCK_REC.MOTO_ZAIKO_QTY, vbUnicode))
    '出荷数計算有無
    Call UniCode_Conv(tmpP_STOCK_REC.SYUKA_NON_F, StrConv(P_STOCK_REC.SYUKA_NON_F, vbUnicode))
    
    
    '

    Do
        sts = BTRV(BtOpInsert, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K0_tmpP_STOCK, Len(K0_tmpP_STOCK), 0)
        
        Select Case sts
            Case BtNoErr, BtErrDuplicates
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                DoEvents
            Case Else
            
                Call File_Error(sts, BtOpInsert, "tmp資材棚卸ﾃﾞｰﾀ")
                Exit Function
        End Select


    Loop
    
          
          
        
        

    tmpP_STOCK_OUTPUT_Proc = False

End Function

' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
' ------------------------------------------------------------------------
Public Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function

