VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00050F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "PR00050F1.dsx":0000
End
Attribute VB_Name = "PR00050F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DET_com     As Integer

Private Const ptxSHIMUKE_CODE% = 0          '仕向け先
Private Const ptxS_YMD% = 1                 '開始　対象年月日
Private Const ptxE_YMD% = 2                 '終了　対象年月日





Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "CLASS_CODE"              'ｸﾗｽｺｰﾄﾞ
    Me.Fields.Add "DET_NAI_CNT"             '内部生産　件数
    Me.Fields.Add "DET_NAI_SURYO"           '内部生産　数量
    Me.Fields.Add "DET_GAI_CNT"             '外部生産　件数
    Me.Fields.Add "DET_GAI_SURYO"           '外部生産　数量
    Me.Fields.Add "DET_GK_CNT"              '合計　件数
    Me.Fields.Add "DET_GK_SURYO"            '合計　数量
    Me.Fields.Add "DET_GK_TANKA"            '合計　単価
    Me.Fields.Add "DET_GK_KIN"              '合計　金額

    Me.Fields.Add "DET_SHIZAI_TANKA"        '資材　単価
    Me.Fields.Add "DET_SHIZAI_KIN"          '資材　金額
    Me.Fields.Add "DET_KOURYO_TANKA"        '工料　単価
    Me.Fields.Add "DET_KOURYO_KIN"          '工料　金額
    Me.Fields.Add "DET_ETC_TANKA"           'その他　単価
    Me.Fields.Add "DET_ETC_KIN"             'その他　金額

    Me.Fields.Add "DET_GENKA_KOSOU"         '仕入原価　個装資材
    Me.Fields.Add "DET_GENKA_GAISOU"        '仕入原価　外装資材
    Me.Fields.Add "DET_GENKA_KOURYO"        '仕入原価　外注工料




    
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
    
Dim sts     As Integer
    
    sts = BTRV(DET_com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "生産実績集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    
    
    
    
    
    
    'ｸﾗｽ
    Me.Fields("CLASS_CODE") = StrConv(P_SEISAN_SUM_REC.CLASS_CODE, vbUnicode)
    '内部生産 件数
    Me.Fields("DET_NAI_CNT") = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).CNT, vbUnicode)), "#0")
    '内部生産 数量
    Me.Fields("DET_NAI_SURYO") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SURYO, vbUnicode)), "#0")
    '外部生産 件数
    Me.Fields("DET_GAI_CNT") = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).CNT, vbUnicode)), "#0")
    '外部生産 数量
    Me.Fields("DET_GAI_SURYO") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SURYO, vbUnicode)), "#0")
    '合計 件数
    Me.Fields("DET_GK_CNT") = Format(CInt(Me.Fields("DET_NAI_CNT")) + CInt(Me.Fields("DET_GAI_CNT")), "#0")
    '合計 数量
    Me.Fields("DET_GK_SURYO") = Format(CDbl(Me.Fields("DET_NAI_SURYO")) + CDbl(Me.Fields("DET_GAI_SURYO")), "#0")
    
    '合計 単価
    Me.Fields("DET_GK_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).TANKA, vbUnicode)), "#,##0.00")
    '合計　金額
    Me.Fields("DET_GK_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KINGAKU, vbUnicode)), "#,##0")
    '資材　単価
    Me.Fields("DET_SHIZAI_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_TANKA, vbUnicode)), "#,##0.00")
    '資材　金額
    Me.Fields("DET_SHIZAI_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SH_KINGAKU, vbUnicode)), "#,##0")
    
    '工料　単価
    Me.Fields("DET_KOURYO_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_TANKA, vbUnicode)), "#,##0.00")

    '工料　金額
    Me.Fields("DET_KOURYO_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KO_KINGAKU, vbUnicode)), "#,##0")
    'その他　単価
    Me.Fields("DET_ETC_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_TANKA, vbUnicode)), "#,##0.00")
    'その他　金額
    Me.Fields("DET_ETC_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).ETC_KINGAKU, vbUnicode)), "#,##0")
    
    
    '仕入原価　個装
    Me.Fields("DET_GENKA_KOSOU") = Format(CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)), "#,##0")
    '仕入原価　外装
    Me.Fields("DET_GENKA_GAISOU") = Format(CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)), "#,##0")
    '仕入原価　工料
    Me.Fields("DET_GENKA_KOURYO") = Format(CLng(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode)), "#,##0")
    
    
    
    
    
    
    
    
    DET_com = BtOpGetNext
    
            
    eof = False

End Sub

Private Sub ActiveReport_Initialize()

Dim sts     As Integer


Dim wkValue As Double

    S_YY.Text = Left(Format(PR000501.Text1(ptxS_YMD).Text, "YYYYMMDD"), 4)
    S_MM.Text = Mid(Format(PR000501.Text1(ptxS_YMD).Text, "YYYYMMDD"), 5, 2)
    S_DD.Text = Right(Format(PR000501.Text1(ptxS_YMD).Text, "YYYYMMDD"), 2)

    E_YY.Text = Left(Format(PR000501.Text1(ptxE_YMD).Text, "YYYYMMDD"), 4)
    E_MM.Text = Mid(Format(PR000501.Text1(ptxE_YMD).Text, "YYYYMMDD"), 5, 2)
    E_DD.Text = Right(Format(PR000501.Text1(ptxE_YMD).Text, "YYYYMMDD"), 2)



    Call UniCode_Conv(P_SEISAN_SUM_REC.SHIMUKE_CODE, Trim(PR000501.Text1(ptxSHIMUKE_CODE).Text))
    Call UniCode_Conv(P_SEISAN_SUM_REC.CLASS_CODE, P_ClassSum_Key)
    
    sts = BTRV(BtOpGetEqual, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
        
            MsgBox "対象ﾃﾞｰﾀがありません"
            Exit Sub
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "生産実績集計ﾃﾞｰﾀ")
            Exit Sub
    End Select


                                            '内部生産　生産件数
    NAI_CNT.Text = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).CNT, vbUnicode)), "#,##0")
                                            '外部生産　生産件数
    GAI_CNT.Text = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).CNT, vbUnicode)), "#,##0")
                                            '合計　 　 生産件数
    GK_CNT.Text = Format(CInt(NAI_CNT.Text) + CInt(GAI_CNT.Text), "#,##0")
                                            
                                            
                                            '内部生産　生産数量
    NAI_SURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SURYO, vbUnicode)), "#,##0")
                                            '外部生産　生産数量
    GAI_SURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SURYO, vbUnicode)), "#,##0")
                                            '合計　 　 生産数量
    GK_SURYO.Text = Format(CDbl(NAI_SURYO.Text) + CDbl(GAI_SURYO.Text), "#,##0")
                                            
                                            
                                            
                                            '内部生産  構成比率
    wkValue = CDbl(NAI_SURYO.Text) / (CDbl(NAI_SURYO.Text) + CDbl(GAI_SURYO.Text)) * 100
    NAI_SU_RITU.Text = Format(wkValue, "#0.00") & "%"
                                            
                                            '外部生産  構成比率
    GAI_SU_RITU.Text = Format(100 - wkValue, "#0.00") & "%"
                                            '構成  構成比率
    GK_SU_RITU.Text = "100.00%"
                                            
                                            
                                            
                                            '内部生産　生産金額
    NAI_KIN.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KINGAKU, vbUnicode)), "#,##0")
                                            '外部生産　生産金額
    GAI_KIN.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KINGAKU, vbUnicode)), "#,##0")
                                            '合計　生産金額
    GK_KIN.Text = Format(CDbl(NAI_KIN.Text) + CDbl(GAI_KIN.Text), "#,##0")
        
                                            
                                            '内部生産  構成比率
    If CDbl(NAI_KIN.Text) + CDbl(GAI_KIN.Text) = 0 Then
        wkValue = 0
    Else
        wkValue = CDbl(NAI_KIN.Text) / (CDbl(NAI_KIN.Text) + CDbl(GAI_KIN.Text)) * 100
    End If
    NAI_KIN_RITU.Text = Format(wkValue, "#0.00") & "%"
                                            
                                            '外部生産  構成比率
    GAI_KIN_RITU.Text = Format(100 - wkValue, "#0.00") & "%"
                                            '構成  構成比率
    GK_KIN_RITU.Text = "100.00%"
        
                                            
                                            '内部生産　内訳　資材
    NAI_UCHI_SHIZAI.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_KINGAKU, vbUnicode)), "#,##0")
                                            '内部生産　内訳　工料
    NAI_UCHI_KOURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_KINGAKU, vbUnicode)), "#,##0")
                                            '内部生産　内訳　その他
    NAI_UCHI_ETC.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_KINGAKU, vbUnicode)), "#,##0")
                                            '外部生産　内訳　資材
    GAI_UCHI_SHIZAI.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SH_KINGAKU, vbUnicode)), "#,##0")
                                            '外部生産　内訳　工料
    GAI_UCHI_KOURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KO_KINGAKU, vbUnicode)), "#,##0")
                                            '外部生産　内訳　その他
    GAI_UCHI_ETC.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).ETC_KINGAKU, vbUnicode)), "#,##0")
                                            '合計　内訳　資材
    GK_UCHI_SHIZAI.Text = Format(CDbl(NAI_UCHI_SHIZAI.Text) + CDbl(GAI_UCHI_SHIZAI.Text), "#,##0")
                                            '合計　内訳　工料
    GK_UCHI_KOURYO.Text = Format(CDbl(NAI_UCHI_KOURYO.Text) + CDbl(GAI_UCHI_KOURYO.Text), "#,##0")
                                            '合計　内訳　その他
    GK_UCHI_ETC.Text = Format(CDbl(NAI_UCHI_ETC.Text) + CDbl(GAI_UCHI_ETC.Text), "#,##0")
                                            
                                            
                                            '価格構成比
    KAKAKU_RITU.Text = "100.00%"
    If (CDbl(GK_UCHI_SHIZAI.Text) + CDbl(GK_UCHI_KOURYO.Text) + CDbl(GK_UCHI_ETC.Text)) = 0 Then
        SHIZAI_RITU.Text = "0.00"
    Else
        SHIZAI_RITU.Text = Format(CDbl(CDbl(GK_UCHI_SHIZAI.Text)) / (CDbl(CDbl(GK_UCHI_SHIZAI.Text) + _
                                                                        CDbl(GK_UCHI_KOURYO.Text) + _
                                                                        CDbl(GK_UCHI_ETC.Text))) * 100, "#0.00") & "%"
    End If
    If (CDbl(GK_UCHI_SHIZAI.Text) + CDbl(GK_UCHI_KOURYO.Text) + CDbl(GK_UCHI_ETC.Text)) = 0 Then
        KOURYO_RITU = "0.00"
    Else
        KOURYO_RITU = Format(CDbl(GK_UCHI_KOURYO.Text) / (CDbl(CDbl(GK_UCHI_SHIZAI.Text) + _
                                                                        CDbl(GK_UCHI_KOURYO.Text) + _
                                                                        CDbl(GK_UCHI_ETC.Text))) * 100, "#0.00") & "%"
    End If
    
    If (CDbl(GK_UCHI_SHIZAI.Text) + CDbl(GK_UCHI_KOURYO.Text) + CDbl(GK_UCHI_ETC.Text)) = 0 Then
        ETC_RITU.Text = "0.00"
    Else
    
        ETC_RITU.Text = Format(CDbl(GK_UCHI_ETC.Text) / (CDbl(GK_UCHI_SHIZAI.Text) + _
                                                            CDbl(GK_UCHI_KOURYO.Text) + _
                                                            CDbl(GK_UCHI_ETC.Text)) * 100, "#0.00") & "%"
    End If
                                            '消費資材
    GENKA_KOSOU.Text = Format(CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)), "#,##0")
    GENKA_GAISOU.Text = Format(CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)), "#,##0")
    GENKA_SHIZAI.Text = Format(CDbl(GENKA_KOSOU.Text) + CDbl(GENKA_GAISOU.Text), "#,##0")
    GENKA_KOURYO.Text = Format(CLng(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode)), "#,##0")
    GENKA_GK.Text = Format(CDbl(GENKA_KOURYO.Text) + CDbl(GENKA_SHIZAI.Text), "#,##0")
    

    





    sts = BTRV(BtOpDelete, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            Call File_Error(sts, BtOpDelete, "生産実績集計ﾃﾞｰﾀ")
            Exit Sub
    End Select



    DET_com = BtOpGetFirst

End Sub

Private Sub ActiveReport_ReportStart()
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 15
    Me.PageTopMargin = 15
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25

    Me.documentName = "商品化生産実績集計表："

    DoEvents

End Sub

