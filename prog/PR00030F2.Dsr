VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00030F2 
   ClientHeight    =   12405
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   33655
   _ExtentY        =   21881
   SectionData     =   "PR00030F2.dsx":0000
End
Attribute VB_Name = "PR00030F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '明細のBtrieve Operation








Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "DET_HIN_GAI"         '品番外部
    Me.Fields.Add "DET_HIN_NAME"        '品名
    Me.Fields.Add "DET_G_SYUSHI"        '在庫元
    Me.Fields.Add "DET_ZEN_ZAIKO_QTY"   '前月在庫
    Me.Fields.Add "DET_NYUKO_QTY"       '入庫数
    Me.Fields.Add "DET_SYUKO_QTY"       '出庫数
    Me.Fields.Add "DET_ZAIKO_QTY"       '在庫数
    Me.Fields.Add "DET_SHI_TANKA"       '仕入単価
    Me.Fields.Add "DET_ZAIKO_KIN"       '当月在庫単価
    Me.Fields.Add "DET_SHI_CODE"        '仕入先
    Me.Fields.Add "DET_LAST_SYUKA_DT"   '最終出荷日
    Me.Fields.Add "DET_LAST_SYUKA_QTY"  '最終出荷数
    Me.Fields.Add "DET_MAEGARI_QTY"     '前借数





End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

    
    
'    sts = BTRV(DET_com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K1_tmpP_STOCK, Len(K1_tmpP_STOCK), 1)
    sts = BTRV(DET_com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K2_tmpP_STOCK, Len(K2_tmpP_STOCK), 2)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "tmp資材棚卸ﾃﾞｰﾀ")
            Exit Sub
    End Select
    
    
    
If Trim(StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode)) = "K203" Then
    Debug.Print
End If
    
    
    '品番
    Me.Fields("DET_HIN_GAI") = StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode)
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(tmpP_STOCK_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(tmpP_STOCK_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Me.Fields("DET_HIN_NAME") = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Me.Fields("DET_HIN_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Sub
    End Select
    '在庫元
    Me.Fields("DET_G_SYUSHI") = StrConv(tmpP_STOCK_REC.G_SYUSHI, vbUnicode)
    '前月在庫数
    If IsNumeric(StrConv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
        Me.Fields("DET_ZEN_ZAIKO_QTY") = Format(CLng(StrConv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)), "#,###")
    Else
        Me.Fields("DET_ZEN_ZAIKO_QTY") = ""
    End If
    '入庫数
        'Clng --> Val 2016.01.08
    Me.Fields("DET_NYUKO_QTY") = Format(Val(StrConv(tmpP_STOCK_REC.NYUKO_QTY, vbUnicode)), "#,###")
    '出庫数
        'Clng --> Val 2016.01.08
    Me.Fields("DET_SYUKO_QTY") = Format(Val(StrConv(tmpP_STOCK_REC.SYUKO_QTY, vbUnicode)), "#,###")
    '当月在庫
        'Clng --> Val 2016.01.08
    Me.Fields("DET_ZAIKO_QTY") = Format(Val(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,###")
    '仕入単価
        'CDbl --> Val 2016.01.08
    Me.Fields("DET_SHI_TANKA") = Format(Val(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)), "#,##0.00")
    '当月在庫金額
'    Me.Fields("DET_ZAIKO_KIN") = Format(CDbl(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) * CLng(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,###")
'
    '>>>>>>>>>> 2016.01.08
    If Not IsNumeric(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) Then
        Call UniCode_Conv(tmpP_STOCK_REC.TANKA, "00000000")
    End If
    If Not IsNumeric(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
        Call UniCode_Conv(tmpP_STOCK_REC.ZAIKO_QTY, "00000000")
    End If
    '>>>>>>>>>> 2016.01.08
    
    Me.Fields("DET_ZAIKO_KIN") = Format(ToRoundUp(CCur(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) * _
                                    CCur(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)), 0), "#,###")
    
    '在庫元
    Me.Fields("DET_SHI_CODE") = StrConv(tmpP_STOCK_REC.CODE, vbUnicode)
    '最終出荷日
    If Trim(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode)) = "" Then
        Me.Fields("DET_LAST_SYUKA_DT") = ""
    Else
        Me.Fields("DET_LAST_SYUKA_DT") = Left(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 4) & "/" & _
                                            Mid(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                           Right(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 2)
    End If
    
    Me.Fields("DET_LAST_SYUKA_QTY") = ""
    If IsNumeric(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) Then
        If CLng(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) <> 0 Then
            Me.Fields("DET_LAST_SYUKA_QTY") = Format(CLng(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)), "#,###")
        End If
    End If
    '前借
    If IsNumeric(StrConv(tmpP_STOCK_REC.MAEGARI_QTY, vbUnicode)) Then
        Me.Fields("DET_MAEGARI_QTY") = Format(CDbl(StrConv(tmpP_STOCK_REC.MAEGARI_QTY, vbUnicode)), "#,###")
    Else
        Me.Fields("DET_MAEGARI_QTY") = ""
    End If
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    
    

End Sub

Private Sub ActiveReport_Initialize()

    
    
Dim sts     As Integer
    
    
    
    
    
    
    
    
    
    
    DET_com = BtOpGetFirst
    
    lblYY.Caption = Left(PR000301.Text1(1).Text, 4)
    lblMM.Caption = Right(PR000301.Text1(1).Text, 2)
 
 
    SHIME_YMD.Caption = "締日：" & PR000301.Text1(0)
 
 
 
 
    Center_Name.Caption = StrConv(P_KANRIREC.Center_Name, vbUnicode)



    DET_com = BtOpGetFirst
 


End Sub


Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    lblPrint_Now = Format(Now, "YYYY/MM/DD HH:MM")
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25

    Me.documentName = "資材棚卸表"

    DoEvents

End Sub

