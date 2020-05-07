VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PI00090F2 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16845
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   29713
   _ExtentY        =   19420
   SectionData     =   "PI00090F2.dsx":0000
End
Attribute VB_Name = "PI00090F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DET_com     As Integer


Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "DET_HIN_GAI"             '子部品コード
    Me.Fields.Add "DET_HIN_NAME"            '品名
    Me.Fields.Add "DET_SO_SUU"              '子部品必要数
    
    Me.Fields.Add "DET_ZAIKO_QTY"           '在庫数
    
    Me.Fields.Add "DET_SHIJI_Z_QTY"         '注文残
    
    Me.Fields.Add "DET_HIKIATE_Z_QTY"       '引当残
    
    Me.Fields.Add "DET_FUSOKU_QTY"          '不足数
    
    Me.Fields.Add "DET_ORDER_QTY"           '注文数
    
    Me.Fields.Add "DET_LOT"                 '発注ﾛｯﾄ
    
    Me.Fields.Add "DET_ORDER_CODE"          '仕入先ｺｰﾄﾞ
    
    Me.Fields.Add "DET_ORDER_NAME"          '仕入先名称
    
    Me.Fields.Add "DET_LT"                  'ﾘｰﾄﾞﾀｲﾑ
    
    Me.Fields.Add "DET_Y_NOUKI_DT"          '予定納期
    

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
Dim wkAVE       As Long
Dim i           As Integer

    sts = BTRV(DET_com, P_SHKENTO_OSAKA_POS, P_SHKENTO_OSAKA_REC, Len(P_SHKENTO_OSAKA_REC), K1_P_SHKENTO_OSAKA, Len(K1_P_SHKENTO_OSAKA), 1)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "資材発注検討ﾌｧｲﾙ")
            Exit Sub
    End Select

    '品番(外部)
    Me.Fields("DET_HIN_GAI") = Trim(StrConv(P_SHKENTO_OSAKA_REC.HIN_GAI, vbUnicode))
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHKENTO_OSAKA_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHKENTO_OSAKA_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHKENTO_OSAKA_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
            Me.Fields("DET_HIN_NAME") = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        
        Case BtErrKeyNotFound
            Me.Fields("DET_HIN_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目ﾏｽﾀ")
            Exit Sub
    End Select
    '子部品必要数
    Me.Fields("DET_SO_SUU") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.SO_SUU, vbUnicode)), "#,##0")
    '在庫数
    Me.Fields("DET_ZAIKO_QTY") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.ZAIKO_QTY, vbUnicode)), "#,##0")
    '注文残
    Me.Fields("DET_SHIJI_Z_QTY") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.SHIJI_Z_QTY, vbUnicode)), "#,##0")
    '引当数
    Me.Fields("DET_HIKIATE_Z_QTY") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.HIKIATE_Z_QTY, vbUnicode)), "#,##0")
    '不足数
    Me.Fields("DET_FUSOKU_QTY") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.FUSOKU_QTY, vbUnicode)), "#,##0")
    '注文数
    Me.Fields("DET_ORDER_QTY") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.ORDER_QTY, vbUnicode)), "#,###")
    '発注ﾛｯﾄ
    Me.Fields("DET_LOT") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.LOT, vbUnicode)), "#,##0")
    '仕入先ｺｰﾄﾞ
    Me.Fields("DET_ORDER_CODE") = Trim(StrConv(P_SHKENTO_OSAKA_REC.ORDER_CODE, vbUnicode))
    '仕入先名
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHKENTO_OSAKA_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        
            Me.Fields("DET_ORDER_NAME") = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
        
        Case BtErrKeyNotFound
            Me.Fields("DET_ORDER_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先ﾏｽﾀ")
            Exit Sub
    End Select
    '発注ﾛｯﾄ
    Me.Fields("DET_LT") = Format(CLng(StrConv(P_SHKENTO_OSAKA_REC.LT, vbUnicode)), "#,##0")
    '納期予定日
    If Trim(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode)) = "" Then
        Me.Fields("DET_Y_NOUKI_DT") = ""
    Else
        Me.Fields("DET_Y_NOUKI_DT") = Left(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode), 4) & "/" & _
                                Mid(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(P_SHKENTO_OSAKA_REC.Y_NOUKI_DT, vbUnicode), 2)
    End If

    DET_com = BtOpGetNext

    eof = False
End Sub

Private Sub ActiveReport_Initialize()



    DET_com = BtOpGetFirst

End Sub

Private Sub ActiveReport_ReportStart()
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25


    Print_Now.Text = Format(Now, "YYYY/MM/DD HH:Mm")
End Sub

Private Sub PageHeader_Format()
    Page_Cnt.Text = Format(PI00090F2.pageNumber, "#0")
End Sub
