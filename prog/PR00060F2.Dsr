VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00060F2 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "PR00060F2.dsx":0000
End
Attribute VB_Name = "PR00060F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DET_com     As Integer

Private Const ptxS_YMD% = 0                 '開始　対象年月日
Private Const ptxE_YMD% = 1                 '終了　対象年月日





Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "UKEIRE_DT"               '受入日
    Me.Fields.Add "SHIJI_NO"                '指示№
    Me.Fields.Add "SHIMUKE_CODE"            '仕向先
    Me.Fields.Add "HIN_GAI"                 '品番
    Me.Fields.Add "UKEIRE_QTY"              '数量
    Me.Fields.Add "S_CLASS_CODE"            '商品化ｸﾗｽ
    Me.Fields.Add "F_CLASS_CODE"            '付加ｸﾗｽ
    Me.Fields.Add "N_CLASS_CODE"            '内職ｸﾗｽ
    Me.Fields.Add "KOURYOU"                 '単価
    Me.Fields.Add "KIN"                     '金額




    
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
    
Dim sts         As Integer

    
    sts = BTRV(DET_com, P_SEISAN_DET_POS, P_SEISAN_DET_REC, Len(P_SEISAN_DET_REC), K0_P_SEISAN_DET, Len(K0_P_SEISAN_DET), 0)
    Select Case sts
        Case BtNoErr
        
            If StrConv(P_SEISAN_DET_REC.TORI_CODE, vbUnicode) <> StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode) Then
                Exit Sub
            End If
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "生産実績明細集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    
    
    Me.Fields("UKEIRE_DT") = Left(StrConv(P_SEISAN_DET_REC.UKEIRE_DT, vbUnicode), 4) & "/" & _
                                Mid(StrConv(P_SEISAN_DET_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(P_SEISAN_DET_REC.UKEIRE_DT, vbUnicode), 2)

    
    '仕向け先
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SEISAN_DET_REC.SHIMUKE_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
            Exit Sub
    End Select
    
    '指示№
    Me.Fields("SHIJI_NO") = StrConv(P_SEISAN_DET_REC.SHIJI_NO, vbUnicode)
    
    '仕向け先
    Me.Fields("SHIMUKE_CODE") = StrConv(P_SEISAN_DET_REC.SHIMUKE_CODE, vbUnicode) & " " & Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    '品番
    Me.Fields("HIN_GAI") = Trim(StrConv(P_SEISAN_DET_REC.HIN_GAI, vbUnicode))
    '数量
    Me.Fields("UKEIRE_QTY") = Format(CDbl(StrConv(P_SEISAN_DET_REC.UKEIRE_QTY, vbUnicode)), "#,##0.00")
    '商品化クラス
    Me.Fields("S_CLASS_CODE") = Trim(StrConv(P_SEISAN_DET_REC.S_CLASS_CODE, vbUnicode))
    '付加クラス
    Me.Fields("F_CLASS_CODE") = Trim(StrConv(P_SEISAN_DET_REC.F_CLASS_CODE, vbUnicode))
    '内職クラス
    Me.Fields("N_CLASS_CODE") = Trim(StrConv(P_SEISAN_DET_REC.N_CLASS_CODE, vbUnicode))
    '単価
    Me.Fields("KOURYOU") = Format(CDbl(StrConv(P_SEISAN_DET_REC.KOURYOU, vbUnicode)), "#,##0.00")
    '金額
    Me.Fields("KIN") = Format(CLng(StrConv(P_SEISAN_DET_REC.KIN, vbUnicode)), "#,##0")
    
    
    
    
    
    DET_com = BtOpGetNext
    
            
    eof = False

End Sub

Private Sub ActiveReport_Initialize()

Dim sts     As Integer
Dim i       As Integer

Dim TOTAL   As Double
    '会社名
    KAISHA_NAME.Caption = Trim(StrConv(P_KANRIREC.KAISHA_NAME, vbUnicode))
    
    'センター名
    CENTER_NAME.Caption = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))


    
    '手配先
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, DET_com, "受払先ﾏｽﾀ")
            Exit Sub
    End Select
        
    TORI_NAME.Caption = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode))
    
    
    YYMM.Text = Left(Format(PR000601.Text1(ptxE_YMD).Text, "YYYYMMDD"), 4) & "年" & _
                Mid(Format(PR000601.Text1(ptxE_YMD).Text, "YYYYMMDD"), 5, 2) & "月分"
 
    '件数
    GK_CNT.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.CNT, vbUnicode)), "#,##0")
    '数量
    GK_QTY.Text = Format(CDbl(StrConv(P_SEISAN_GK_REC.QTY, vbUnicode)), "#,##0.00")
            
    TOTAL = 0
    For i = 0 To UBound(SHIMUKE_TBL)
    
        If Trim(SHIMUKE_TBL(i)) = "" Then
        Else
    
            TOTAL = TOTAL + CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, vbUnicode))
        End If
    
    
    Next i
    '金額
    GK_KIN.Text = Format(TOTAL, "#,##0")
        
 
 
 
    Call UniCode_Conv(K0_P_SEISAN_DET.TORI_CODE, StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode))
    DET_com = BtOpGetGreaterEqual

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

    Me.documentName = "商品化生産実績集計表："

    DoEvents

End Sub

