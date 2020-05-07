VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00060F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "PR00060F1.dsx":0000
End
Attribute VB_Name = "PR00060F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DET_com     As Integer

Private Const ptxS_YMD% = 0                 '開始　対象年月日
Private Const ptxE_YMD% = 1                 '終了　対象年月日





Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "TORI_CODE"               '取引先
    Me.Fields.Add "KIN01"                   '仕向け先別金額（1）
    Me.Fields.Add "KIN02"                   '仕向け先別金額（2）
    Me.Fields.Add "KIN03"                   '仕向け先別金額（3）
    Me.Fields.Add "KIN04"                   '仕向け先別金額（4）
    Me.Fields.Add "KIN05"                   '仕向け先別金額（5）
    Me.Fields.Add "KIN06"                   '仕向け先別金額（6）
    Me.Fields.Add "KIN07"                   '仕向け先別金額（7）
    Me.Fields.Add "KIN08"                   '仕向け先別金額（8）
    Me.Fields.Add "KIN09"                   '仕向け先別金額（9）
    Me.Fields.Add "KIN10"                   '仕向け先別金額（10）
    Me.Fields.Add "TOTAL"                   '合計
    Me.Fields.Add "ZEI"                     '消費税
    Me.Fields.Add "SHIHARAI"                '支払い




    
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
    
Dim sts         As Integer

Dim TOTAL       As Long
Dim ZEI         As Long
Dim i           As Integer
    
    sts = BTRV(DET_com, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "生産実績明細集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    
    
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
    
    
    '受払先
    Me.Fields("TORI_CODE") = StrConv(P_SEISAN_GK_REC.TORI_CODE, vbUnicode) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
    
    '内訳金額（１）
    If Trim(SHIMUKE_TBL(0)) = "" Then
        Me.Fields("KIN01") = ""
    Else
        Me.Fields("KIN01") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(0).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（２）
    If Trim(SHIMUKE_TBL(1)) = "" Then
        Me.Fields("KIN02") = ""
    Else
        Me.Fields("KIN02") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(1).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（３）
    If Trim(SHIMUKE_TBL(2)) = "" Then
        Me.Fields("KIN03") = ""
    Else
        Me.Fields("KIN03") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(2).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（４）
    If Trim(SHIMUKE_TBL(3)) = "" Then
        Me.Fields("KIN04") = ""
    Else
        Me.Fields("KIN04") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(3).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（５）
    If Trim(SHIMUKE_TBL(4)) = "" Then
        Me.Fields("KIN05") = ""
    Else
        Me.Fields("KIN05") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(4).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（６）
    If Trim(SHIMUKE_TBL(5)) = "" Then
        Me.Fields("KIN06") = ""
    Else
        Me.Fields("KIN06") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(5).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（７）
    If Trim(SHIMUKE_TBL(6)) = "" Then
        Me.Fields("KIN07") = ""
    Else
        Me.Fields("KIN07") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(6).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（８）
    If Trim(SHIMUKE_TBL(7)) = "" Then
        Me.Fields("KIN08") = ""
    Else
        Me.Fields("KIN08") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(7).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（９）
    If Trim(SHIMUKE_TBL(8)) = "" Then
        Me.Fields("KIN09") = ""
    Else
        Me.Fields("KIN09") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(8).KIN, vbUnicode)), "#,##0")
    End If
    '内訳金額（10）
    If Trim(SHIMUKE_TBL(9)) = "" Then
        Me.Fields("KIN10") = ""
    Else
        Me.Fields("KIN10") = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(9).KIN, vbUnicode)), "#,##0")
    End If
    
    For i = 0 To UBound(SHIMUKE_TBL)
        
        TOTAL = TOTAL + CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, vbUnicode))
    
    Next i
    Me.Fields("TOTAL") = Format(TOTAL, "#,##0")
    
    
    
    Select Case StrConv(P_SEISAN_GK_REC.TORI_KBN, vbUnicode)
        Case P_TORI_GENERAL
            
            If GAICYU_F Then        '2007.07.17
                Me.Fields("ZEI") = ""
                Me.Fields("SHIHARAI") = Format(TOTAL, "#,##0")
            Else
                ZEI = Int(Int(TOTAL * CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode) / 100)) + CInt(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode) / 10))
                Me.Fields("ZEI") = Format(ZEI, "#,##0")
                Me.Fields("SHIHARAI") = Format(TOTAL + ZEI, "#,##0")
            End If
        Case Else
            Me.Fields("ZEI") = ""
            Me.Fields("SHIHARAI") = Format(TOTAL, "#,##0")
    End Select

    
    
    
    
    DET_com = BtOpGetNext
    
            
    eof = False

End Sub

Private Sub ActiveReport_Initialize()

Dim sts     As Integer
Dim i       As Integer

Dim TOTAL   As Double
Dim ZEI     As Double

    S_YY.Text = Left(Format(PR000601.Text1(ptxS_YMD).Text, "YYYYMMDD"), 4)
    S_MM.Text = Mid(Format(PR000601.Text1(ptxS_YMD).Text, "YYYYMMDD"), 5, 2)
    S_DD.Text = Right(Format(PR000601.Text1(ptxS_YMD).Text, "YYYYMMDD"), 2)

    E_YY.Text = Left(Format(PR000601.Text1(ptxE_YMD).Text, "YYYYMMDD"), 4)
    E_MM.Text = Mid(Format(PR000601.Text1(ptxE_YMD).Text, "YYYYMMDD"), 5, 2)
    E_DD.Text = Right(Format(PR000601.Text1(ptxE_YMD).Text, "YYYYMMDD"), 2)


    Call UniCode_Conv(K0_P_SEISAN_GK.TORI_CODE, "")
    sts = BTRV(BtOpGetEqual, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
        
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrKeyNotFound
        
            MsgBox "対象ﾃﾞｰﾀがありません"
            Exit Sub
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "生産実績明細集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    
    
    If Trim(SHIMUKE_TBL(0)) = "" Then
        SHIMUKE_CD01.Text = ""
        SHIMUKE_NM01.Text = ""
        SHIMUKE_GK01.Text = ""
    Else
        SHIMUKE_CD01.Text = SHIMUKE_TBL(0)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(0))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM01.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM01.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK01.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(0).KIN, vbUnicode)), "#,##0")
    End If

    If Trim(SHIMUKE_TBL(1)) = "" Then
        SHIMUKE_CD02.Text = ""
        SHIMUKE_NM02.Text = ""
        SHIMUKE_GK02.Text = ""
    Else
        SHIMUKE_CD02.Text = SHIMUKE_TBL(1)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(1))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM02.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM02.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK02.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(1).KIN, vbUnicode)), "#,##0")
    End If

    If Trim(SHIMUKE_TBL(2)) = "" Then
        SHIMUKE_CD03.Text = ""
        SHIMUKE_NM03.Text = ""
        SHIMUKE_GK03.Text = ""
    Else
        SHIMUKE_CD03.Text = SHIMUKE_TBL(2)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(2))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM03.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM03.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK02.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(2).KIN, vbUnicode)), "#,##0")
    End If

    If Trim(SHIMUKE_TBL(3)) = "" Then
        SHIMUKE_CD04.Text = ""
        SHIMUKE_NM04.Text = ""
        SHIMUKE_GK04.Text = ""
    Else
        SHIMUKE_CD04.Text = SHIMUKE_TBL(3)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(3))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM04.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM04.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK04.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(3).KIN, vbUnicode)), "#,##0")
    End If

    If Trim(SHIMUKE_TBL(4)) = "" Then
        SHIMUKE_CD05.Text = ""
        SHIMUKE_NM05.Text = ""
        SHIMUKE_GK05.Text = ""
    Else
        SHIMUKE_CD05.Text = SHIMUKE_TBL(4)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(4))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM05.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM05.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK05.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(4).KIN, vbUnicode)), "#,##0")
    End If


    If Trim(SHIMUKE_TBL(5)) = "" Then
        SHIMUKE_CD06.Text = ""
        SHIMUKE_NM06.Text = ""
        SHIMUKE_GK06.Text = ""
    Else
        SHIMUKE_CD06.Text = SHIMUKE_TBL(5)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(5))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM06.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM06.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK06.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(5).KIN, vbUnicode)), "#,##0")
    End If


    If Trim(SHIMUKE_TBL(6)) = "" Then
        SHIMUKE_CD07.Text = ""
        SHIMUKE_NM07.Text = ""
        SHIMUKE_GK07.Text = ""
    Else
        SHIMUKE_CD07.Text = SHIMUKE_TBL(6)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(6))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM07.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM07.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK07.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(6).KIN, vbUnicode)), "#,##0")
    End If


    If Trim(SHIMUKE_TBL(7)) = "" Then
        SHIMUKE_CD08.Text = ""
        SHIMUKE_NM08.Text = ""
        SHIMUKE_GK08.Text = ""
    Else
        SHIMUKE_CD08.Text = SHIMUKE_TBL(7)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(7))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM08.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM08.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK08.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(7).KIN, vbUnicode)), "#,##0")
    End If



    If Trim(SHIMUKE_TBL(8)) = "" Then
        SHIMUKE_CD09.Text = ""
        SHIMUKE_NM09.Text = ""
        SHIMUKE_GK09.Text = ""
    Else
        SHIMUKE_CD09.Text = SHIMUKE_TBL(8)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(8))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM09.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM09.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK09.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(8).KIN, vbUnicode)), "#,##0")
    End If



    If Trim(SHIMUKE_TBL(9)) = "" Then
        SHIMUKE_CD10.Text = ""
        SHIMUKE_NM10.Text = ""
        SHIMUKE_GK10.Text = ""
    Else
        SHIMUKE_CD10.Text = SHIMUKE_TBL(9)
        
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, SHIMUKE_TBL(9))
        
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                SHIMUKE_NM10.Text = StrConv(P_CODEREC.C_RNAME, vbUnicode)
            
            
            
            Case BtErrKeyNotFound
                SHIMUKE_NM10.Text = ""
            Case Else
                Call File_Error(sts, BtOpGetEqual, "ｺｰﾄﾞﾏｽﾀ")
                Exit Sub
        End Select
            
        SHIMUKE_GK10.Text = Format(CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(9).KIN, vbUnicode)), "#,##0")
    End If


    TOTAL = 0
    For i = 0 To UBound(SHIMUKE_TBL)
        TOTAL = TOTAL + CLng(StrConv(P_SEISAN_GK_REC.UCHIWAKE_TBL(i).KIN, vbUnicode))
    Next i

    GK_TOTAL.Text = Format(TOTAL, "#,##0")
    
    If GAICYU_F Then    '2007.07.17
        GK_ZEI.Text = ""
        GK_SHIHARAI.Text = Format(TOTAL, "#,##0")
    Else
        ZEI = Int(Int(CLng(StrConv(P_SEISAN_GK_REC.KAZEI, vbUnicode)) * CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode) / 100)) + CInt(StrConv(P_KANRIREC.NOW_MARUME, vbUnicode) / 10))
        GK_ZEI.Text = Format(ZEI, "#,##0")
        GK_SHIHARAI.Text = Format(TOTAL + ZEI, "#,##0")
    End If
    
    

    sts = BTRV(BtOpDelete, P_SEISAN_GK_POS, P_SEISAN_GK_REC, Len(P_SEISAN_GK_REC), K0_P_SEISAN_GK, Len(K0_P_SEISAN_GK), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            Call File_Error(sts, BtOpDelete, "生産実績明細集計ﾃﾞｰﾀ")
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
    
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25

    Me.documentName = "商品化生産実績集計表："

    DoEvents

End Sub

