VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00015F1 
   ClientHeight    =   10545
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   26882
   _ExtentY        =   18600
   SectionData     =   "PR00015F1.dsx":0000
End
Attribute VB_Name = "PR00015F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '明細のBtrieve Operation


'計上年月用添字
Private Const ptxKEIJYO_YM% = 3             '対象年月






Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "TOKUI_CODE"          'CODE
    Me.Fields.Add "TOKUI_NAME"          '得意先名称
    
    Me.Fields.Add "DET01"               '販売
    Me.Fields.Add "DET02"               '製造
    Me.Fields.Add "DET03"               '家賃
    Me.Fields.Add "DET04"               'その他
    Me.Fields.Add "DET05"               '小計
    Me.Fields.Add "DET06"               '派遣
    Me.Fields.Add "DET07"               '税抜金額
    Me.Fields.Add "DET08"               '消費税
    Me.Fields.Add "DET09"               '合計


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
    
    
    sts = BTRV(DET_com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "資材資材売上集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    
    
    
    Me.Fields("TOKUI_CODE") = StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode)
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
            Me.Fields("TOKUI_NAME") = StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
        Case BtErrKeyNotFound
            Me.Fields("TOKUI_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先ﾏｽﾀ")
            Exit Sub
    End Select


    Me.Fields("DET01") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)), "#,##0")
    Me.Fields("DET02") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)), "#,##0")
    Me.Fields("DET03") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)), "#,##0")
    Me.Fields("DET04") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)), "#,##0")

    TOTAL = CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)) + _
            CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)) + _
            CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)) + _
            CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode))


    Me.Fields("DET05") = Format(TOTAL, "#,##0")

    Me.Fields("DET06") = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)), "#,##0")

    TOTAL = TOTAL + CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode))

    Me.Fields("DET07") = Format(TOTAL, "#,##0")

''    YMD = Mid(PR000101.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
''            Mid(PR000101.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
''            StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
''
''
''    If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
''        ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
''                CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
''    Else
''        ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
''                CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
''
''    End If
    ZEI = CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(5).URIAGE, vbUnicode))

    Me.Fields("DET08") = Format(ZEI, "#,##0")

    Me.Fields("DET09") = Format(TOTAL + ZEI, "#,##0")
    
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
Dim Skip_Flg    As Boolean
 
Dim TOTAL       As Long
Dim ZEI         As Long
Dim com         As Integer
 
Dim YMD         As String * 8
Dim i           As Integer
 
    Label1.Caption = PR000151.Text1(ptxKEIJYO_YM).Text                  '計上年月
    Label5.Caption = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))   'センター名


    'ｾﾞﾛﾚｺｰﾄﾞ削除
    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
        Select Case sts
            Case BtNoErr
                If Trim(StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode)) = "" Then
                Else
                    For i = 0 To 5
                        If IsNumeric(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(i).URIAGE, vbUnicode)) Then
                            If CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(i).URIAGE, vbUnicode)) <> 0 Then
                                Exit For
                            End If
                        End If
                    Next i
                
                
                    If i > 5 Then
                        sts = BTRV(BtOpDelete, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case Else
                                Call File_Error(sts, BtOpDelete, "資材売上集計ﾃﾞｰﾀ")
                                Exit Sub
                        End Select
                    End If
                
                End If
            
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "資材資材売上集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
        com = BtOpGetNext
    Loop

    '合計ﾚｺｰﾄﾞ（売掛金）の読み込み
    Call UniCode_Conv(K0_P_SHURI_SUM.TORI_KBN, P_TORI_GENERAL)
    Call UniCode_Conv(K0_P_SHURI_SUM.TOKUI_CODE, "")

    Skip_Flg = False

    sts = BTRV(BtOpGetEqual, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
                    
            Skip_Flg = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材売上集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    

    If Skip_Flg Then
        URIAGE01.Text = "0"
        URIAGE02.Text = "0"
        URIAGE03.Text = "0"
        URIAGE04.Text = "0"
        URIAGE05.Text = "0"
        URIAGE06.Text = "0"
        URIAGE07.Text = "0"
        URIAGE08.Text = "0"
        URIAGE09.Text = "0"
    Else
        URIAGE01.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)), "#,##0")
        URIAGE02.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)), "#,##0")
        URIAGE03.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)), "#,##0")
        URIAGE04.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)), "#,##0")

        TOTAL = CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)) + _
                CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)) + _
                CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)) + _
                CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode))


        URIAGE05.Text = Format(TOTAL, "#,##0")

        URIAGE06.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)), "#,##0")

        TOTAL = TOTAL + CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode))

        URIAGE07.Text = Format(TOTAL, "#,##0")

        YMD = Mid(PR000151.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
                Mid(PR000151.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
                StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
        
        
''        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
''            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
''                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
''        Else
''            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
''                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
''
''        End If

        
        ZEI = CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(5).URIAGE, vbUnicode))
        
        URIAGE08.Text = Format(ZEI, "#,##0")

        URIAGE09.Text = Format(TOTAL + ZEI, "#,##0")
    
        sts = BTRV(BtOpDelete, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "資材売上集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
    
    
    End If

    '合計ﾚｺｰﾄﾞ（振替）の読み込み
    Call UniCode_Conv(K0_P_SHURI_SUM.TORI_KBN, P_TORI_SYANAI)
    Call UniCode_Conv(K0_P_SHURI_SUM.TOKUI_CODE, "")

    Skip_Flg = False

    sts = BTRV(BtOpGetEqual, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
                    
            Skip_Flg = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材売上集計ﾃﾞｰﾀ")
            Exit Sub
    End Select


    If Skip_Flg Then
        FURIKAE01.Text = "0"
        FURIKAE02.Text = "0"
        FURIKAE03.Text = "0"
        FURIKAE04.Text = "0"
        FURIKAE05.Text = "0"
        FURIKAE06.Text = "0"
        FURIKAE07.Text = "0"
        FURIKAE08.Text = "0"
        FURIKAE09.Text = "0"
    Else
        FURIKAE01.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)), "#,##0")
        FURIKAE02.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)), "#,##0")
        FURIKAE03.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)), "#,##0")
        FURIKAE04.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)), "#,##0")

        TOTAL = CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)) + _
                CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)) + _
                CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)) + _
                CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode))


        FURIKAE05.Text = Format(TOTAL, "#,##0")

        FURIKAE06.Text = Format(CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)), "#,##0")

        TOTAL = TOTAL + CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode))

        FURIKAE07.Text = Format(TOTAL, "#,##0")

''        YMD = Mid(PR000101.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
''                Mid(PR000101.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
''                StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
''
''
''        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
''            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
''                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
''        Else
''            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
''                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
''
''        End If

        ZEI = CLng(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(5).URIAGE, vbUnicode))
        
        
        FURIKAE08.Text = Format(ZEI, "#,##0")

        FURIKAE09.Text = Format(TOTAL + ZEI, "#,##0")
    
        sts = BTRV(BtOpDelete, P_SHURI_SUM_POS, P_SHURI_SUM_REC, Len(P_SHURI_SUM_REC), K0_P_SHURI_SUM, Len(K0_P_SHURI_SUM), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpDelete, "資材売上集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
    
    
    
    End If


    TOTAL01.Text = Format(CLng(URIAGE01.Text) + CLng(FURIKAE01.Text), "#,##0")
    TOTAL02.Text = Format(CLng(URIAGE02.Text) + CLng(FURIKAE02.Text), "#,##0")
    TOTAL03.Text = Format(CLng(URIAGE03.Text) + CLng(FURIKAE03.Text), "#,##0")
    TOTAL04.Text = Format(CLng(URIAGE04.Text) + CLng(FURIKAE04.Text), "#,##0")
    TOTAL05.Text = Format(CLng(URIAGE05.Text) + CLng(FURIKAE05.Text), "#,##0")
    TOTAL06.Text = Format(CLng(URIAGE06.Text) + CLng(FURIKAE06.Text), "#,##0")
    TOTAL07.Text = Format(CLng(URIAGE07.Text) + CLng(FURIKAE07.Text), "#,##0")
    TOTAL08.Text = Format(CLng(URIAGE08.Text) + CLng(FURIKAE08.Text), "#,##0")
    TOTAL09.Text = Format(CLng(URIAGE09.Text) + CLng(FURIKAE09.Text), "#,##0")


    Call UniCode_Conv(K0_P_SHURI_SUM.TORI_KBN, P_TORI_GENERAL)
    Call UniCode_Conv(K0_P_SHURI_SUM.TOKUI_CODE, "")

    DET_com = BtOpGetGreater


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

    Me.documentName = "得意先別売上集計表："

    DoEvents

End Sub

