VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00010F2 
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Μωθl
   _ExtentX        =   26882
   _ExtentY        =   17621
   SectionData     =   "PR00010F2.dsx":0000
End
Attribute VB_Name = "PR00010F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      'ΎΧΜBtrieve Operation


'vγNpY
Private Const ptxG_HANBAI_KBN% = 2          'Μζͺ
Private Const ptxKEIJYO_YM% = 3             'ΞΫN






Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "DET_URIAGE_DT"       'γNϊ
    Me.Fields.Add "DET_TOKUI"           'ΎΣζ
    Me.Fields.Add "DET_HIN_GAI"         'iΤ
    Me.Fields.Add "DET_HIN_NAME"        'iΌ
    Me.Fields.Add "DET_HANBAI_KBN"      'Μζͺ
    Me.Fields.Add "DET_SYUSHI"          'ϋxPΚ
    Me.Fields.Add "DET_URIAGE_QTY"      'Κ
    Me.Fields.Add "DET_TANKA"           'PΏ
    Me.Fields.Add "DET_KINGAKU"         'ΰz





End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
    
    
    sts = BTRV(DET_com, P_tmpSHURIAGE_POS, P_tmpSHURIAGE_REC, Len(P_tmpSHURIAGE_REC), K0_P_tmpSHURIAGE, Len(K0_P_tmpSHURIAGE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "ήήγΓή°ΐ(κΜ§²Ω)")
            Exit Sub
    End Select
    
    If StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode) <> StrConv(P_tmpSHURIAGE_REC.G_SYUSHI, vbUnicode) Or _
        StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode) <> StrConv(P_tmpSHURIAGE_REC.TOKUI_CODE, vbUnicode) Then
        Exit Sub
    End If
    'γNϊ
    Me.Fields("DET_URIAGE_DT") = Mid(StrConv(P_tmpSHURIAGE_REC.URIAGE_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_tmpSHURIAGE_REC.URIAGE_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_tmpSHURIAGE_REC.URIAGE_DT, vbUnicode), 7, 2)
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_tmpSHURIAGE_REC.TOKUI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "σ₯ζΟ½ΐ")
            Exit Sub
    End Select
    
    Me.Fields("DET_TOKUI") = StrConv(P_tmpSHURIAGE_REC.TOKUI_CODE, vbUnicode) & " " & Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
    'iΤ
    Me.Fields("DET_HIN_GAI") = StrConv(P_tmpSHURIAGE_REC.HIN_GAI, vbUnicode)
    'iΌ
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_tmpSHURIAGE_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_tmpSHURIAGE_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_tmpSHURIAGE_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Me.Fields("DET_HIN_NAME") = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Me.Fields("DET_HIN_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "iΪΟ½ΐ")
            Exit Sub
    End Select
    'Μζͺ
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN02_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_tmpSHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "Ί°ΔήΟ½ΐ")
            Exit Sub
    End Select
'    Me.Fields("DET_HANBAI_KBN") = Trim(StrConv(P_tmpSHURIAGE_REC.G_HANBAI_KBN, vbUnicode)) & " " & Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    Me.Fields("DET_HANBAI_KBN") = Trim(StrConv(P_tmpSHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
    
    'ϋxPΚ
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_tmpSHURIAGE_REC.G_SYUSHI, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "Ί°ΔήΟ½ΐ")
            Exit Sub
    End Select
    Me.Fields("DET_SYUSHI") = Trim(StrConv(P_tmpSHURIAGE_REC.G_SYUSHI, vbUnicode)) & " " & Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    'Κ
    Me.Fields("DET_URIAGE_QTY") = Format(CDbl(StrConv(P_tmpSHURIAGE_REC.URIAGE_QTY, vbUnicode)), "#,##0")
    'PΏ
    Me.Fields("DET_TANKA") = Format(CDbl(StrConv(P_tmpSHURIAGE_REC.TANKA, vbUnicode)), "#,##0.00")
    'ΰz
    Me.Fields("DET_KINGAKU") = Format(CDbl(StrConv(P_tmpSHURIAGE_REC.KINGAKU, vbUnicode)), "#,##0")
    
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
Dim Skip_Flg    As Boolean
 
Dim TOTAL       As Long
 
Dim YMD         As String * 8
 
 
    'ϋxΌΜ
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode))
    
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "Ί°ΔήΟ½ΐ")
            Exit Sub
    End Select
    SYUSHI.Text = Trim(StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode)) & " " & _
                    Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    'ΎΣζ
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "σ₯ζΟ½ΐ")
            Exit Sub
    End Select
    TOKUI.Text = Trim(StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode)) & " " & _
                    Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
    'Μ
    URIAGE01.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)), "#,##0")
    TOTAL = CDbl(URIAGE01.Text)
    '»’
    URIAGE02.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE02.Text)
    'Ζΐ
    URIAGE03.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE03.Text)
    '»ΜΌ
    URIAGE04.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE04.Text)
    '»ΜΌ
    URIAGE05.Text = Format(TOTAL, "#,##0")
    'h­
    URIAGE06.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE06.Text)
    'v
    URIAGE07.Text = Format(TOTAL, "#,##0")






    Call UniCode_Conv(K0_P_tmpSHURIAGE.KEIJYO_YM, Mid(PR000101.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
                                                    Mid(PR000101.Text1(ptxKEIJYO_YM).Text, 6, 2))
    Call UniCode_Conv(K0_P_tmpSHURIAGE.G_SYUSHI, StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode))
    Call UniCode_Conv(K0_P_tmpSHURIAGE.TOKUI_CODE, StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode))
    Call UniCode_Conv(K0_P_tmpSHURIAGE.URIAGE_DT, "")
    Call UniCode_Conv(K0_P_tmpSHURIAGE.URIAGE_NO, "")


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

    Me.documentName = "ΎΣζΚγWv\F"

    DoEvents

End Sub

