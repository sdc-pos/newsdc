VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00010F2 
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows �̊���l
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

Private DET_com         As Integer      '���ׂ�Btrieve Operation


'�v��N���p�Y��
Private Const ptxG_HANBAI_KBN% = 2          '�̔��敪
Private Const ptxKEIJYO_YM% = 3             '�Ώ۔N��






Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "DET_URIAGE_DT"       '����N����
    Me.Fields.Add "DET_TOKUI"           '���Ӑ�
    Me.Fields.Add "DET_HIN_GAI"         '�i��
    Me.Fields.Add "DET_HIN_NAME"        '�i��
    Me.Fields.Add "DET_HANBAI_KBN"      '�̔��敪
    Me.Fields.Add "DET_SYUSHI"          '���x�P��
    Me.Fields.Add "DET_URIAGE_QTY"      '����
    Me.Fields.Add "DET_TANKA"           '�P��
    Me.Fields.Add "DET_KINGAKU"         '���z





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
            Call File_Error(sts, DET_com, "���ގ��ޔ����ް�(�ꎞ̧��)")
            Exit Sub
    End Select
    
    If StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode) <> StrConv(P_tmpSHURIAGE_REC.G_SYUSHI, vbUnicode) Or _
        StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode) <> StrConv(P_tmpSHURIAGE_REC.TOKUI_CODE, vbUnicode) Then
        Exit Sub
    End If
    '����N����
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
            Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
            Exit Sub
    End Select
    
    Me.Fields("DET_TOKUI") = StrConv(P_tmpSHURIAGE_REC.TOKUI_CODE, vbUnicode) & " " & Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
    '�i��
    Me.Fields("DET_HIN_GAI") = StrConv(P_tmpSHURIAGE_REC.HIN_GAI, vbUnicode)
    '�i��
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
            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
            Exit Sub
    End Select
    '�̔��敪
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN02_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_tmpSHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "����Ͻ�")
            Exit Sub
    End Select
'    Me.Fields("DET_HANBAI_KBN") = Trim(StrConv(P_tmpSHURIAGE_REC.G_HANBAI_KBN, vbUnicode)) & " " & Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    Me.Fields("DET_HANBAI_KBN") = Trim(StrConv(P_tmpSHURIAGE_REC.G_HANBAI_KBN, vbUnicode))
    
    '���x�P��
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_tmpSHURIAGE_REC.G_SYUSHI, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "����Ͻ�")
            Exit Sub
    End Select
    Me.Fields("DET_SYUSHI") = Trim(StrConv(P_tmpSHURIAGE_REC.G_SYUSHI, vbUnicode)) & " " & Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    '����
    Me.Fields("DET_URIAGE_QTY") = Format(CDbl(StrConv(P_tmpSHURIAGE_REC.URIAGE_QTY, vbUnicode)), "#,##0")
    '�P��
    Me.Fields("DET_TANKA") = Format(CDbl(StrConv(P_tmpSHURIAGE_REC.TANKA, vbUnicode)), "#,##0.00")
    '���z
    Me.Fields("DET_KINGAKU") = Format(CDbl(StrConv(P_tmpSHURIAGE_REC.KINGAKU, vbUnicode)), "#,##0")
    
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
Dim Skip_Flg    As Boolean
 
Dim TOTAL       As Long
 
Dim YMD         As String * 8
 
 
    '���x����
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode))
    
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "����Ͻ�")
            Exit Sub
    End Select
    SYUSHI.Text = Trim(StrConv(P_SHURI_SUM_REC.G_SYUSHI, vbUnicode)) & " " & _
                    Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    '���Ӑ�
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
            Exit Sub
    End Select
    TOKUI.Text = Trim(StrConv(P_SHURI_SUM_REC.TOKUI_CODE, vbUnicode)) & " " & _
                    Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
    '�̔�
    URIAGE01.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(0).URIAGE, vbUnicode)), "#,##0")
    TOTAL = CDbl(URIAGE01.Text)
    '����
    URIAGE02.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(1).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE02.Text)
    '�ƒ�
    URIAGE03.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(2).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE03.Text)
    '���̑�
    URIAGE04.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(3).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE04.Text)
    '���̑�
    URIAGE05.Text = Format(TOTAL, "#,##0")
    '�h��
    URIAGE06.Text = Format(CDbl(StrConv(P_SHURI_SUM_REC.URIAGE_TBL(4).URIAGE, vbUnicode)), "#,##0")
    TOTAL = TOTAL + CDbl(URIAGE06.Text)
    '���v
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

    Me.documentName = "���Ӑ�ʔ���W�v�\�F"

    DoEvents

End Sub

