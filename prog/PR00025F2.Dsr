VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00025F2 
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   26882
   _ExtentY        =   17621
   SectionData     =   "PR00025F2.dsx":0000
End
Attribute VB_Name = "PR00025F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '���ׂ�Btrieve Operation

Private Print_Mode      As Integer

'�w�����
Private Const ptxKEIJYO_YM% = 0         '�Ώ۔N��
Private Const ptxSHIIRE_CODE% = 1       '�d���溰��






Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "DET_UKEIRE_DT"       '�������
    Me.Fields.Add "DET_HIN_GAI"         '�i��
    Me.Fields.Add "DET_HIN_NAME"        '�i��
    Me.Fields.Add "DET_SHIIRE_NAME"     '�d����
    Me.Fields.Add "DET_SHIIRE_KBN"      '�d���敪
    Me.Fields.Add "DET_SYUSHI"          '���x�P��
    Me.Fields.Add "DET_UKEIRE_QTY"      '�������
    Me.Fields.Add "DET_UKEIRE_TANKA"    '����P��
    Me.Fields.Add "DET_UKEIRE_KINGAKU"  '������z





End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

 
    
    
    If Print_Mode = 0 Then
        sts = BTRV(DET_com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K2_P_SHUKEIRE, Len(K2_P_SHUKEIRE), 2)
    Else
        sts = BTRV(DET_com, P_SHUKEIRE_POS, P_SHUKEIRE_REC, Len(P_SHUKEIRE_REC), K1_P_SHUKEIRE, Len(K1_P_SHUKEIRE), 1)
    End If
    Select Case sts
        Case BtNoErr
        
            If StrConv(P_SHUKEIRE_REC.KEIJYO_YM, vbUnicode) <> Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
                                                                Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2) Then
                Exit Sub
            End If
        
            If Print_Mode = 1 Then
                If Trim(StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode)) <> PR000251.Text1(ptxSHIIRE_CODE).Text Then
                    Exit Sub
                End If
            End If
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "���ގ���ް�")
            Exit Sub
    End Select
    
    
    '�����f�[�^�ǂݍ���
    Call UniCode_Conv(K0_P_SHORDER.ORDER_NO, StrConv(P_SHUKEIRE_REC.ORDER_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K0_P_SHORDER, Len(K0_P_SHORDER), 0)
    Select Case sts
        Case BtNoErr
            
            '�i��
            Me.Fields("DET_HIN_GAI") = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
            '�i��
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHORDER_REC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHORDER_REC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Sub
            End Select
            Me.Fields("DET_HIN_NAME") = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
            '�d���敪
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN01_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                    Exit Sub
            End Select
            Me.Fields("DET_SHIIRE_KBN") = Trim(StrConv(P_SHORDER_REC.G_SHIIRE_KBN, vbUnicode)) & " " & _
                                        Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
    
            '���x�敪
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                    Exit Sub
            End Select
        
            Me.Fields("DET_SYUSHI") = Trim(StrConv(P_SHORDER_REC.G_SYUSHI, vbUnicode)) & " " & _
                                        Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))
        
        
        Case BtErrKeyNotFound
                    
            Me.Fields("DET_HIN_GAI") = "****"
            Me.Fields("DET_SHIIRE_KBN") = "**"
            Me.Fields("DET_SYUSHI") = "***"
                    
                    
                    
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
            Exit Sub
    End Select
    
    '����N����
    Me.Fields("DET_UKEIRE_DT") = Mid(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_SHUKEIRE_REC.UKEIRE_DT, vbUnicode), 7, 2)
    '�d����
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
            Exit Sub
    End Select
    Me.Fields("DET_SHIIRE_NAME") = StrConv(P_SHUKEIRE_REC.ORDER_CODE, vbUnicode) & " " & Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
    '����
    Me.Fields("DET_UKEIRE_QTY") = Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_QTY, vbUnicode)), "#,##0")
    '�P��
    Me.Fields("DET_UKEIRE_TANKA") = Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_TANKA, vbUnicode)), "#,##0.00")
    '���z
    Me.Fields("DET_UKEIRE_KINGAKU") = Format(CDbl(StrConv(P_SHUKEIRE_REC.UKEIRE_KINGAKU, vbUnicode)), "#,##0")
       
    
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
 
 

    If Trim(PR000251.Text1(ptxSHIIRE_CODE).Text) = "" Then
        Print_Mode = 0
    Else
        Print_Mode = 1
    End If

    If Print_Mode = 0 Then

        Call UniCode_Conv(K2_P_SHUKEIRE.KEIJYO_YM, Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
                                                    Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2))
        Call UniCode_Conv(K2_P_SHUKEIRE.UKEIRE_DT, "")
    
    Else

        Call UniCode_Conv(K1_P_SHUKEIRE.KEIJYO_YM, Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
                                                    Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2))
        Call UniCode_Conv(K1_P_SHUKEIRE.ORDER_CODE, PR000251.Text1(ptxSHIIRE_CODE).Text)
        Call UniCode_Conv(K1_P_SHUKEIRE.UKEIRE_DT, "")

    End If

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

