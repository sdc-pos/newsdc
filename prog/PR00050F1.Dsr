VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00050F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows �̊���l
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

Private Const ptxSHIMUKE_CODE% = 0          '�d������
Private Const ptxS_YMD% = 1                 '�J�n�@�Ώ۔N����
Private Const ptxE_YMD% = 2                 '�I���@�Ώ۔N����





Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "CLASS_CODE"              '�׽����
    Me.Fields.Add "DET_NAI_CNT"             '�������Y�@����
    Me.Fields.Add "DET_NAI_SURYO"           '�������Y�@����
    Me.Fields.Add "DET_GAI_CNT"             '�O�����Y�@����
    Me.Fields.Add "DET_GAI_SURYO"           '�O�����Y�@����
    Me.Fields.Add "DET_GK_CNT"              '���v�@����
    Me.Fields.Add "DET_GK_SURYO"            '���v�@����
    Me.Fields.Add "DET_GK_TANKA"            '���v�@�P��
    Me.Fields.Add "DET_GK_KIN"              '���v�@���z

    Me.Fields.Add "DET_SHIZAI_TANKA"        '���ށ@�P��
    Me.Fields.Add "DET_SHIZAI_KIN"          '���ށ@���z
    Me.Fields.Add "DET_KOURYO_TANKA"        '�H���@�P��
    Me.Fields.Add "DET_KOURYO_KIN"          '�H���@���z
    Me.Fields.Add "DET_ETC_TANKA"           '���̑��@�P��
    Me.Fields.Add "DET_ETC_KIN"             '���̑��@���z

    Me.Fields.Add "DET_GENKA_KOSOU"         '�d�������@������
    Me.Fields.Add "DET_GENKA_GAISOU"        '�d�������@�O������
    Me.Fields.Add "DET_GENKA_KOURYO"        '�d�������@�O���H��




    
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)
    
Dim sts     As Integer
    
    sts = BTRV(DET_com, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "���Y���яW�v�ް�")
            Exit Sub
    End Select
    
    
    
    
    
    
    '�׽
    Me.Fields("CLASS_CODE") = StrConv(P_SEISAN_SUM_REC.CLASS_CODE, vbUnicode)
    '�������Y ����
    Me.Fields("DET_NAI_CNT") = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).CNT, vbUnicode)), "#0")
    '�������Y ����
    Me.Fields("DET_NAI_SURYO") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SURYO, vbUnicode)), "#0")
    '�O�����Y ����
    Me.Fields("DET_GAI_CNT") = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).CNT, vbUnicode)), "#0")
    '�O�����Y ����
    Me.Fields("DET_GAI_SURYO") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SURYO, vbUnicode)), "#0")
    '���v ����
    Me.Fields("DET_GK_CNT") = Format(CInt(Me.Fields("DET_NAI_CNT")) + CInt(Me.Fields("DET_GAI_CNT")), "#0")
    '���v ����
    Me.Fields("DET_GK_SURYO") = Format(CDbl(Me.Fields("DET_NAI_SURYO")) + CDbl(Me.Fields("DET_GAI_SURYO")), "#0")
    
    '���v �P��
    Me.Fields("DET_GK_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).TANKA, vbUnicode)), "#,##0.00")
    '���v�@���z
    Me.Fields("DET_GK_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KINGAKU, vbUnicode)), "#,##0")
    '���ށ@�P��
    Me.Fields("DET_SHIZAI_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_TANKA, vbUnicode)), "#,##0.00")
    '���ށ@���z
    Me.Fields("DET_SHIZAI_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SH_KINGAKU, vbUnicode)), "#,##0")
    
    '�H���@�P��
    Me.Fields("DET_KOURYO_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_TANKA, vbUnicode)), "#,##0.00")

    '�H���@���z
    Me.Fields("DET_KOURYO_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KO_KINGAKU, vbUnicode)), "#,##0")
    '���̑��@�P��
    Me.Fields("DET_ETC_TANKA") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_TANKA, vbUnicode)), "#,##0.00")
    '���̑��@���z
    Me.Fields("DET_ETC_KIN") = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_KINGAKU, vbUnicode)) + _
                                CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).ETC_KINGAKU, vbUnicode)), "#,##0")
    
    
    '�d�������@��
    Me.Fields("DET_GENKA_KOSOU") = Format(CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)), "#,##0")
    '�d�������@�O��
    Me.Fields("DET_GENKA_GAISOU") = Format(CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)), "#,##0")
    '�d�������@�H��
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
        
            MsgBox "�Ώ��ް�������܂���"
            Exit Sub
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���Y���яW�v�ް�")
            Exit Sub
    End Select


                                            '�������Y�@���Y����
    NAI_CNT.Text = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).CNT, vbUnicode)), "#,##0")
                                            '�O�����Y�@���Y����
    GAI_CNT.Text = Format(CInt(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).CNT, vbUnicode)), "#,##0")
                                            '���v�@ �@ ���Y����
    GK_CNT.Text = Format(CInt(NAI_CNT.Text) + CInt(GAI_CNT.Text), "#,##0")
                                            
                                            
                                            '�������Y�@���Y����
    NAI_SURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SURYO, vbUnicode)), "#,##0")
                                            '�O�����Y�@���Y����
    GAI_SURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SURYO, vbUnicode)), "#,##0")
                                            '���v�@ �@ ���Y����
    GK_SURYO.Text = Format(CDbl(NAI_SURYO.Text) + CDbl(GAI_SURYO.Text), "#,##0")
                                            
                                            
                                            
                                            '�������Y  �\���䗦
    wkValue = CDbl(NAI_SURYO.Text) / (CDbl(NAI_SURYO.Text) + CDbl(GAI_SURYO.Text)) * 100
    NAI_SU_RITU.Text = Format(wkValue, "#0.00") & "%"
                                            
                                            '�O�����Y  �\���䗦
    GAI_SU_RITU.Text = Format(100 - wkValue, "#0.00") & "%"
                                            '�\��  �\���䗦
    GK_SU_RITU.Text = "100.00%"
                                            
                                            
                                            
                                            '�������Y�@���Y���z
    NAI_KIN.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KINGAKU, vbUnicode)), "#,##0")
                                            '�O�����Y�@���Y���z
    GAI_KIN.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KINGAKU, vbUnicode)), "#,##0")
                                            '���v�@���Y���z
    GK_KIN.Text = Format(CDbl(NAI_KIN.Text) + CDbl(GAI_KIN.Text), "#,##0")
        
                                            
                                            '�������Y  �\���䗦
    If CDbl(NAI_KIN.Text) + CDbl(GAI_KIN.Text) = 0 Then
        wkValue = 0
    Else
        wkValue = CDbl(NAI_KIN.Text) / (CDbl(NAI_KIN.Text) + CDbl(GAI_KIN.Text)) * 100
    End If
    NAI_KIN_RITU.Text = Format(wkValue, "#0.00") & "%"
                                            
                                            '�O�����Y  �\���䗦
    GAI_KIN_RITU.Text = Format(100 - wkValue, "#0.00") & "%"
                                            '�\��  �\���䗦
    GK_KIN_RITU.Text = "100.00%"
        
                                            
                                            '�������Y�@����@����
    NAI_UCHI_SHIZAI.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).SH_KINGAKU, vbUnicode)), "#,##0")
                                            '�������Y�@����@�H��
    NAI_UCHI_KOURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).KO_KINGAKU, vbUnicode)), "#,##0")
                                            '�������Y�@����@���̑�
    NAI_UCHI_ETC.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(0).ETC_KINGAKU, vbUnicode)), "#,##0")
                                            '�O�����Y�@����@����
    GAI_UCHI_SHIZAI.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).SH_KINGAKU, vbUnicode)), "#,##0")
                                            '�O�����Y�@����@�H��
    GAI_UCHI_KOURYO.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).KO_KINGAKU, vbUnicode)), "#,##0")
                                            '�O�����Y�@����@���̑�
    GAI_UCHI_ETC.Text = Format(CDbl(StrConv(P_SEISAN_SUM_REC.UCHIWAKE(1).ETC_KINGAKU, vbUnicode)), "#,##0")
                                            '���v�@����@����
    GK_UCHI_SHIZAI.Text = Format(CDbl(NAI_UCHI_SHIZAI.Text) + CDbl(GAI_UCHI_SHIZAI.Text), "#,##0")
                                            '���v�@����@�H��
    GK_UCHI_KOURYO.Text = Format(CDbl(NAI_UCHI_KOURYO.Text) + CDbl(GAI_UCHI_KOURYO.Text), "#,##0")
                                            '���v�@����@���̑�
    GK_UCHI_ETC.Text = Format(CDbl(NAI_UCHI_ETC.Text) + CDbl(GAI_UCHI_ETC.Text), "#,##0")
                                            
                                            
                                            '���i�\����
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
                                            '�����
    GENKA_KOSOU.Text = Format(CLng(StrConv(P_SEISAN_SUM_REC.KO_GENKA, vbUnicode)), "#,##0")
    GENKA_GAISOU.Text = Format(CLng(StrConv(P_SEISAN_SUM_REC.GA_GENKA, vbUnicode)), "#,##0")
    GENKA_SHIZAI.Text = Format(CDbl(GENKA_KOSOU.Text) + CDbl(GENKA_GAISOU.Text), "#,##0")
    GENKA_KOURYO.Text = Format(CLng(StrConv(P_SEISAN_SUM_REC.GK_GENKA, vbUnicode)), "#,##0")
    GENKA_GK.Text = Format(CDbl(GENKA_KOURYO.Text) + CDbl(GENKA_SHIZAI.Text), "#,##0")
    

    





    sts = BTRV(BtOpDelete, P_SEISAN_SUM_POS, P_SEISAN_SUM_REC, Len(P_SEISAN_SUM_REC), K0_P_SEISAN_SUM, Len(K0_P_SEISAN_SUM), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            Call File_Error(sts, BtOpDelete, "���Y���яW�v�ް�")
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

    Me.documentName = "���i�����Y���яW�v�\�F"

    DoEvents

End Sub

