VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00020F1 
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   17160
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   30268
   _ExtentY        =   17621
   SectionData     =   "PR00020F1.dsx":0000
End
Attribute VB_Name = "PR00020F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '���ׂ�Btrieve Operation


'�v��N���p�Y��
Private Const ptxS_ORDER_DT% = 0            '�������@�J�n
Private Const ptxE_ORDER_DT% = 1            '�������@�I��

Private Const ptxS_Y_NOUKI_DT% = 2          '��]�[�� �J�n�@2008.01.10
Private Const ptxE_Y_NOUKI_DT% = 3          '��]�[�� �J�n�@2008.01.10

Private Const ptxUSE_YM% = 4                '�g�p��         2008.01.10




Private Const ptxORDER_CODE% = 5            '�����溰��






Private Sub ActiveReport_DataInitialize()

''2007.10.31    Me.Fields.Add "DET_CANCEL"          '��ݾ�  2007.07.26
    Me.Fields.Add "DET_ORDER_DT"        '������
    Me.Fields.Add "DET_ORDER"           '������
    Me.Fields.Add "DET_HIN_GAI"         '�i��
    Me.Fields.Add "DET_HIN_NAME"        '�i��
    Me.Fields.Add "DET_ORDER_QTY"       '������
    Me.Fields.Add "DET_ZAN_QTY"         '�����c
    Me.Fields.Add "DET_ZAIKO_QTY"       '�݌ɐ�
    Me.Fields.Add "DET_Y_NOUKI_DT"      '�\��[��


    Me.Fields.Add "DET_ANS_NOUKI_DT"    '�񓚔[��   2008.01.10
    Me.Fields.Add "DET_USE_YM"          '�g�p��     2008.01.10

    Me.Fields.Add "DET_ORDER_NO"        '������     2012.02.28
    Me.Fields.Add "DET_ORDER_NO_BC"     '������(BC) 2012.02.28

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
Dim Mi_QTY      As Long
Dim Sumi_QTY    As Long
Dim SKIP_Flg    As Boolean
    
Dim work        As String       '2012.03.01
Dim i           As Integer      '2012.03.01
    
    
    sts = BTRV(DET_com, tmpP_SHORDER_POS, tmpP_SHORDER_REC, Len(tmpP_SHORDER_REC), K4_tmpP_SHORDER, Len(K4_tmpP_SHORDER), 4)
'    sts = BTRV(DET_com, tmpP_SHORDER_POS, tmpP_SHORDER_REC, Len(tmpP_SHORDER_REC), K3_tmpP_SHORDER, Len(K3_tmpP_SHORDER), 3)
    Select Case sts
        Case BtNoErr
        
            If StrConv(tmpP_SHORDER_REC.KAN_F, vbUnicode) <> P_KAN_OFF Then
                '�����׸�
                Exit Sub
            End If
            
            If Trim(PR000201.Text1(ptxE_ORDER_DT).Text) <> "" Then
                If StrConv(tmpP_SHORDER_REC.ORDER_DT, vbUnicode) > Mid(PR000201.Text1(ptxE_ORDER_DT).Text, 1, 4) & _
                                                                    Mid(PR000201.Text1(ptxE_ORDER_DT).Text, 6, 2) & _
                                                                    Mid(PR000201.Text1(ptxE_ORDER_DT).Text, 9, 2) Then
                    '�������͈̔�
                    Exit Sub
                End If
            End If
            If Trim(PR000201.Text1(ptxORDER_CODE).Text) <> "" Then
                If Trim(StrConv(tmpP_SHORDER_REC.ORDER_CODE, vbUnicode)) <> Trim(PR000201.Text1(ptxORDER_CODE).Text) Then
                    Exit Sub
                End If
            End If
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "���ޒ����ް�")
            Exit Sub
    End Select
    
    
    If Not SKIP_Flg Then
        '��ݾ�  2007.07.26
''2007.10.31        If StrConv(tmpP_SHORDER_REC.CANCEL_F, vbUnicode) = P_CANCEL_ON Then
''2007.10.31            Me.Fields("DET_CANCEL") = "*"
''2007.10.31        Else
''2007.10.31            Me.Fields("DET_CANCEL") = " "
''2007.10.31        End If
        
        '������
        Me.Fields("DET_ORDER_DT") = Mid(StrConv(tmpP_SHORDER_REC.ORDER_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(tmpP_SHORDER_REC.ORDER_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(tmpP_SHORDER_REC.ORDER_DT, vbUnicode), 7, 2)
        '������
        Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(tmpP_SHORDER_REC.ORDER_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
                Exit Sub
        End Select
        
        '-------------------------- 2012.03.01
        For i = 1 To Len(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
        
            work = Left(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode), i)    '�P�����C�Q�����E�E�E�Ə��ԂɏZ����؂�o��
            work = StrConv(work, vbFromUnicode)                                 'Unicode����V�X�e���̊���̃R�[�h�ɕϊ�
            If (LenB(work) > 19) Then                                           '�Q�O�o�C�g�𒴂�����d�w�h�s
                Exit For                                                        '�d�w�h�s�����address2�̐擪�����ʒu��ii�Ɋi�[�����
            End If
        Next i
        work = Left(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode), i - 1)
        '-------------------------- 2012.03.01
        
        
        
        Me.Fields("DET_ORDER") = StrConv(tmpP_SHORDER_REC.ORDER_CODE, vbUnicode) & " " & work
        '�i��
        Me.Fields("DET_HIN_GAI") = Trim(StrConv(tmpP_SHORDER_REC.HIN_GAI, vbUnicode))
        '�i��
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(tmpP_SHORDER_REC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(tmpP_SHORDER_REC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(tmpP_SHORDER_REC.HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Sub
        End Select
        Me.Fields("DET_HIN_NAME") = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        '������
        Me.Fields("DET_ORDER_QTY") = Format(CLng(StrConv(tmpP_SHORDER_REC.ORDER_QTY, vbUnicode)), "#,##0")
        '�����c
        Me.Fields("DET_ZAN_QTY") = Format(CLng(StrConv(tmpP_SHORDER_REC.ORDER_QTY, vbUnicode)) - CLng(StrConv(tmpP_SHORDER_REC.UKEIRE_QTY, vbUnicode)), "#,##0")
        '�݌ɐ�
        '���݌�
        If Zaiko_Syukei_Proc(Sumi_QTY, Mi_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
            Exit Sub
        End If
        Me.Fields("DET_ZAIKO_QTY") = Format(Mi_QTY + Sumi_QTY, "#,##0")
        '�\��[��
        Me.Fields("DET_Y_NOUKI_DT") = Mid(StrConv(tmpP_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(tmpP_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(tmpP_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
    
        '�񓚔[���� 2008.01.10
        
        If OSAKA_MODE Then
        
            Me.Fields("DET_ANS_NOUKI_DT") = Mid(StrConv(tmpP_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(tmpP_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(tmpP_SHORDER_REC.ANS_NOUKI_DT, vbUnicode), 7, 2)
        Else
            Me.Fields("DET_ANS_NOUKI_DT") = ""
        End If
        
        '�g�p�� 2008.01.10
        If OSAKA_MODE Then
            Me.Fields("DET_USE_YM") = Mid(StrConv(tmpP_SHORDER_REC.USE_YM, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(tmpP_SHORDER_REC.USE_YM, vbUnicode), 5, 2)
        Else
            Me.Fields("DET_USE_YM") = ""
        End If
    
    
    
    
        '������ 2012.02.28
        Me.Fields("DET_ORDER_NO") = StrConv(tmpP_SHORDER_REC.ORDER_NO, vbUnicode)
        '������(BC) 2012.02.28
        Me.Fields("DET_ORDER_NO_BC") = "*" & StrConv(tmpP_SHORDER_REC.ORDER_NO, vbUnicode) & "*"
    
    
    
    End If
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
 
 

'2007.10.31    Call UniCode_Conv(K_tmpP_SHORDER.KAN_F, P_KAN_OFF)
'2007.10.31    Call UniCode_Conv(K4_tmpP_SHORDER.ORDER_CODE, PR000201.Text1(ptxORDER_CODE).Text)
'2007.10.31    Call UniCode_Conv(K4_tmpP_SHORDER.ORDER_DT, Mid(PR000201.Text1(ptxS_ORDER_DT).Text, 1, 4) & _
'                                                                Mid(PR000201.Text1(ptxS_ORDER_DT).Text, 6, 2) & _
'                                                                Mid(PR000201.Text1(ptxS_ORDER_DT).Text, 9, 2))
'

'2007.10.31    DET_com = BtOpGetGreaterEqual
 
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

    Me.documentName = "�d���c�ꗗ�\�F"



    Line182.Visible = OSAKA_MODE    '2008.01.10
    Line191.Visible = OSAKA_MODE    '2008.01.10
    Line211.Visible = OSAKA_MODE    '2008.01.10
    Line212.Visible = OSAKA_MODE    '2008.01.10
    Line213.Visible = OSAKA_MODE    '2008.01.10
    Line214.Visible = OSAKA_MODE    '2008.01.10
    Line215.Visible = OSAKA_MODE    '2008.01.10


    Label45.Visible = OSAKA_MODE    '2008.01.10
    Label46.Visible = OSAKA_MODE    '2008.01.10

    DoEvents

End Sub

