VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} ODR3010F2 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "ODR30101F2.dsx":0000
End
Attribute VB_Name = "ODR3010F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DET_com     As Integer
Dim W_CNT       As Long
Dim P_Cnt       As Long
Dim sw          As Integer

Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "DET_HIN_GAI"             '�q���i�R�[�h
    Me.Fields.Add "DET_HIN_NAME"            '�i��
    Me.Fields.Add "DET_SO_SUU"              '�q���i�K�v��
    
    Me.Fields.Add "DET_ZAIKO_QTY"           '�݌ɐ�
    
    Me.Fields.Add "DET_SHIJI_Z_QTY"         '�����c
    
    Me.Fields.Add "DET_HIKIATE_Z_QTY"       '�����c
    
    Me.Fields.Add "DET_FUSOKU_QTY"          '�s����
    
    Me.Fields.Add "DET_ORDER_QTY"           '������
    
    Me.Fields.Add "DET_LOT"                 '����ۯ�
    
    Me.Fields.Add "DET_ORDER_CODE"          '�d���溰��
    
    Me.Fields.Add "DET_ORDER_NAME"          '�d���於��
    
    Me.Fields.Add "DET_LT"                  'ذ�����
    
    Me.Fields.Add "DET_Y_NOUKI_DT"          '�\��[��
    
    W_CNT = 0
    P_Cnt = 0
End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
Dim wkAVE       As Long
Dim i           As Integer

    Do
        sw = 1
        sts = BTRV(DET_com, ODR_KNT_POS, ODR_KNT_R, Len(ODR_KNT_R), K0_ODR_KENTO, Len(K0_ODR_KENTO), 0)
        Select Case sts
            Case BtNoErr
            
            
            Case BtErrEOF
                Exit Sub
            Case Else
                Call File_Error(sts, DET_com, "��������̧��3")
                Exit Sub
        End Select
        
        
        '   �g�p���̔���I
        'If StrConv(ODR_KNT_R.USE_YM, vbUnicode) > Key_USE_YM Then Exit Sub
        'If StrConv(ODR_KNT_R.USE_YM, vbUnicode) > Key_USE_YM Then Exit Do
        
        '   �K�v�������O���ΏہI�H
        'If CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode))) <> 0 Then
        If CDbl(Trim(StrConv(ODR_KNT_R.NED_QTY, vbUnicode))) = 0 Then sw = 0
        
        '   �s�������O���ΏہI�H
        'If CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode))) < 0 Then
        If CDbl(Trim(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode))) >= 0 Then sw = 0
        
        If sw = 1 Then
            Exit Do
        End If
        DET_com = BtOpGetNext
    Loop
    
    
    
    If StrConv(ODR_KNT_R.USE_YM, vbUnicode) > Key_USE_YM Then Exit Sub
    
    W_CNT = W_CNT + 1
    '�i��(�O��)
    Me.Fields("DET_HIN_GAI") = Trim(StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode))
    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_KNT_R.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_KNT_R.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_KNT_R.KO_HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        
            Me.Fields("DET_HIN_NAME") = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        
        Case BtErrKeyNotFound
            Me.Fields("DET_HIN_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
            Exit Sub
    End Select
    '�q���i�K�v��
    Me.Fields("DET_SO_SUU") = Format(CLng(StrConv(ODR_KNT_R.NED_QTY, vbUnicode)), "#,##0")
    '�݌ɐ�
    Me.Fields("DET_ZAIKO_QTY") = Format(CLng(StrConv(ODR_KNT_R.ZAI_QTY, vbUnicode)), "#,##0")
    '�����c
    Me.Fields("DET_SHIJI_Z_QTY") = Format(CLng(StrConv(ODR_KNT_R.ODR_QTY, vbUnicode)), "#,##0")
    '�s����
    Me.Fields("DET_FUSOKU_QTY") = Format(CLng(StrConv(ODR_KNT_R.MAI_QTY, vbUnicode)), "#,##0")
    '������
    Me.Fields("DET_ORDER_QTY") = Format(CLng(StrConv(ODR_KNT_R.ODR_QTY, vbUnicode)), "#,###")
    '����ۯ�
    Me.Fields("DET_LOT") = Format(CLng(StrConv(ODR_KNT_R.LOT_QTY, vbUnicode)), "#,##0")
    '�d���溰��
    Me.Fields("DET_ORDER_CODE") = Trim(StrConv(ODR_KNT_R.SECT, vbUnicode))
    '�d���於
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(ODR_KNT_R.SECT, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
        
            Me.Fields("DET_ORDER_NAME") = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
        
        Case BtErrKeyNotFound
            Me.Fields("DET_ORDER_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�Ͻ�")
            Exit Sub
    End Select
    '���[�h�^�C��
    Me.Fields("DET_LT") = ""
    '�[���\���
    If Trim(StrConv(ODR_KNT_R.NOUKI, vbUnicode)) = "" Then
        Me.Fields("DET_Y_NOUKI_DT") = ""
    Else
        Me.Fields("DET_Y_NOUKI_DT") = Left(StrConv(ODR_KNT_R.NOUKI, vbUnicode), 4) & "/" & _
                                Mid(StrConv(ODR_KNT_R.NOUKI, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_KNT_R.NOUKI, vbUnicode), 2)
    End If
       
    
    
    DET_com = BtOpGetNext

    eof = False
End Sub

Private Sub ActiveReport_Initialize()



    DET_com = BtOpGetGreater



    Call UniCode_Conv(K0_ODR_KENTO.USE_YM, Key_USE_YM)
    
    Call UniCode_Conv(K0_ODR_KENTO.KO_JGYOBU, "")
    Call UniCode_Conv(K0_ODR_KENTO.KO_NAIGAI, "")
    Call UniCode_Conv(K0_ODR_KENTO.KO_HIN_GAI, "")




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
    
    USE_YM.Text = Left(Key_USE_YM, 4) & "/" & Right(Key_USE_YM, 2)
    
End Sub

Private Sub PageHeader_Format()
    
    Page_Cnt.Text = Format(ODR3010F2.pageNumber, "##0")
    
    P_Cnt = P_Cnt + 1
    Page_Cnt.Text = Format(P_Cnt, "##0")
End Sub
