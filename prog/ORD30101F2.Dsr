VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} ORD3010F2 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16845
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   29713
   _ExtentY        =   19420
   SectionData     =   "ORD30101F2.dsx":0000
End
Attribute VB_Name = "ORD3010F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DET_com     As Integer


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
    

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
Dim wkAVE       As Long
Dim i           As Integer

    sts = BTRV(DET_com, ODR_TP3_POS, ODR_TP3_R, Len(ODR_TP3_R), K0_ODR_TEMP3, Len(K0_ODR_TEMP3), 0)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "��������̧��3")
            Exit Sub
    End Select

    '�i��(�O��)
    Me.Fields("DET_HIN_GAI") = Trim(StrConv(ODR_TP3_R.KO_HIN_GAI, vbUnicode))
    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ODR_TP3_R.KO_JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ODR_TP3_R.KO_NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ODR_TP3_R.KO_HIN_GAI, vbUnicode))
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
    Me.Fields("DET_SO_SUU") = Format(CLng(StrConv(ODR_TP3_R.REQ_QTY, vbUnicode)), "#,##0")
    '�݌ɐ�
    Me.Fields("DET_ZAIKO_QTY") = Format(CLng(StrConv(ODR_TP3_R.ZAI_QTY, vbUnicode)), "#,##0")
    '�����c
    Me.Fields("DET_SHIJI_Z_QTY") = Format(CLng(StrConv(ODR_TP3_R.ODR_QTY, vbUnicode)), "#,##0")
    '�s����
    Me.Fields("DET_FUSOKU_QTY") = Format(CLng(StrConv(ODR_TP3_R.MAI_QTY, vbUnicode)), "#,##0")
    '������
    Me.Fields("DET_ORDER_QTY") = Format(CLng(StrConv(ODR_TP3_R.ODR_QTY, vbUnicode)), "#,###")
    '����ۯ�
    Me.Fields("DET_LOT") = Format(CLng(StrConv(ODR_TP3_R.LOT_QTY, vbUnicode)), "#,##0")
    '�d���溰��
    Me.Fields("DET_ORDER_CODE") = Trim(StrConv(ODR_TP3_R.SECT, vbUnicode))
    '�d���於
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(ODR_TP3_R.SECT, vbUnicode))
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
    '�[���\���
    If Trim(StrConv(ODR_TP3_R.NOUKI, vbUnicode)) = "" Then
        Me.Fields("DET_Y_NOUKI_DT") = ""
    Else
        Me.Fields("DET_Y_NOUKI_DT") = Left(StrConv(ODR_TP3_R.NOUKI, vbUnicode), 4) & "/" & _
                                Mid(StrConv(ODR_TP3_R.NOUKI, vbUnicode), 5, 2) & "/" & _
                                Right(StrConv(ODR_TP3_R.NOUKI, vbUnicode), 2)
    End If

    DET_com = BtOpGetNext

    eof = False
End Sub

Private Sub ActiveReport_Initialize()



    DET_com = BtOpGetGreater



    Call UniCode_Conv(K0_ODR_TEMP3.USE_YM, Key_USE_YM)
    
    Call UniCode_Conv(K0_ODR_TEMP3.KO_JGYOBU, "")
    Call UniCode_Conv(K0_ODR_TEMP3.KO_NAIGAI, "")
    Call UniCode_Conv(K0_ODR_TEMP3.KO_HIN_GAI, "")




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
    Page_Cnt.Text = Format(ORD3010F2.pageNumber, "#0")
End Sub
