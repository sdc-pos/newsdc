VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00070F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16875
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   29766
   _ExtentY        =   19420
   SectionData     =   "PR00070F1.dsx":0000
End
Attribute VB_Name = "PR00070F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DET_com     As Integer


Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "DET_HIN_GAI"             '�i�ں���
    Me.Fields.Add "DET_HIN_NAME"            '�i��
    Me.Fields.Add "DET_JITU_QTY1"           '���с@�O�X��
    Me.Fields.Add "DET_JITU_QTY2"           '���с@�O��
    Me.Fields.Add "DET_JITU_QTY3"           '���с@����

    Me.Fields.Add "DET_JITU_AVE"            '���с@������

    Me.Fields.Add "DET_LT_DAYS"             'LT
    
    Me.Fields.Add "DET_SYUSHI_CODE"         '���x
    
    Me.Fields.Add "DET_Zaiko_Standard"      '��݌�
    Me.Fields.Add "DET_ZAIKO_QTY"           '���_�݌�
    
    Me.Fields.Add "DET_LOT"                 '����ۯ�
    Me.Fields.Add "DET_ORDER_CODE"       '������
    
    Me.Fields.Add "DET_SHIJI_Z_QTY"         '�����c
    
    Me.Fields.Add "DET_SHIJI_QTY_R"        '�������@���_
    Me.Fields.Add "DET_SHIJI_QTY_K"        '�������@�m��
    
    Me.Fields.Add "DET_TANKA"               '�P��
    Me.Fields.Add "DET_KINGAKU"             '���z
    

End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
Dim wkAVE       As Long
Dim i           As Integer

    sts = BTRV(DET_com, P_SHKENTO_POS, P_SHKENTO_REC, Len(P_SHKENTO_REC), K1_P_SHKENTO, Len(K1_P_SHKENTO), 1)
    Select Case sts
        Case BtNoErr
        
        
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "���ޔ�������̧��")
            Exit Sub
    End Select

    If DET_com = BtOpGetLast Then
        JITU_YM1.Caption = StrConv(P_SHKENTO_REC.JITU_TBL(2).JITU_YM, vbUnicode)
        JITU_YM2.Caption = StrConv(P_SHKENTO_REC.JITU_TBL(1).JITU_YM, vbUnicode)
        JITU_YM3.Caption = StrConv(P_SHKENTO_REC.JITU_TBL(0).JITU_YM, vbUnicode)
    End If


    '�i��(�O��)
    Me.Fields("DET_HIN_GAI") = Trim(StrConv(P_SHKENTO_REC.HIN_GAI, vbUnicode))
    '�i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SHKENTO_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SHKENTO_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SHKENTO_REC.HIN_GAI, vbUnicode))
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
    '����
    Me.Fields("DET_JITU_QTY1") = Format(CLng(StrConv(P_SHKENTO_REC.JITU_TBL(2).JITU_QTY, vbUnicode)), "#,##0")
    Me.Fields("DET_JITU_QTY2") = Format(CLng(StrConv(P_SHKENTO_REC.JITU_TBL(1).JITU_QTY, vbUnicode)), "#,##0")
    Me.Fields("DET_JITU_QTY3") = Format(CLng(StrConv(P_SHKENTO_REC.JITU_TBL(0).JITU_QTY, vbUnicode)), "#,##0")
    
    wkAVE = 0
    For i = 0 To 2
        wkAVE = wkAVE + CLng(StrConv(P_SHKENTO_REC.JITU_TBL(i).JITU_QTY, vbUnicode))
    Next i
    
    Me.Fields("DET_JITU_AVE") = Format(Round(wkAVE / 3, 1), "#,##0.0")
    'LT
    If IsNumeric(StrConv(P_SHKENTO_REC.LT_DAYS, vbUnicode)) Then
        Me.Fields("DET_LT_DAYS") = Format(CInt(StrConv(P_SHKENTO_REC.LT_DAYS, vbUnicode)), "##0")
    Else
        Me.Fields("DET_LT_DAYS") = 0
    End If
    '���x
    Me.Fields("DET_SYUSHI_CODE") = StrConv(P_SHKENTO_REC.SYUSHI_CODE, vbUnicode)
    '��݌�
    If IsNumeric(StrConv(P_SHKENTO_REC.ZAIKO_STANDARD, vbUnicode)) Then
        Me.Fields("DET_ZAIKO_STANDARD") = Format(CLng(StrConv(P_SHKENTO_REC.ZAIKO_STANDARD, vbUnicode)), "#,##0")
    Else
        Me.Fields("DET_ZAIKO_STANDARD") = 0
    End If
    '���_�݌�
    Me.Fields("DET_ZAIKO_QTY") = Format(CLng(StrConv(P_SHKENTO_REC.ZAIKO_QTY, vbUnicode)), "#,##0")
    'ۯ�
    If IsNumeric(StrConv(P_SHKENTO_REC.LOT, vbUnicode)) Then
        Me.Fields("DET_LOT") = Format(CLng(StrConv(P_SHKENTO_REC.LOT, vbUnicode)), "#,##0")
    Else
        Me.Fields("DET_LOT") = 0
    End If
    '������
    Me.Fields("DET_ORDER_CODE") = StrConv(P_SHKENTO_REC.ORDER_CODE, vbUnicode)
    '�����c�@����
    Me.Fields("DET_SHIJI_Z_QTY") = Format(CLng(StrConv(P_SHKENTO_REC.SHIJI_Z_QTY, vbUnicode)), "#,##0")
    '�������@���_
    Me.Fields("DET_SHIJI_QTY_R") = Format(CLng(StrConv(P_SHKENTO_REC.SHIJI_QTY_R, vbUnicode)), "#,##0")
    '�������@�m��
    Me.Fields("DET_SHIJI_QTY_K") = Format(CLng(StrConv(P_SHKENTO_REC.SHIJI_QTY_K, vbUnicode)), "#,##0")
    '�P��
    Me.Fields("DET_TANKA") = Format(CLng(StrConv(P_SHKENTO_REC.TANKA, vbUnicode)), "#,##0.0")
    '���z
    Me.Fields("DET_KINGAKU") = Format(CLng(StrConv(P_SHKENTO_REC.KINGAKU, vbUnicode)), "#,##0.0")


    DET_com = BtOpGetPrev

    eof = False
End Sub

Private Sub ActiveReport_Initialize()


    S_YY.Text = Mid(GLB_S_YMD, 1, 4)
    S_MM.Text = Mid(GLB_S_YMD, 6, 2)
    S_DD.Text = Mid(GLB_S_YMD, 9, 2)

    E_YY.Text = Mid(GLB_E_YMD, 1, 4)
    E_MM.Text = Mid(GLB_E_YMD, 6, 2)
    E_DD.Text = Mid(GLB_E_YMD, 9, 2)

    
    PRI_YY.Text = Mid(Format(Now, "YYYYMMDD"), 1, 4)
    PRI_MM.Text = Mid(Format(Now, "YYYYMMDD"), 5, 2)
    PRI_DD.Text = Mid(Format(Now, "YYYYMMDD"), 7, 2)

    PRI_CENTER.Text = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))

    TOTAL_KINGAKU.Caption = Format(GLB_TOTAL_KINGAKU, "#,##0")

    DET_com = BtOpGetLast

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

End Sub

