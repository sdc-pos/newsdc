VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PI00090F1 
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   24765
   _ExtentY        =   18600
   SectionData     =   "PI00090F1.dsx":0000
End
Attribute VB_Name = "PI00090F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '���ׂ�Btrieve Operation


Private DET_eof         As Integer      '���� Eof

Private DET_cnt         As Integer      '���ׂ�LINE COUNT




Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "HIN_GAI"             '�i��
    Me.Fields.Add "HIN_NAME"            '�i��
    Me.Fields.Add "ORDER_QTY"           '����
    Me.Fields.Add "Y_NOUKI_DT"          '�\��[��
    Me.Fields.Add "ORDER_NO"            '������
    Me.Fields.Add "DELI_NAME"           '�[����



End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
Dim ans         As Integer
    
    
    If DET_eof Then
        If DET_cnt > 18 Then
            Exit Sub
        End If
    End If
    
    If DET_eof Then
        Me.Fields("HIN_GAI") = ""           '�i��
        Me.Fields("HIN_NAME") = ""          '�i��
        Me.Fields("ORDER_QTY") = ""         '����
        Me.Fields("Y_NOUKI_DT") = ""        '�\��[��
        Me.Fields("ORDER_NO") = ""          '������
        Me.Fields("DELI_NAME") = ""           '�[����

    Else
        Do
            sts = BTRV(DET_com + BtSNoWait, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K2_P_SHORDER, Len(K2_P_SHORDER), 2)
            Select Case sts
                Case BtNoErr
                    If StrConv(P_SHORDER_REC.WS_NO, vbUnicode) <> StrConv(wP_SHORDER_REC.WS_NO, vbUnicode) Or _
                        StrConv(P_SHORDER_REC.PRINT_F, vbUnicode) <> P_PRINT_OFF Or _
                        StrConv(P_SHORDER_REC.ORDER_CODE, vbUnicode) <> StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode) Then
                        sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K2_P_SHORDER, Len(K2_P_SHORDER), 2)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "���ޒ����ް�")
                            Exit Sub
                        End If
                        DET_eof = True
                    End If
                    Exit Do
                Case BtErrEOF
                    
                    DET_eof = True
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Sub
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�")
                    Exit Sub
            
            End Select
        Loop
                                            
        If DET_eof Then
            Me.Fields("HIN_GAI") = ""           '�i��
            Me.Fields("HIN_NAME") = ""          '�i��
            Me.Fields("ORDER_QTY") = ""         '����
            Me.Fields("Y_NOUKI_DT") = ""        '�\��[��
            Me.Fields("ORDER_NO") = ""          '������
            Me.Fields("DELI_NAME") = ""           '�[����
                                            
                                            
                                            
        Else
            If DET_cnt > 18 Then
                Detail.NewPage = ddNPBefore
                DET_cnt = 0
            Else
                Detail.NewPage = ddNPNone
            End If
                                                '�i��
            Me.Fields("HIN_GAI") = StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)
            '�i�ڃ}�X�^�ǂݍ���
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
                                                '�i��
            Me.Fields("HIN_NAME") = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                                                '����
            Me.Fields("ORDER_QTY") = Format(CLng(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)), "#0")
                                                '�\��[��
            Me.Fields("Y_NOUKI_DT") = Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_SHORDER_REC.Y_NOUKI_DT, vbUnicode), 7, 2)
                                                '������
            Me.Fields("ORDER_NO") = StrConv(P_SHORDER_REC.ORDER_NO, vbUnicode)
            '�[����
            Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHORDER_REC.DELI_CODE, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
                    Exit Sub
            End Select
            Me.Fields("DELI_NAME") = StrConv(P_SHORDER_REC.DELI_CODE, vbUnicode) & " " & _
                                        StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
            
            
            
            
            Call UniCode_Conv(P_SHORDER_REC.PRINT_F, P_PRINT_ON)
                                                                                '�X�V����
            Call UniCode_Conv(P_SHORDER_REC.Print_datetime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
            
            
            Do
                
                DoEvents
                
                sts = BTRV(BtOpUpdate, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K2_P_SHORDER, Len(K2_P_SHORDER), 2)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SHORDER.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            sts = BTRV(BtOpUnlock, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K2_P_SHORDER, Len(K2_P_SHORDER), 2)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "���ޒ����ް�")
                                Exit Sub
                            End If
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "���ޒ����ް�")
                        Exit Sub
                End Select
            Loop
            
            Call UniCode_Conv(K2_P_SHORDER.WS_NO, StrConv(wP_SHORDER_REC.WS_NO, vbUnicode))
            Call UniCode_Conv(K2_P_SHORDER.PRINT_F, P_PRINT_OFF)
            Call UniCode_Conv(K2_P_SHORDER.ORDER_CODE, StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode))
            Call UniCode_Conv(K2_P_SHORDER.ORDER_NO, "")
            
            
            DET_com = BtOpGetGreaterEqual
    
        End If
    End If
    
    
    
    DET_cnt = DET_cnt + 1
    
            
    eof = False

End Sub

Private Sub ActiveReport_Initialize()


Dim sts     As Integer

    Field1.Text = StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode)     '�����溰��
    '�����於
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Sub
    
    End Select
    Field2.Text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))

    Field3.Text = Trim(StrConv(P_KANRIREC.KAISHA_NAME, vbUnicode))      '��Ж�
    Field4.Text = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))      '�Z���^�[��
    Field5.Text = "TEL:" & Trim(StrConv(P_KANRIREC.TEL_NO, vbUnicode))  '�Z���^�[��
    Field6.Text = "FAX:" & Trim(StrConv(P_KANRIREC.FAX_NO, vbUnicode))  '�Z���^�[��


    Print_datetime.Caption = Format(Now, "YYYY/MM/DD")


    LabBikou_1.Caption = pubBikou_1             '���l�P     2007.07.20
    LabBikou_2.Caption = pubBikou_2             '���l�Q     2007.07.20
    LabBikou_3.Caption = pubBikou_3             '���l�R     2007.07.20

    Call UniCode_Conv(K2_P_SHORDER.WS_NO, StrConv(wP_SHORDER_REC.WS_NO, vbUnicode))
    Call UniCode_Conv(K2_P_SHORDER.PRINT_F, P_PRINT_OFF)
    Call UniCode_Conv(K2_P_SHORDER.ORDER_CODE, StrConv(wP_SHORDER_REC.ORDER_CODE, vbUnicode))
    Call UniCode_Conv(K2_P_SHORDER.ORDER_NO, "")




    DET_com = BtOpGetGreaterEqual
    
    
    
    
    
    DET_eof = False
    DET_cnt = 0



End Sub

Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORPortrait
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageTopMargin = 800
    Me.PageBottomMargin = 25
    
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25

    Me.documentName = "�������F"

    DoEvents

End Sub

