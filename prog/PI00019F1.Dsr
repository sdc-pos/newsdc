VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PI00019F1 
   ClientHeight    =   14805
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   19050
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   33602
   _ExtentY        =   26114
   SectionData     =   "PI00019F1.dsx":0000
End
Attribute VB_Name = "PI00019F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Doukon_com      As Integer      '�\���^������Btrieve Operation
Private Doukon_eof      As Integer      '�\���^���� Eof

Private Doukon_cnt      As Integer      '�\���^������LINE COUNT

Private EOF_F           As Boolean      '2012.04.17


Private SHIJI_QTY       As Double       '����w����


Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "KO_NO"               'No
    Me.Fields.Add "KO_HIN_GAI"          '�i��
    Me.Fields.Add "KO_SYUBETSU"         '���
    Me.Fields.Add "KO_QTY"              '����
    Me.Fields.Add "KO_SHIJI_QTY"        '����

    Me.Fields.Add "KO_ST_LOCATION"      '�I��
    Me.Fields.Add "KO_ZAIKO_QTY"        '���_�݌�
    Me.Fields.Add "KO_ID_NO"            'ID_NO
    Me.Fields.Add "KO_ID_BCR"           'ID_NO�ް����
    Me.Fields.Add "KO_BIKOU"            '���l



End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
    
Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
    
Dim SURYO       As String

Dim ST_SOKO     As String
Dim c           As String * 128
    
    If Doukon_cnt > 19 Then
        
        If Doukon_eof Then
            Exit Sub
        Else
            Doukon_cnt = 0
        End If
    End If
    
    Me.Fields("ko_no").Value = Doukon_Tbl_No(Doukon_cnt)
    
    If Doukon_eof Then
        Me.Fields("KO_HIN_GAI") = " "       '�i��
        Me.Fields("KO_SYUBETSU") = " "      '���
        Me.Fields("KO_QTY") = " "           '����
        Me.Fields("KO_SHIJI_QTY") = " "     '����
        Me.Fields("KO_ST_LOCATION") = " "   '�I��
        Me.Fields("KO_ZAIKO_QTY") = " "     '���_�݌�
        Me.Fields("KO_ID_NO") = " "         'ID_NO
        Me.Fields("KO_ID_BCR") = " "        'ID_NO�ް����
        Me.Fields("KO_BIKOU") = " "         '���l
    
        
    Else
        
        
        sts = BTRV(Doukon_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Taget_Key Or _
                    StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                    Doukon_eof = True
                    EOF_F = True                '2012.04.17
                
                End If
            
            
                If Doukon_cnt = 0 Then              '2012.04.17
                    If Doukon_eof Then              '2012.04.17
                        Doukon_cnt = Doukon_cnt + 1 '2012.04.17
'                        eof = False                '2012.04.17
                        Exit Sub                   '2012.04.17
                    End If                          '2012.04.17
                End If                              '2012.04.17
            
            
            Case BtErrEOF
                
                Doukon_eof = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�Ώێw�}�[�ް��i�e�j")
                Exit Sub
        
        End Select
                                            
                                            
                                            
        If Doukon_eof Then
            Me.Fields("KO_HIN_GAI") = " "        '�i��
            Me.Fields("KO_SYUBETSU") = " "       '���
            Me.Fields("KO_QTY") = " "            '����
            Me.Fields("KO_SHIJI_QTY") = " "      '����
            Me.Fields("KO_ST_LOCATION") = " "    '�I��
            Me.Fields("KO_ZAIKO_QTY") = " "      '���_�݌�
            Me.Fields("KO_ID_NO") = " "          'ID_NO
            Me.Fields("KO_ID_BCR") = " "         'ID_NO�ް����
            Me.Fields("KO_BIKOU") = " "          '���l
                                            
                                            
                                            
        Else
                                                '�i��
            Me.Fields("KO_HIN_GAI") = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                                                '���
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                    Exit Sub
            
            End Select
            Me.Fields("KO_SYUBETSU") = StrConv(P_CODEREC.C_RNAME, vbUnicode)
                                                '����
            If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
            Else
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
            End If
                                                '����
'            If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'            Else
'                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'            End If
        
            SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
            If CInt(Right(SURYO, 2)) = 0 Then
                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(SURYO), "#0")
            Else
                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(SURYO), "#0.00")
            End If
        
        
        
            '�i�ڃ}�X�^�ǂݍ���
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    
                    
                    Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Sub
    
            End Select
        
        
        
            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Me.Fields("KO_ST_LOCATION") = ""
            Else
                '�W���I��
                
                ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                'P_SYS.INI--> PI00010.INI   2011.08.04
                If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                Else
                    ST_SOKO = RTrim(c)
                End If
                
                
                
                Me.Fields("KO_ST_LOCATION") = Trim(ST_SOKO) & "-" & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If
        
        
            '�݌ɐ�
            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Sub
            
            End If
            Me.Fields("KO_ZAIKO_QTY") = Format(Sumi_Qty + Mi_Qty, "#0")
            '���lOR�o���ް���h
            If PRI_BIKOU_BCR Then
                                                                                        
                                                                                        
                If Trim(StrConv(P_SSHIJI_K_REC.KO_ID_NO, vbUnicode)) = "" Then
                                                                                        'ID_NO
                    Me.Fields("KO_ID_NO") = ""
                                                                                        'ID_NO�ް����
                    Me.Fields("KO_ID_BCR") = ""
                Else
                                                                                        'ID_NO
                    Me.Fields("KO_ID_NO") = StrConv(ITEMREC.JGYOBU, vbUnicode) & StrConv(P_SSHIJI_K_REC.KO_ID_NO, vbUnicode)
                                                                                            'ID_NO�ް����
                    Me.Fields("KO_ID_BCR") = "*" & StrConv(ITEMREC.JGYOBU, vbUnicode) & StrConv(P_SSHIJI_K_REC.KO_ID_NO, vbUnicode) & "*"
                End If
            Else
                Me.Fields("KO_BIKOU") = StrConv(P_SSHIJI_K_REC.KO_BIKOU, vbUnicode)     '���l
        
            End If
        End If
            
    
    
    
    
    
    
        Doukon_com = BtOpGetNext
    End If
    
    
    
    Doukon_cnt = Doukon_cnt + 1
    
            
    eof = False
    
    
    


End Sub

Private Sub ActiveReport_Initialize()

Dim sts             As Integer

Dim cnt             As Integer
Dim com             As Integer


Dim i               As Integer
Dim Total_Times     As Double
Dim AVE             As Double


Dim SURYO           As String

Dim ST_SOKO         As String
Dim c               As String * 128

Dim Target          As Double

Dim wkValue         As String
Dim wkEDIT_NIN      As String
Dim wkEDIT_TIMES    As String
Dim wkAVE           As String


Dim Tanka_F         As Boolean      '2008.09.20
Dim wkDate          As String * 8   '2008.09.20

Dim wkTOTAL         As Double       '2008.09.20

Dim wkTarget        As String       '2008.09.20




    '�Ώێw�}�[�ް��i�e�j�̓ǂݍ���
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Taget_Key)
    sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�w�}�[�ް��i�e�j")
            Exit Sub
    
    End Select

    '�d�����於
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
            Exit Sub
    
    End Select
       
    Field1.text = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))             '�d�����於
    
    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        Field2.text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)   '�w�}�[��
    Else
        Field2.text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) & "-" & _
                        Format(CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) + 1, "#")
    End If
    Field3.text = Format(Now, "YYYY/MM/DD HH:MM")                   '���s����

'    Field3.Text = Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 1, 4) & "/" & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 5, 2) & "/" & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 7, 2) & " " & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 9, 2) & ":" & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 11, 2)

    '���F��
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Sub
    
    End Select
    Field4.text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)             '���F��
    
    '�S����
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Exit Sub
    
    End Select
    Field5.text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)             '�S����
    
    '���P�^�S����
    lblS_Tanto1.Visible = PRI_S_TANTO
    fldS_Tanto.Visible = PRI_S_TANTO
    speS_tanto1.Visible = PRI_S_TANTO
    l_S_Tanto1.Visible = PRI_S_TANTO
    If PRI_S_TANTO Then
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN05_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(P_CODEREC.C_RNAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                Exit Sub
        
        End Select
        fldS_Tanto.text = StrConv(P_CODEREC.C_RNAME, vbUnicode)         '���P�^�S����
    End If
    
    
    Select Case StrConv(P_SSHIJI_O_REC.SHIJI_F, vbUnicode)              '2007.08.31 �w���`��
        Case P_SHIJI_F_NORMAL           '���O
            lblSHIJI_F.Caption = " ���@�O "
        Case P_SHIJI_F_SPOT             '��߯�
            lblSHIJI_F.Caption = "�X�|�b�g"
        Case P_SHIJI_F_KEPPIN           '���i����
            lblSHIJI_F.Caption = "���i����"
        Case Else
            lblSHIJI_F.Caption = ""
    End Select
    
    
    
    
    
    
    Field7.text = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)            '�i��
                                                                        '����
    SHIJI_QTY = CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))
'    SHIJI_QTY = CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode))
    Field8.text = Format(SHIJI_QTY, "#0")
    '�i���^�I��
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
            Call UniCode_Conv(ITEMREC.ST_RETU, "")
            Call UniCode_Conv(ITEMREC.ST_REN, "")
            Call UniCode_Conv(ITEMREC.ST_DAN, "")
        
            Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Sub
    
    End Select
    Field9.text = StrConv(ITEMREC.HIN_NAME, vbUnicode)                      '�i��

    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
        Field10.text = ""                                                   '�W���I��
    Else
        ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        'P_SYS.INI--> PI0010.INI
        If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
        Else
            ST_SOKO = RTrim(c)
        End If
        
        
        Field10.text = Trim(ST_SOKO) & "-" & _
                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_DAN, vbUnicode)
    End If

    Field11.text = Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))    '���i���׽
    Field12.text = Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))    '�t���׽
    Field13.text = Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))    '���E�׽


    '���x���\�t�v��L��
    If StrConv(ITEMREC.G_LABEL_NON, vbUnicode) = P_G_LABEL_OFF Then
        lblLabel_NIN.Caption = "******"
        lblLabel_TIMES.Caption = "******"
    Else
        lblLabel_NIN.Caption = ""
        lblLabel_TIMES.Caption = ""
    End If


    '�󕥐�
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            Exit Sub
    
    End Select
    Field14.text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))   '���i����z��
    

    '�����ނ̃��[�v
    cnt = 0

    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Taget_Key Or _
                    StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_KOSOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�w�}�[�ް��i�q�j")
                Exit Sub
        
        End Select
        '�i�ڃ}�X�^�ǂݍ���
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                '�����އ�
                Field15.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field16.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field16.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field17.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field17.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field17.text = Format(CDbl(SURYO), "#0")
                Else
                    Field17.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field18.text = ""
                Else
                    
                    
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI-->PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    
                    Field18.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                '�����އ�
                Field19.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field20.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field20.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field21.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field21.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field21.text = Format(CDbl(SURYO), "#0")
                Else
                    Field21.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field22.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI-->PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field22.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                '�����އ�
                Field23.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field24.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field24.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field25.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field25.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field25.text = Format(CDbl(SURYO), "#0")
                Else
                    Field25.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field26.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field26.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
            Case 4
            
                '�����އ�
                Field27.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field28.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field28.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field29.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field29.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field29.text = Format(CDbl(SURYO), "#0")
                Else
                    Field29.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field30.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field30.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 5
                '�����އ�
                Field31.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field32.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field32.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field33.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field33.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field33.text = Format(CDbl(SURYO), "#0")
                Else
                    Field33.text = Format(CDbl(SURYO), "#0.00")
                End If
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field34.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field34.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
        
        End Select
        com = BtOpGetNext
    
    Loop


    '�O�����ނ̃��[�v
    cnt = 0

    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_GAISOU)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Taget_Key Or _
                    StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�w�}�[�ް��i�q�j")
                Exit Sub
        
        End Select
        '�i�ڃ}�X�^�ǂݍ���
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                '�O�����އ�
                Field35.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�O�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field36.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field36.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�O�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field37.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field37.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                
                
                
                SURYO = Format(Int(CDbl(SHIJI_QTY / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))), "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field37.text = Format(CDbl(SURYO), "#0")
                Else
                    Field37.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field38.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field38.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                '�O�����އ�
                Field39.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�O�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field40.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field40.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�O�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field41.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field41.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                SURYO = Format(Int(CDbl(SHIJI_QTY / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))), "00000000.00")
'                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field41.text = Format(CDbl(SURYO), "#0")
                Else
                    Field41.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field42.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field42.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                '�O�����އ�
                Field43.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '�O�����ށ@����
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field44.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field44.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '�O�����ށ@����
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field45.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field45.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
'                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                SURYO = Format(Int(CDbl(SHIJI_QTY / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))), "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field45.text = Format(CDbl(SURYO), "#0")
                Else
                    Field45.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field46.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field46.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
        
        End Select
    
        com = BtOpGetNext
    
    Loop

    Field47.text = Trim(StrConv(P_SSHIJI_O_REC.BIKOU, vbUnicode))               '���l
    
    
        
    '��Ɠ��^���ʁ^�S��     2013.01.16 �폜
'    ShaSagyo_Day.Visible = PRI_SAGYO_DAY
'    LineSagyo_Day1.Visible = PRI_SAGYO_DAY
'    LineSagyo_Day2.Visible = PRI_SAGYO_DAY
'    LineSagyo_Day3.Visible = PRI_SAGYO_DAY
'    lblSagyo_day1.Visible = PRI_SAGYO_DAY
'    lblSagyo_day2.Visible = PRI_SAGYO_DAY
'    lblSagyo_day3.Visible = PRI_SAGYO_DAY
        
    
    
    '���{�쐬�̈󎚗L��
    If StrConv(P_SSHIJI_O_REC.SAMPLE_F, vbUnicode) = P_SAMPLE_F_OFF Then        '���{�쐬
        lblSample.Visible = False
        Shape10.Visible = False
    Else
        lblSample.Visible = True
        Shape10.Visible = True
    End If

    
    
    'Ҳ��ް����
    fldMain_Bcr.Visible = PRI_MAIN_BCR
    If PRI_MAIN_BCR Then
        fldMain_Bcr.text = "*" & Trim(Field2.text) & "*"
    End If

    
    '���ה��l
    If PRI_BIKOU_BCR Then
        fldBIKOU.Visible = False
    
        fldSyuka_No.Visible = True
        fldSyuka_Bcr.Visible = True
    
    Else
        fldBIKOU.Visible = True
        
        fldSyuka_No.Visible = False
        fldSyuka_Bcr.Visible = False
    End If

    '�������i   '2011.08.04
'    lblDOUKON.Visible = PRI_DOUKON
'    lblDOUKON_GOUHI.Visible = PRI_DOUKON


    Select Case PRI_DOUKON
        Case 0
            lblDOUKON.Visible = False
            lblKOKUIN.Visible = False
            lblDOUKON_GOUHI.Visible = False
        Case 1
            lblDOUKON.Visible = True
            lblKOKUIN.Visible = False
            lblDOUKON_GOUHI.Visible = True
        Case 2
            lblDOUKON.Visible = False
            lblKOKUIN.Visible = True
            lblDOUKON_GOUHI.Visible = False
    End Select
    '�������i   '2011.08.04




    '�\���^����
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    '���Ɋ�����
    l_Nyuko_IN1.Visible = PRI_NYUKO_IN
    l_Nyuko_IN2.Visible = PRI_NYUKO_IN
    l_Nyuko_IN3.Visible = PRI_NYUKO_IN
    l_Nyuko_IN4.Visible = PRI_NYUKO_IN

    lblNyuko_In.Visible = PRI_NYUKO_IN

    '���͊�����
    l_Input_IN1.Visible = PRI_INPUT_IN
    l_Input_IN2.Visible = PRI_INPUT_IN
    l_Input_IN3.Visible = PRI_INPUT_IN
    l_Input_IN4.Visible = PRI_INPUT_IN

    lblInput_In.Visible = PRI_INPUT_IN


    If Not PRI_NYUKO_IN And Not PRI_NYUKO_IN Then
        l_IN_Center.Visible = False
    Else
        l_IN_Center.Visible = True
    End If

    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        lblBunnou.Visible = False
    Else
        lblBunnou.Visible = True
    End If

    '�����@�i�ԁ^���^����   2007.05.22
    ShaHINBAN_BIKOU.Visible = PRI_HINBAN_BIKOU
    
    LineHINBAN_BIKOU1.Visible = PRI_HINBAN_BIKOU
    LineHINBAN_BIKOU2.Visible = PRI_HINBAN_BIKOU
    LineHINBAN_BIKOU3.Visible = PRI_HINBAN_BIKOU
    LineHINBAN_BIKOU4.Visible = PRI_HINBAN_BIKOU

    lblHINBAN_BIKOU1.Visible = PRI_HINBAN_BIKOU
    lblHINBAN_BIKOU2.Visible = PRI_HINBAN_BIKOU
    lblHINBAN_BIKOU3.Visible = PRI_HINBAN_BIKOU

    Field60.Visible = PRI_HINBAN_BIKOU
    Field61.Visible = PRI_HINBAN_BIKOU
    Field62.Visible = PRI_HINBAN_BIKOU
    
    Field60.text = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)           '�i��
    
    
    
'2011.08.04    Field61.text = StrConv(P_SSHIJI_O_REC.Shiji_No, vbUnicode)          '��
                                                                        '����
    Field62.text = Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0")


    If JISSEKI_DSP = "s" Then           '2008.08.19
        Label116.Caption = "�b"
    End If

    '���Ӄ^�C�g��
    If CStr(JISEKI_TITLE(0)) = "" Then
    Else
        lblJISEKI_TITLE.Caption = CStr(JISEKI_TITLE(0)) & "/" & CStr(JISEKI_TITLE(1))
        
    End If
    '���Ӄ^�C�g��
    If CStr(TASEKI_TITLE(0)) = "" Then
    Else
        LblTASEKI_TITLE.Caption = CStr(TASEKI_TITLE(0)) & "/" & CStr(TASEKI_TITLE(1))
        
    End If
        
        
        
'2008.09.20    '�O����т̊l��
'2008.09.20    Call UniCode_Conv(K1_wP_SSHIJI_O.KAN_F, P_KAN_ON)   '�����׸�
'2008.09.20                                                        '�d������
'2008.09.20    Call UniCode_Conv(K1_wP_SSHIJI_O.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
'2008.09.20                                                        '���ƕ�
'2008.09.20    Call UniCode_Conv(K1_wP_SSHIJI_O.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
'2008.09.20                                                        '�����O
'2008.09.20    Call UniCode_Conv(K1_wP_SSHIJI_O.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
'2008.09.20                                                        '�i��
'2008.09.20    Call UniCode_Conv(K1_wP_SSHIJI_O.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
'2008.09.20                                                        '������
'2008.09.20    Call UniCode_Conv(K1_wP_SSHIJI_O.KAN_DT, "zzzzzzzz")
'2008.09.20                                                        '�w�}�\��
'2008.09.20    Call UniCode_Conv(K1_wP_SSHIJI_O.SHIJI_NO, "zzzzzzzz")
'2008.09.20    sts = BTRV(BtOpGetLess, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K1_wP_SSHIJI_O, Len(K1_wP_SSHIJI_O), 1)
'2008.09.20    Select Case sts
'2008.09.20        Case BtNoErr
'2008.09.20            If StrConv(wP_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_ON Or _
'2008.09.20                StrConv(wP_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) Or _
'2008.09.20                StrConv(wP_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) Or _
'2008.09.20                StrConv(wP_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) Or _
'2008.09.20                StrConv(wP_SSHIJI_O_REC.HIN_GAI, vbUnicode) <> StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode) Then
'2008.09.20                    lblBEF_JISSEKI.Caption = ""
'2008.09.20                    lblBEF_BEFORE1.Caption = ""
'2008.09.20                    lblBEF_BEFORE2.Caption = ""
'2008.09.20                    lblBEF_BEFORE3.Caption = ""
'2008.09.20'                    lblBEF_BEFORE4.Caption = ""
'2008.09.20                    lblBEF_SAGYO1.Caption = ""
'2008.09.20                    lblBEF_SAGYO2.Caption = ""
'2008.09.20                    lblBEF_SAGYO3.Caption = ""
'2008.09.20                    lblBEF_AFTER1.Caption = ""
'2008.09.20                    lblBEF_AFTER2.Caption = ""
'2008.09.20                    lblBEF_KAKOU1.Caption = ""
'2008.09.20                    lblBEF_JISEKI.Caption = ""
'2008.09.20                    lblBEF_TASEKI.Caption = ""
'2008.09.20
'2008.09.20            Else
'2008.09.20
'2008.09.20
'2008.09.20                    '��Ƈ@
'2008.09.20
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_SAGYO1.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_SAGYO1.Caption = ""
'2008.09.20                        Else
'2008.09.20
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_SAGYO1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_SAGYO1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20
'2008.09.20                            End If
'2008.09.20
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '��ƇA
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_SAGYO2.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_SAGYO2.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_SAGYO2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_SAGYO2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '��ƇB
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_SAGYO3.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_SAGYO3.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_SAGYO3.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20
'2008.09.20
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_SAGYO3.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '�����@
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_BEFORE1.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_BEFORE1.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_BEFORE1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_BEFORE1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '�����A
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_BEFORE2.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_BEFORE2.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_BEFORE2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_BEFORE2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '�����B
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_BEFORE3.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_BEFORE3.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_BEFORE3.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20                            Else
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_BEFORE3.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '��Еt���@
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_AFTER1.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_AFTER1.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_AFTER1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20                            Else
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_AFTER1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '��Еt���A
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_AFTER2.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_AFTER2.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_AFTER2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20                            Else
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_AFTER2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '���H�@
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_KAKOU1.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(9).NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_KAKOU1.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_KAKOU1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "�b"
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(9).TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20                                lblBEF_KAKOU1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
'2008.09.20
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20
'2008.09.20                    '����
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_JISEKI.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.JISEKI_NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.JISEKI_TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_JISEKI.Caption = ""
'2008.09.20                        Else
'2008.09.20
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_JISEKI.Caption = wkEDIT_NIN & "�l�~" & _
'2008.09.20                                                        wkEDIT_TIMES & "�b " & _
'2008.09.20                                                        StrConv(wP_SSHIJI_O_REC.JISEKI_NAME, vbUnicode)
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_JISEKI.Caption = wkEDIT_NIN & "�l�~" & _
'2008.09.20                                                        wkEDIT_TIMES & "�� " & _
'2008.09.20                                                        StrConv(wP_SSHIJI_O_REC.JISEKI_NAME, vbUnicode)
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20                    '����
'2008.09.20                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) Or _
'2008.09.20                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) Then
'2008.09.20                        lblBEF_TASEKI.Caption = ""
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.TASEKI_NIN, "0.0")
'2008.09.20                        Call UniCode_Conv(wP_SSHIJI_O_REC.TASEKI_TIMES, "000.00")
'2008.09.20                    Else
'2008.09.20                        If CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) = 0 Then
'2008.09.20                            lblBEF_TASEKI.Caption = ""
'2008.09.20                        Else
'2008.09.20                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0.0")
'2008.09.20                            If Right(wkValue, 1) = "0" Then
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0")
'2008.09.20                            Else
'2008.09.20                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0.0")
'2008.09.20                            End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                            If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) * 60, "#0")
'2008.09.20                                Else
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) * 60, "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) * 60, "#0.00")
'2008.09.20                                    End If
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_TASEKI.Caption = wkEDIT_NIN & "�l�~" & _
'2008.09.20                                                        wkEDIT_TIMES & "�b " & _
'2008.09.20                                                        StrConv(wP_SSHIJI_O_REC.TASEKI_NAME, vbUnicode)
'2008.09.20
'2008.09.20
'2008.09.20                            Else
'2008.09.20                                wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.00")
'2008.09.20                                If Right(wkValue, 2) = "00" Then
'2008.09.20                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0")
'2008.09.20                                Else
'2008.09.20
'2008.09.20                                    If Right(wkValue, 1) = "0" Then
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.0")
'2008.09.20                                    Else
'2008.09.20                                        wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.00")
'2008.09.20                                    End If
'2008.09.20
'2008.09.20                                End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                                lblBEF_TASEKI.Caption = wkEDIT_NIN & "�l�~" & _
'2008.09.20                                                        wkEDIT_TIMES & "�� " & _
'2008.09.20                                                        StrConv(wP_SSHIJI_O_REC.TASEKI_NAME, vbUnicode)
'2008.09.20
'2008.09.20
'2008.09.20                            End If
'2008.09.20                        End If
'2008.09.20                    End If
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                    If JISSEKI_DSP = "s" Then       '2008.08.19
'2008.09.20
'2008.09.20
'2008.09.20                        '���v�̌v�Z
'2008.09.20                        Total_Times = 0
'2008.09.20                        For i = 0 To 9
'2008.09.20                            If i <> 3 Then
'2008.09.20                                Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)) * 60))
'2008.09.20                            End If
'2008.09.20                        Next i
'2008.09.20
'2008.09.20                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) * (CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) * 60))
'2008.09.20                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) * (CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) * 60))
'2008.09.20
'2008.09.20                        If Total_Times = 0 Or CDbl(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) = 0 Then
'2008.09.20                            AVE = 0
'2008.09.20                        Else
'2008.09.20                            AVE = Round(CDbl(Total_Times / CDbl(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))), 1)
'2008.09.20                        End If
'2008.09.20
'2008.09.20                        wkValue = Format(Total_Times, "#0.00")
'2008.09.20                        If Right(wkValue, 2) = "00" Then
'2008.09.20                            wkEDIT_TIMES = Format(Total_Times, "#0")
'2008.09.20                        Else
'2008.09.20                            wkEDIT_TIMES = Format(Total_Times, "#0.00")
'2008.09.20                        End If
'2008.09.20
'2008.09.20                        lblBEF_JISSEKI.Caption = "�O��:" & Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
'2008.09.20                                                    Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
'2008.09.20                                                    Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2) & ":" & _
'2008.09.20                                                    Format(CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0") & _
'2008.09.20                                                    "�� " & _
'2008.09.20                                                    wkEDIT_TIMES & "�b(" & Format(AVE, "#0.0") & "�b/��)"
'2008.09.20
'2008.09.20                        '�ڕW�̌v�Z
'2008.09.20                        Total_Times = 0
'2008.09.20                        For i = 0 To 2
'2008.09.20                            Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)) * 60))
'2008.09.20                        Next i
'2008.09.20
'2008.09.20                        If CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) = 0 Then
'2008.09.20                            AVE = 0
'2008.09.20                        Else
'2008.09.20                            AVE = Round(Total_Times / CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), 1)
'2008.09.20                        End If
'2008.09.20
'2008.09.20
'2008.09.20                        Target = AVE * CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode))
'2008.09.20                        lblTarget1.Caption = "����ڕW�F" & Format(Target, "#0") & "�b"
'2008.09.20
'2008.09.20                        wkValue = Format(AVE, "#0.0")
'2008.09.20                        If Right(wkValue, 1) = "0" Then
'2008.09.20                            wkAVE = Format(AVE, "#0")
'2008.09.20                        Else
'2008.09.20                            wkAVE = Format(AVE, "#0.0")
'2008.09.20                        End If
'2008.09.20                        lblTarget2.Caption = wkAVE & "�b/�~" & Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0") & "��"
'2008.09.20
'2008.09.20                    Else
'2008.09.20
'2008.09.20
'2008.09.20
'2008.09.20                        '���v�̌v�Z
'2008.09.20                        Total_Times = 0
'2008.09.20                        For i = 0 To 9
'2008.09.20                            If i <> 3 Then
'2008.09.20                                Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)))
'2008.09.20                            End If
'2008.09.20                        Next i
'2008.09.20
'2008.09.20                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)))
'2008.09.20                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)))
'2008.09.20
'2008.09.20                        If Total_Times = 0 Or CDbl(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) = 0 Then
'2008.09.20                            AVE = 0
'2008.09.20                        Else
'2008.09.20                            AVE = Round(CDbl(Total_Times / CDbl(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))), 1)
'2008.09.20                        End If
'2008.09.20
'2008.09.20                        wkValue = Format(Total_Times, "#0.00")
'2008.09.20                        If Right(wkValue, 2) = "00" Then
'2008.09.20                            wkEDIT_TIMES = Format(Total_Times, "#0")
'2008.09.20                        Else
'2008.09.20                            wkEDIT_TIMES = Format(Total_Times, "#0.00")
'2008.09.20                        End If
'2008.09.20
'2008.09.20                        lblBEF_JISSEKI.Caption = "�O��:" & Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
'2008.09.20                                                    Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
'2008.09.20                                                    Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2) & ":" & _
'2008.09.20                                                    Format(CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0") & _
'2008.09.20                                                    "�� " & _
'2008.09.20                                                    wkEDIT_TIMES & "��(" & Format(AVE, "#0.0") & "��/��)"
'2008.09.20
'2008.09.20                        '�ڕW�̌v�Z
'2008.09.20                        Total_Times = 0
'2008.09.20                        For i = 0 To 2
'2008.09.20                            Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)))
'2008.09.20                        Next i
'2008.09.20
'2008.09.20                        If CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) = 0 Then
'2008.09.20                            AVE = 0
'2008.09.20                        Else
'2008.09.20                            AVE = Round(Total_Times / CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), 1)
'2008.09.20                        End If
'2008.09.20
'2008.09.20
'2008.09.20                        Target = AVE * CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode))
'2008.09.20                        lblTarget1.Caption = "����ڕW�F" & Format(Target, "#0") & "��"
'2008.09.20
'2008.09.20                        wkValue = Format(AVE, "#0.0")
'2008.09.20                        If Right(wkValue, 1) = "0" Then
'2008.09.20                            wkAVE = Format(AVE, "#0")
'2008.09.20                        Else
'2008.09.20                            wkAVE = Format(AVE, "#0.0")
'2008.09.20                        End If
'2008.09.20                        lblTarget2.Caption = wkAVE & "��/�~" & Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0") & "��"
'2008.09.20
'2008.09.20                    End If
'2008.09.20
'2008.09.20                End If
'2008.09.20
'2008.09.20        Case BtErrEOF
'2008.09.20            lblBEF_JISSEKI.Caption = ""
'2008.09.20            lblBEF_BEFORE1.Caption = ""
'2008.09.20            lblBEF_BEFORE2.Caption = ""
'2008.09.20            lblBEF_BEFORE3.Caption = ""
'2008.09.20'            lblBEF_BEFORE4.Caption = ""
'2008.09.20            lblBEF_SAGYO1.Caption = ""
'2008.09.20            lblBEF_SAGYO2.Caption = ""
'2008.09.20            lblBEF_SAGYO3.Caption = ""
'2008.09.20            lblBEF_AFTER1.Caption = ""
'2008.09.20            lblBEF_AFTER2.Caption = ""
'2008.09.20            lblBEF_KAKOU1.Caption = ""
'2008.09.20            lblBEF_JISEKI.Caption = ""
'2008.09.20            lblBEF_TASEKI.Caption = ""
'2008.09.20
'2008.09.20
'2008.09.20        Case Else
'2008.09.20            Call File_Error(sts, BtOpGetEqual, "�w�}�[�ް��i�e�j")
'2008.09.20            Exit Sub
'2008.09.20
'2008.09.20    End Select
    
    
    
    
    
    '2008.09.20 ��
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Tanka_F = False
        
        
    Select Case sts
        Case BtNoErr
            
            wkDate = StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode)
            
            If IsDate(Mid(wkDate, 1, 4) & "/" & Mid(wkDate, 5, 2) & "/" & Mid(wkDate, 7, 2)) Then
                Tanka_F = True
            End If
        
        
        Case BtErrKeyNotFound
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Sub
    End Select
    
    
    If Tanka_F Then
    
    
            lblBEF_JISSEKI = "���ύ�Ǝ��ԁ^���b�g���F" & Format(StrConv(ITEMREC.SEI_LOT, vbUnicode), "#0")
                    
                    
'�]�T���𖳌��Ƃ��� 2008.10.03
Call UniCode_Conv(P_KANRIREC.KOUTEI_R_RATE, "1.00")
                    
                    
            Call UniCode_Conv(K0_SE_LOC_TANKA_M.SE_IO_TANKA_No, StrConv(ITEMREC.SE_IO_TANKA_No, vbUnicode))
            sts = BTRV(BtOpGetEqual, SE_LOC_TANKA_M_POS, SE_LOC_TANKA_M_REC, Len(SE_LOC_TANKA_M_REC), K0_SE_LOC_TANKA_M, Len(K0_SE_LOC_TANKA_M), 0)
            Select Case sts
                Case BtNoErr
                    
                
                
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(SE_LOC_TANKA_M_REC.SE_Name, "")
                        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���o�ɒP���ݒ�}�X�^")
                    Exit Sub
            End Select
                    
                    
                    
            wkTOTAL = 0
            '���i����
            For i = 3 To 8
            
                If IsNumeric(StrConv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, vbUnicode)) Then
                Else
                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "0.0")
            
                End If
            
            
            
            
            Next i
                                    
                                    
            lblBEF_BEFORE1 = Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(3).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0") & "��" & " " & _
                                Trim(StrConv(SE_LOC_TANKA_M_REC.SE_Name, vbUnicode))
            wkTOTAL = CDbl(Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(3).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0"))
            
            
            
            '�����ޏ���
            lblBEF_BEFORE2 = Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(4).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0") & "��" & " " & _
                                wkDate & " " & Trim(StrConv(ITEMREC.SE_TANKA_MEMO, vbUnicode))

            wkTOTAL = wkTOTAL + CDbl(Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(4).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0"))
            '�������i����
            lblBEF_BEFORE3 = Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(5).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0") & "��"
            wkTOTAL = wkTOTAL + CDbl(Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(5).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0"))
            
            
            
            
            '���x���\��
            
            For i = 0 To 4
            
                If IsNumeric(StrConv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, vbUnicode)) Then
                Else
                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "0.0")
            
                End If
            
            
            
            
            Next i
            
            
            lblBEF_SAGYO1 = Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(0).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2), "#0.0") & "��"
            wkTOTAL = wkTOTAL + CDbl(Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(0).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2), "#0.0"))
            '�����
            lblBEF_SAGYO2 = Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(7).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(8).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(1).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(2).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(3).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(4).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2), "#0.0") & "��"
            wkTOTAL = wkTOTAL + CDbl(Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(7).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.BEF_KOUTEI(8).BEF_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(1).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(2).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(3).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2) + _
                            ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.MAIN_KOUTEI(4).MAIN_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) * SHIJI_QTY / 60)), 2), "#0.0"))
            
            
            lblBEF_SAGYO3 = ""
        
            
            '���i����
            
            For i = 1 To 2
            
                If IsNumeric(StrConv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, vbUnicode)) Then
                Else
                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "0.0")
            
                End If
            
            
            
            
            Next i
            
            
            lblBEF_AFTER1 = Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.AFT_KOUTEI(1).AFT_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0") & "��"
            wkTOTAL = wkTOTAL + CDbl(Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.AFT_KOUTEI(1).AFT_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0"))
            lblBEF_AFTER2 = Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.AFT_KOUTEI(2).AFT_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0") & "��"
            wkTOTAL = wkTOTAL + CDbl(Format(ToHalfAdjust(CCur(CDbl(StrConv(ITEMREC.AFT_KOUTEI(2).AFT_KOUTEI, vbUnicode) * CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) / 60)), 2), "#0.0"))
            
            lblBEF_KAKOU1 = ""
                    
            lblBEF_JISEKI = ""
            lblBEF_TASEKI = ""





            lblTarget1 = Trim(Format(wkTOTAL, "#0.0")) & "��" & " "
            lblTarget2 = ""
    

    
    
    
    Else
        '�P�����ݒ莞
        
            lblBEF_JISSEKI = "�P �� �� �� ��"
                    
            lblBEF_BEFORE1 = ""
            lblBEF_BEFORE2 = ""
            lblBEF_BEFORE3 = ""
        
            lblBEF_SAGYO1 = ""
            lblBEF_SAGYO2 = ""
            lblBEF_SAGYO3 = ""
        
            lblBEF_AFTER1 = ""
            lblBEF_AFTER2 = ""
            lblBEF_KAKOU1 = ""
                    
            lblBEF_JISEKI = ""
            lblBEF_TASEKI = ""

            lblTarget1 = ""
            lblTarget2 = ""
    
    
    
    
    End If
    
    
    
    
    '�O����т̊l��
    Call UniCode_Conv(K1_wP_SSHIJI_O.KAN_F, P_KAN_ON)   '�����׸�
                                                        '�d������
    Call UniCode_Conv(K1_wP_SSHIJI_O.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                        '���ƕ�
    Call UniCode_Conv(K1_wP_SSHIJI_O.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
                                                        '�����O
    Call UniCode_Conv(K1_wP_SSHIJI_O.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
                                                        '�i��
    Call UniCode_Conv(K1_wP_SSHIJI_O.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
                                                        '������
    Call UniCode_Conv(K1_wP_SSHIJI_O.KAN_DT, "zzzzzzzz")
                                                        '�w�}�\��
    Call UniCode_Conv(K1_wP_SSHIJI_O.SHIJI_No, "zzzzzzzz")
    sts = BTRV(BtOpGetLess, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K1_wP_SSHIJI_O, Len(K1_wP_SSHIJI_O), 1)
    Select Case sts
        Case BtNoErr
            If StrConv(wP_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_ON Or _
                StrConv(wP_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) Or _
                StrConv(wP_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) Or _
                StrConv(wP_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) Or _
                StrConv(wP_SSHIJI_O_REC.HIN_GAI, vbUnicode) <> StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode) Then
    
    
    
            Else
    
    
                lblTarget1 = lblTarget1 & "�O��F" & Mid(StrConv(wP_SSHIJI_O_REC.KAN_DT, vbUnicode), 1, 4) & "/" & _
                                                    Mid(StrConv(wP_SSHIJI_O_REC.KAN_DT, vbUnicode), 5, 2) & "/" & _
                                                    Mid(StrConv(wP_SSHIJI_O_REC.KAN_DT, vbUnicode), 7, 4) & _
                                                    "�F" & _
                                                    Format(CInt(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0") & "��"
                                                    
    
    
    
            End If
    
    
        Case BtErrEOF
        
        
        Case Else
            Call File_Error(sts, BtOpGetLess, "�w�}�[�f�[�^")
            Exit Sub
    End Select
    
    '2008.09.20 ��
            
    
    
    
    
    
    
    
    
    
    
    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        lblBunnou.Visible = False
    
    
        lblTarget1.Visible = True
        lblTarget2.Visible = True
    
    
    
    
    
    Else
        lblBunnou.Visible = True
    
'        lblTarget1.Visible = False
        lblTarget1.Visible = True
        lblTarget2.Visible = False
    
    
    End If

    Doukon_com = BtOpGetGreater
    Doukon_eof = False

    Doukon_cnt = 0

    EOF_F = False       '2012.04.17



' 2013.01.08 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    If GENSANKOKU_MSG_F Then                        '2013.02.19
        If Trim(chk_TORI_GENSANKOKU) <> "" Then
            GENSANKOKU_Alart.Visible = True
        End If
    End If                                          '2013.02.19
' 2013.01.08 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


' 2013.01.16 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    LblKAIKON_1.Visible = KAIKON_PRI
    LblKAIKON_2.Visible = KAIKON_PRI
' 2013.01.16 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

' 2013.11.05 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    lblNYUKA_KANSYOZAI.Visible = NYUKA_KANSYOZAI
' 2013.11.05 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

End Sub

Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORPortrait
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 5     '2012.04.17 10-->5
    Me.PageTopMargin = 5        '2012.04.17 10-->5
    Me.PageLeftMargin = 20
    Me.PageRightMargin = 20

    Me.documentName = "���i���w�}�[�F"

    DoEvents

End Sub

