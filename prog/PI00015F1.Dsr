VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PI00015F1 
   ClientHeight    =   15615
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   28560
   StartUpPosition =   3  'Windows �̊���l
   _ExtentX        =   50377
   _ExtentY        =   27543
   SectionData     =   "PI00015F1.dsx":0000
End
Attribute VB_Name = "PI00015F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Doukon_com      As Integer      '�\���^������Btrieve Operation
Private Doukon_eof      As Integer      '�\���^���� Eof

Private Doukon_cnt      As Integer      '�\���^������LINE COUNT


Private SHIJI_QTY       As Double       '����w����
Private Function fStrCut(ByRef CutTxt As String, _
                         ByVal CutLen As Long) As String
'���p�E�S�p�̍��݂��镶����𔼊p���Z�������Ŏ��o��
    Dim myLen As Long, SysCodeTxt As String
    SysCodeTxt = StrConv(CutTxt, vbFromUnicode)     '�������ϊ�
    myLen = LenB(SysCodeTxt)    '���p���Z�̃o�C�g�����擾
    If myLen <= CutLen Then     '�w��̒������Z���ꍇ
        fStrCut = CutTxt & Space$(CutLen - myLen)   '����Ȃ����̓X�y�[�X��
    Else    '�Y���̕�����̕��������ꍇ�A�w��̃o�C�g�ŃJ�b�g����
        fStrCut = StrConv(LeftB$(SysCodeTxt, CutLen), vbUnicode)
        If InStr(fStrCut, vbNullChar) > 0 Then
            '�����P�o�C�g�ڂŕ��f���ꂽ�ꍇ�̏���
            fStrCut = Left$(fStrCut, InStr(fStrCut, vbNullChar) - 1) & " "
        End If
    End If
End Function


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
    Me.Fields.Add "KO_HIN_NAME"         '�i��



End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
    
Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
    
Dim SURYO       As String

Dim ST_SOKO     As String
Dim c           As String * 128
    
Dim wkJgyobu    As String * 1
    
Dim ST_LOCATION As String * 8   '2013.03.31
    
Dim IN_WORD     As String       '2017.12.15
Dim OUT_WORD    As String       '2017.12.15
    


'    If Doukon_cnt > 19 Then                '2013.11.21
'        Exit Sub                           '2013.11.21
'    End If                                 '2013.11.21
    
    
    If Doukon_cnt > 19 Then                     '2013.11.21
        If Doukon_eof Then                      '2013.11.21
            Exit Sub                            '2013.11.21
        Else                                    '2013.11.21
            Doukon_cnt = 0                      '2013.11.21
        End If                                  '2013.11.21
    End If                                      '2013.11.21
    
    
    
    
    
    
    Me.Fields("ko_no").Value = Doukon_Tbl_No(Doukon_cnt)
    
    If Doukon_eof Then
        Me.Fields("KO_HIN_GAI") = ""        '�i��
        Me.Fields("KO_SYUBETSU") = ""       '���
        Me.Fields("KO_QTY") = ""            '����
        Me.Fields("KO_SHIJI_QTY") = ""      '����
        Me.Fields("KO_ST_LOCATION") = ""    '�I��
        Me.Fields("KO_ZAIKO_QTY") = ""      '���_�݌�
        Me.Fields("KO_ID_NO") = ""          'ID_NO
        Me.Fields("KO_ID_BCR") = ""         'ID_NO�ް����
        Me.Fields("KO_BIKOU") = ""          '���l
    
    
    Else
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
'        sts = BTRV(Doukon_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        sts = BTRV(Doukon_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K3_P_SSHIJI_K, Len(K3_P_SSHIJI_K), 3)
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Taget_Key Or _
                    StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                    Doukon_eof = True
                End If
            
            
                If Doukon_cnt = 0 Then              '2016.01.14
                    If Doukon_eof Then              '2016.01.14
                        Doukon_cnt = Doukon_cnt + 1 '2016.01.14
'                        eof = False                '2016.01.14
                        Exit Sub                   '2016.01.14
                    End If                          '2016.01.14
                End If                              '2016.01.14
            
            
            
            
            Case BtErrEOF
                
                Doukon_eof = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�Ώێw�}�[�ް��i�e�j")
                Exit Sub
        
        End Select
                                            
        If Doukon_eof Then
            Me.Fields("KO_HIN_GAI") = ""        '�i��
            Me.Fields("KO_SYUBETSU") = ""       '���
            Me.Fields("KO_QTY") = ""            '����
            Me.Fields("KO_SHIJI_QTY") = ""      '����
            Me.Fields("KO_ST_LOCATION") = ""    '�I��
            Me.Fields("KO_ZAIKO_QTY") = ""      '���_�݌�
            Me.Fields("KO_ID_NO") = ""          'ID_NO
            Me.Fields("KO_ID_BCR") = ""         'ID_NO�ް����
            Me.Fields("KO_BIKOU") = ""          '���l
            Me.Fields("KO_HIN_NAME") = ""       '�i��
                                            
                                            
                                            
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
'            Me.Fields("KO_SYUBETSU") = StrConv(P_CODEREC.C_RNAME, vbUnicode)               '2017.12.15
            Me.Fields("KO_SYUBETSU") = fStrCut(StrConv(P_CODEREC.C_RNAME, vbUnicode), 6)    '2017.12.15
                                                
                                                
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
            
            
'>>>>>  2016.01.27 �ǂݑւ��p�~
'            If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then                           '2013.03.31
'                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                                            '2013.03.31
'            Else                                                                                    '2013.03.31
'                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
'            End If                                                                                  '2013.03.31
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
'>>>>>  2016.01.27 �ǂݑւ��p�~
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    
                    
                    Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")     '2008.02.27
                    
                    
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
'                If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
                Else
                    ST_SOKO = RTrim(c)
                End If
                
                
                
                Me.Fields("KO_ST_LOCATION") = Trim(ST_SOKO) & "-" & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If
        
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            '�݌ɐ�
            
            
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.03.31
'            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
'                wkJgyobu = BUZAI
'            Else
'                'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
'                wkJgyobu = YUKO_JGYOBU                          '2012.04.04
'            End If
            
 '>>>>>>>   �ǂݑւ��p�~ 2016.01.27
'            Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                Case SHIZAI
'                    wkJgyobu = BUZAI
'                Case SETSUBI
'                    wkJgyobu = YUKO_JGYOBU
'                Case Else
'                    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'            End Select
            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
 '>>>>>>>   �ǂݑւ��p�~ 2016.01.27
            
            
            
'            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
'                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , Jyogai_Soko_umu) Then
                
                
            ST_LOCATION = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), ST_LOCATION, , , Jyogai_Soko_umu) Then
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.03.31
                
                
                
                Exit Sub
            
            End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
            Me.Fields("KO_ZAIKO_QTY") = Format(Sumi_Qty + Mi_Qty, "#0")
            '���lOR�o���ް���h
            
            
        
            Select Case PRI_BIKOU_BCR
                Case 0          '���l
                    Me.Fields("KO_BIKOU") = Trim(StrConv(P_SSHIJI_K_REC.KO_BIKOU, vbUnicode))
            
                Case 1          'ID_NO�ް����
            
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
            
                Case 2
                    Me.Fields("KO_HIN_NAME") = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                
            End Select
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
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
    If Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) = "" Then
        Field61.text = "�����Ȃ�"
    Else
        Field61.text = "������:" & Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) & Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT_SEQ, vbUnicode))
    End If
'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
    
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
    
    
    Select Case StrConv(P_SSHIJI_O_REC.SHIJI_F, vbUnicode)              '2007.11.08 �w���`��
        Case P_SHIJI_F_NORMAL           '���O
            lblSHIJI_F.Caption = " ���@�O "
        Case P_SHIJI_F_SPOT             '��߯�
            lblSHIJI_F.Caption = "�X�|�b�g"
        Case P_SHIJI_F_KEPPIN           '���i����
            lblSHIJI_F.Caption = "���i����"
        Case P_SHIJI_F_SAIKON           '�č��� 2007.11.09
            lblSHIJI_F.Caption = "�č���"
        Case Else
            lblSHIJI_F.Caption = ""
    End Select
    
    
    
    
    
    
    
    Field7.text = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)            '�i��
                                                                        '����
    SHIJI_QTY = CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))
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
'        If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
        If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
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
        If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then                           '2013.03.31
            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                                            '2013.03.31
        Else                                                                                    '2013.03.31
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
        End If                                                                                  '2013.03.31
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then  '2016.01.13
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then  '2016.01.13
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then  '2016.01.13
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then      '2016.01.13
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
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then      '2016.01.13
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
    Select Case PRI_BIKOU_BCR
        Case 0
            fldBIKOU.Visible = True
            
            fldSyuka_No.Visible = False
            fldSyuka_Bcr.Visible = False
            fldHin_Name.Visible = False

        Case 1
        
            fldSyuka_No.Visible = True
            fldSyuka_Bcr.Visible = True

            fldBIKOU.Visible = False
            fldHin_Name.Visible = False

        Case 2
            
            fldHin_Name.Visible = True
        
            fldSyuka_No.Visible = False
            fldSyuka_Bcr.Visible = False

            fldBIKOU.Visible = False

        Case Else
            fldHin_Name.Visible = False
        
            fldSyuka_No.Visible = False
            fldSyuka_Bcr.Visible = False

            fldBIKOU.Visible = False
    End Select

    '�������i
    lblDOUKON.Visible = PRI_DOUKON
    lblDOUKON_GOUHI.Visible = PRI_DOUKON


'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18
    '�\���^����
'    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
'    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_DOUKON)
'    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")


    Call UniCode_Conv(K3_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K3_P_SSHIJI_K.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K3_P_SSHIJI_K.ST_TANABAN, "")

'--------------------------------------------------- ���  ���ޑΉ��@2012.03.18

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
                    lblBEF_JISSEKI.Caption = ""
                    lblBEF_BEFORE1.Caption = ""
                    lblBEF_BEFORE2.Caption = ""
                    lblBEF_BEFORE3.Caption = ""
                    lblBEF_BEFORE4.Caption = ""
                    lblBEF_SAGYO1.Caption = ""
                    lblBEF_SAGYO2.Caption = ""
                    lblBEF_SAGYO3.Caption = ""
                    lblBEF_AFTER1.Caption = ""
                    lblBEF_AFTER2.Caption = ""
                    lblBEF_JISEKI.Caption = ""
                    lblBEF_TASEKI.Caption = ""
            
            Else
                    

                    '��Ƈ@
                    
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) Then
                        lblBEF_SAGYO1.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 Then
                            lblBEF_SAGYO1.Caption = ""
                        Else
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            
                            
                            lblBEF_SAGYO1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '��ƇA
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) Then
                        lblBEF_SAGYO2.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 Then
                            lblBEF_SAGYO2.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_SAGYO2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '��ƇB
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) Then
                        lblBEF_SAGYO3.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 Then
                            lblBEF_SAGYO3.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0")
                            Else
                                
                                
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_SAGYO3.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '�����@
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE1.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE1.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '�����A
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE2.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE2.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '�����B
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE3.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE3.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE3.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '�����C
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE4.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE4.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE4.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '��Еt���@
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) Then
                        lblBEF_AFTER1.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 Then
                            lblBEF_AFTER1.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_AFTER1.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    '��Еt���A
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) Then
                        lblBEF_AFTER2.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 Then
                            lblBEF_AFTER2.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_AFTER2.Caption = wkEDIT_NIN & "�l�~" & wkEDIT_TIMES & "��"
                        End If
                    End If
                    
                    '����
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) Then
                        lblBEF_JISEKI.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.JISEKI_NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.JISEKI_TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) = 0 Then
                            lblBEF_JISEKI.Caption = ""
                        Else
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0")
                            Else
                            
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.00")
                                End If
                            
                            End If
                            
                            
                            
                            lblBEF_JISEKI.Caption = wkEDIT_NIN & "�l�~" & _
                                                    wkEDIT_TIMES & "�� " & _
                                                    StrConv(wP_SSHIJI_O_REC.JISEKI_NAME, vbUnicode)
                        End If
                    End If
                    '����
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) Then
                        lblBEF_TASEKI.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.TASEKI_NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.TASEKI_TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) = 0 Then
                            lblBEF_TASEKI.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0")
                            Else
                                
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.00")
                                End If
                                
                            End If
                            
                            
                            
                            lblBEF_TASEKI.Caption = wkEDIT_NIN & "�l�~" & _
                                                    wkEDIT_TIMES & "�� " & _
                                                    StrConv(wP_SSHIJI_O_REC.TASEKI_NAME, vbUnicode)
                        End If
                    End If
                    
                    
                    
                    
                    '���v�̌v�Z
                    Total_Times = 0
                    For i = 0 To 8
                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)))
                    Next i
                    
                    Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)))
                    Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)))
                    
                    If Total_Times = 0 Then
                        AVE = 0
                    Else
                        AVE = Round(CDbl(Total_Times / CDbl(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))), 1)
                    End If
            
                    wkValue = Format(Total_Times, "#0.00")
                    If Right(wkValue, 2) = "00" Then
                        wkEDIT_TIMES = Format(Total_Times, "#0")
                    Else
                        wkEDIT_TIMES = Format(Total_Times, "#0.00")
                    End If
            
                    lblBEF_JISSEKI.Caption = "�O��:" & Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2) & ":" & _
                                                Format(CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0") & _
                                                "�� " & _
                                                wkEDIT_TIMES & "��(" & Format(AVE, "#0.0") & "��/��)"

                    '�ڕW�̌v�Z
                    Total_Times = 0
                    For i = 0 To 2
                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)))
                    Next i
                    
                    
                    
                    If CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) = 0 Then
                        AVE = 0
                    Else
                        AVE = Round(Total_Times / CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), 1)
                    End If
                                        
                    
                    Target = AVE * CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode))
                    lblTarget1.Caption = "����ڕW�F" & Format(Target, "#0") & "��"
                    
                    wkValue = Format(AVE, "#0.0")
                    If Right(wkValue, 1) = "0" Then
                        wkAVE = Format(AVE, "#0")
                    Else
                        wkAVE = Format(AVE, "#0.0")
                    End If
                    lblTarget2.Caption = wkAVE & "��/�~" & Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0") & "��"
                               
            
            
            
            End If
        
        Case BtErrEOF
            lblBEF_JISSEKI.Caption = ""
            lblBEF_BEFORE1.Caption = ""
            lblBEF_BEFORE2.Caption = ""
            lblBEF_BEFORE3.Caption = ""
            lblBEF_BEFORE4.Caption = ""
            lblBEF_SAGYO1.Caption = ""
            lblBEF_SAGYO2.Caption = ""
            lblBEF_SAGYO3.Caption = ""
            lblBEF_AFTER1.Caption = ""
            lblBEF_AFTER2.Caption = ""
            lblBEF_JISEKI.Caption = ""
            lblBEF_TASEKI.Caption = ""
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�w�}�[�ް��i�e�j")
            Exit Sub

    End Select


    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        lblBunnou.Visible = False
    
    
        lblTarget1.Visible = True
        lblTarget2.Visible = True
    
    
    
    
    
    Else
        lblBunnou.Visible = True
    
        lblTarget1.Visible = False
        lblTarget2.Visible = False
    
    
    End If

'    Doukon_com = BtOpGetGreater            '2013.03.31
    Doukon_com = BtOpGetGreaterEqual        '2013.03.31
    
    
    
    Doukon_eof = False

    Doukon_cnt = 0



End Sub

Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORPortrait
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 10
    Me.PageTopMargin = 10
    Me.PageLeftMargin = 20
    Me.PageRightMargin = 20

    Me.documentName = "���i���w�}�[�F"

    DoEvents

End Sub

