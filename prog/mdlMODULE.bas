Attribute VB_Name = "mdlMODULE"
Option Explicit


Public Function MODULE_INSPE_CHECK_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���W���[�����i�����x
'       2014.06.24
'2016.05.14 Private--> Public
'-------------------------------------------------------
Dim sts             As Integer



Dim Hinban          As String * 20

Dim Location        As String * 8

Dim i               As Integer
Dim j               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim MI_QTY          As Long
Dim SUMI_QTY        As Long
Dim Zaiko_QTY       As Long

Dim Use_QTY         As Long

Dim HANTEI_MARK     As String

Dim wkDate          As String * 10

Dim MENU_NO         As String * 2


    MODULE_INSPE_CHECK_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MODULE_INSPE_CHECK_PROC = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                
                        End Select
    
    
    
    
                        '---------------------  <�i�ԁi���W���[���j> --------------------------------------------
                        Call UniCode_Conv(K0_M_ITEM.JGYOBU, RET_JGYOBU)
                        Call UniCode_Conv(K0_M_ITEM.NAIGAI, RET_NAIGAI)
                        Call UniCode_Conv(K0_M_ITEM.HIN_GAI, Hinban)
                        sts = BTRV(BtOpGetEqual, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                            
                                
                                '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "���W���[���ΏۊO", "�~�p�����", "")
                                '
                                '
                                'Sendbuf = Text_Create_Proc()
                                'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                'MODULE_INSPE_CHECK_PROC = False
                                'Exit Function
                                
                                                            
                                ID_KANRI_TBL(ING_No).Hinban = Hinban
                                
                                ID_KANRI_TBL(ING_No).MEMO = "���W���[���i�ږ��o�^"
                                
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                            
                                Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "���W���[���i�ږ��o�^", "", "", Buzzer_DEF)
                                Sendbuf = Text_Create_Proc()
                            
                                MODULE_INSPE_CHECK_PROC = False
                                
                                Exit Function
                            
                                '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                            
                            
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^(���W���[��)", 0)
                                Exit Function
                        End Select
    
                        '---------------------  <���W���[���Ώۃ`�F�b�N>
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) <> "1" Then
                        
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "���W���[���ΏۊO", "�~�p�����", "")
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "���W���[���ΏۊO"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "���W���[���ΏۊO", "�~�p�����", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            
                            Exit Function
                            
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
         
                        End If
                        '---------------------  <�����ł��؂�`�F�b�N>
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "0" Then
                            
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "0��Ώ�", "�~�p�����", "")
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                        
                        
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "0��Ώ�"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "0��Ώ�", "�~�p�����", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            
                            Exit Function
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                        End If
                            
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "3" Then
                            
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "3�Ő؂�", "�~�p�����", "")
                            '
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "3�Ő؂�"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "3�Ő؂�", "�~�p�����", "", Buzzer_DEF)
                            
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            Exit Function
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                        
                        End If
                            
                        '---------------------  <���j�b�g�`�F�b�N>      2014.07.03 DELETE
                        'If StrConv(M_ITEM_REC.MODULE_UNIT_KBN, vbUnicode) = "2" Then
                        '
                        '    '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                        '    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "2���j�b�g�q", "�S���Ҋm�F", "")
                        '    '
                        '    '
                        '    '
                        '    'Sendbuf = Text_Create_Proc()
                        '    'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        '    'MODULE_INSPE_CHECK_PROC = False
                        '    'Exit Function
                        '
                        '
                        '    ID_KANRI_TBL(ING_No).Hinban = Hinban
                        '
                        '    ID_KANRI_TBL(ING_No).MEMO = "2���j�b�g�q"
                        '
                        '    ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        '
                        '    Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "2���j�b�g�q", "�S���Ҋm�F", "")
                        '    Sendbuf = Text_Create_Proc()
                        '
                        '    MODULE_INSPE_CHECK_PROC = False
                        '    Exit Function
                        '
                        '    '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                        'End If
                        '---------------------  <����`�F�b�N>
                        If StrConv(M_ITEM_REC.KENSA_JIGU, vbUnicode) = "1" Then
                            
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "����Ȃ�", "10�Ԃֈړ����", "")
                            '
                            '
                            '
                            'Sendbuf = Text_Create_Proc()
                            'ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            'MODULE_INSPE_CHECK_PROC = False
                            'Exit Function
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "����Ȃ�"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "����Ȃ�", "10�Ԃֈړ����", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            Exit Function
                            
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                        
                        End If
                        
                        '---------------------  <�ݕσ`�F�b�N>  2014.07.02
                        If StrConv(M_ITEM_REC.SETUHEN_KBN, vbUnicode) = "1" Then
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "�~�ݕϗL��"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�~�ݕϗL��", "", "", Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK_PROC = False
                            Exit Function
                            
                            '--------------------   �G���[���b�Z�[�W���m�F���b�Z�[�W�ɕύX  2014.07.01
                        
                        End If
                        
                        
                        '---------------------  <����>
                        '���݌�
                        
                        Zaiko_QTY = 0
                        For j = 0 To UBound(Nara_Soko_T)
                        
                            Location = Nara_Soko_T(j)
                        
                            'If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then     '2018.09.18
                            If NEW_SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then  '2018.09.18
                                Exit Function
                            End If
                            Zaiko_QTY = Zaiko_QTY + (SUMI_QTY + MI_QTY)
                        Next j
                    
                    
                        '4�����݌�
                        Use_QTY = Val(StrConv(M_ITEM_REC.HITUYO_SU, vbUnicode)) * Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))
                        
                        
                        HANTEI_MARK = "D �Đ����"
                        If Use_QTY >= 200 Then
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "A �N�x�m�F"
                            Else
                                HANTEI_MARK = "B �Đ����"
                            End If
                        Else
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "C �N�x�m�F"
                            Else
                                HANTEI_MARK = "D �Đ����"
                            End If
                        End If
        
    
    
                        'wkDate = Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 1, 4) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 5, 2) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 7, 2)
    
    
                        
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        ID_KANRI_TBL(ING_No).MEMO = HANTEI_MARK
                        

'               --- 2014.07.17
'                        Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
'                                                Hinban, _
'                                                "����:" & HANTEI_MARK, _
'                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#") & "�����݌�:" & Format(Use_QTY), _
'                                                "���݌�:" & Format(Zaiko_QTY))
                        
                        
                        Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                Hinban, _
                                                "����:" & HANTEI_MARK, _
                                                "����݌�:" & Format(Use_QTY), _
                                                "�� �� ��:" & Format(Zaiko_QTY), Buzzer_DEF)
                        
'               --- 2014.07.17
                        
                        Sendbuf = Text_Create_Proc()
                        
                        
                        '-----------------------------------------------�w�b�_�[
                        'Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        'ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                        '
                        'Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
                        'ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                        '
                        'Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        'ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                        '
                        'Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        'ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                        '
                        'Send_Text.FileName = ""                                 '���M�f�[�^�t�@�C����
                        'ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                        '
                        'Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        'ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        '
                        '-----------------------------------------------�P�s��
                        '                                                        'BOX����
                        'Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        '                                                        '�\�����e
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, "����:" & HANTEI_MARK)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, "����:" & HANTEI_MARK)
                        '                                                        '���l�����\��
                        'Send_Text.Box_Type(0).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                        '                                                        '�����J�[�\���ʒu
                        'Send_Text.Box_Type(0).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                        '                                                        '���͌���
                        'Send_Text.Box_Type(0).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                        '
                        'Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------�Q�s��
                        '                                                        'BOX����
                        'Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        '                                                        '�\�����e
                        'Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "���Đ����")
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "���Đ����")
                        '                                                        '���l�����\��
                        'Send_Text.Box_Type(1).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                        '                                                        '�����J�[�\���ʒu
                        'Send_Text.Box_Type(1).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                        '                                                        '���͌���
                        'Send_Text.Box_Type(1).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                        '
                        'Send_Text.Box_Type(1).MENU = ""                         '���j���\�ԍ�
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------�R�s��
                        '                                                        'BOX����
                        'Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        '                                                        '�\�����e
                        '
                        '
                        'wkDate = Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 1, 4) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 5, 2) & _
                        '        "/" & Mid(StrConv(M_ITEM_REC.SETUHEN_LAST_DATE, vbUnicode), 7, 2)
                        '
                        'Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�ݕ�:" & wkDate)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�ݕ�:" & wkDate)
                        '                                                        '���l�����\��
                        'Send_Text.Box_Type(2).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                        '                                                        '�����J�[�\���ʒu
                        'Send_Text.Box_Type(2).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                        '                                                        '���͌���
                        '
                        'Send_Text.Box_Type(2).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                        '
                        '
                        '
                        'Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
                        '                                                        'BOX����
                        'Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        '                                                        '�\�����e
                        'Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#") & "�����݌�:" & Format(Use_QTY))
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode)), "#") & "�����݌�:" & Format(Use_QTY))
                        '                                                        '���l�����\��
                        'Send_Text.Box_Type(3).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                        '                                                        '�����J�[�\���ʒu
                        'Send_Text.Box_Type(3).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                        '                                                        '���͌���
                        '
                        'Send_Text.Box_Type(3).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                        '
                        '
                        '
                        'Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�T�s��
                        '
                        '                                                        'BOX����
                        'Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        '                                                        '�\�����e
                        'Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "���݌�:" & Format(Zaiko_QTY))
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "���݌�:" & Format(Zaiko_QTY))
                        '                                                        '���l�����\��
                        'Send_Text.Box_Type(4).INIT = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                        '                                                        '�����J�[�\���ʒu
                        'Send_Text.Box_Type(4).Start_Pos = ""
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                        '                                                        '���͌���
                        '
                        'Send_Text.Box_Type(4).Max_Size = "00"
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                        '
                        'Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        'ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        '
                        '
                        'Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�iENT�j
                        
                        
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                            
                            
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                            
            If Trim(MENU_NO) = "" Then
            Else
            '���۸ޏo��
                
                If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    MENU_NO, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                     ID_KANRI_TBL(ING_No).Hinban, , , , , , , , , , , , , , _
                                                     ID_KANRI_TBL(ING_No).MEMO) Then
                    MODULE_INSPE_CHECK_PROC = SYS_ERR
                    GoTo Abort_Tran
                End If
            End If
                                
                                                        
'-----------------------------<�ړ����o��>













'-----------------------------<�ړ����o��>
                                                        
                                                        
                                        '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
            '���̍�Ɨv��
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- �G���[���b�Z�[�W�쐬
                Case Else
                '�d�v�ȗv���Ȃ̂Ŗ��o�^�̓V�X�e����~�Ƃ���
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "�v���}�X�^", 0)
                    Exit Function
            End Select
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        
            Sendbuf = Text_Create_Proc()
        
        
            MODULE_INSPE_CHECK_PROC = False

            Exit Function
                        
        
    End Select
            
    MODULE_INSPE_CHECK_PROC = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Public Function MODULE_INSPE_CHECK2_PROC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���W���[�����i����2�x
'       2015.02.19
'2016.05.14 Private--> Public
'
'-------------------------------------------------------
Dim sts             As Integer



Dim Hinban          As String * 20

Dim Location        As String * 8

Dim i               As Integer
Dim j               As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim MI_QTY          As Long
Dim SUMI_QTY        As Long
Dim Zaiko_QTY       As Long

Dim Use_QTY         As Long

Dim HANTEI_MARK     As String

Dim wkDate          As String * 10

Dim MENU_NO         As String * 2


    MODULE_INSPE_CHECK2_PROC = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MODULE_INSPE_CHECK2_PROC = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                
                        End Select
    
    
    
    
                        '---------------------  <�i�ԁi���W���[���j> --------------------------------------------
                        Call UniCode_Conv(K0_M_ITEM.JGYOBU, RET_JGYOBU)
                        Call UniCode_Conv(K0_M_ITEM.NAIGAI, RET_NAIGAI)
                        Call UniCode_Conv(K0_M_ITEM.HIN_GAI, Hinban)
                        sts = BTRV(BtOpGetEqual, M_ITEM_POS, M_ITEM_REC, Len(M_ITEM_REC), K0_M_ITEM, Len(K0_M_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                
                                                            
                                ID_KANRI_TBL(ING_No).Hinban = Hinban
                                
                                ID_KANRI_TBL(ING_No).MEMO = "���W���[���i�ږ��o�^"
                                
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                            
                                Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "���W���[���i�ږ��o�^", "", "", Buzzer_DEF)
                                Sendbuf = Text_Create_Proc()
                            
                                MODULE_INSPE_CHECK2_PROC = False
                                
                                Exit Function
                            
                            
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^(���W���[��)", 0)
                                Exit Function
                        End Select
    
                        '---------------------  <�݌ɏW�v>
    
                        '���݌�
                        
                        Zaiko_QTY = 0
                        For j = 0 To UBound(Nara_Soko_T)
                        
                            Location = Nara_Soko_T(j)
                        
                            'If SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then     '2018.09.18
                            If NEW_SOKO_Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, RET_JGYOBU, RET_NAIGAI, Hinban, Location) Then  '2018.09.18
                                Exit Function
                            End If
                            Zaiko_QTY = Zaiko_QTY + (SUMI_QTY + MI_QTY)
                        Next j
                    
                    
                        '4�����݌�
                        Use_QTY = Val(StrConv(M_ITEM_REC.HITUYO_SU, vbUnicode)) * Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))
    
    
    
                        '---------------------  <���W���[���Ώۃ`�F�b�N>
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) = "0" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "���W���[���ΏۊO"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "���W���[���ΏۊO", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DEF)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                            
                        End If
                        
                        
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) = "9" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "���U����"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "���U����", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                            
                        End If
                        
                        '>>>>>>>>>  2017.11.24
                        If StrConv(M_ITEM_REC.MODULE_KBN, vbUnicode) = "8" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "�S���c��"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "�S���c��", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                            
                        End If
                        
                        
                        '---------------------  <�����ł��؂�`�F�b�N>
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "0" Then
                        
                        
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "0��Ώ�"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "0��Ώ�", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DEF)
                            
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            
                            Exit Function
                        End If
                            
                        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) = "3" Then
                            
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "3�Ő؂�"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "3�Ő؂�", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            Exit Function
                        
                        End If
                            
                        '---------------------  <����`�F�b�N>
                        If StrConv(M_ITEM_REC.KENSA_JIGU, vbUnicode) = "1" Then
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "����Ȃ�"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "����Ȃ�", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DEF)
                        
                            Sendbuf = Text_Create_Proc()
                        
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            Exit Function
                            
                        
                        End If
                        
                        '---------------------  <�ݕσ`�F�b�N>  2014.07.02
                        If StrConv(M_ITEM_REC.SETUHEN_KBN, vbUnicode) = "1" Then
                            
                            ID_KANRI_TBL(ING_No).Hinban = Hinban
                            
                            ID_KANRI_TBL(ING_No).MEMO = "�~�ݕϗL��"
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                        
                                                        
                            Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                                                Hinban, _
                                                                                "�~�ݕϗL��", _
                                                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DOUBLE)
                                                        
                            Sendbuf = Text_Create_Proc()
                        
                            MODULE_INSPE_CHECK2_PROC = False
                            Exit Function
                            
                        
                        End If
                        
                        
                        '---------------------  <����>
                        
                        
                        HANTEI_MARK = "D �Đ����"
                        If Use_QTY >= 200 Then
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "A �N�x�m�F"
                            Else
                                HANTEI_MARK = "B �Đ����"
                            End If
                        Else
                            If Zaiko_QTY >= Use_QTY Then
                                HANTEI_MARK = "C �N�x�m�F"
                            Else
                                HANTEI_MARK = "D �Đ����"
                            End If
                        End If
        
                        
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        
                        ID_KANRI_TBL(ING_No).MEMO = HANTEI_MARK
                        

                        
                        
                        Call MODULE_TEXT_PROC(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, _
                                                Hinban, _
                                                "����:" & HANTEI_MARK, _
                                                Format(Val(StrConv(M_ITEM_REC.HITUYO_TUKI, vbUnicode))) & "�����݌�:" & Format(Use_QTY), _
                                                "���݌�:" & Format(Zaiko_QTY), Buzzer_DEF)
                        
                        
                        Sendbuf = Text_Create_Proc()
                        
                        
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�iENT�j
                        
                        
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                            
                            
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                            
            If Trim(MENU_NO) = "" Then
            Else
            '���۸ޏo��
                
                If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    MENU_NO, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                     ID_KANRI_TBL(ING_No).Hinban, , , , , , , , , , , , , , _
                                                     ID_KANRI_TBL(ING_No).MEMO) Then
                    MODULE_INSPE_CHECK2_PROC = SYS_ERR
                    GoTo Abort_Tran
                End If
            End If
                                
                                                        

'-----------------------------<�ړ����o��>
                                                        
                                                        
                                        '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
            '���̍�Ɨv��
            Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
            Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
            sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                '   -------------------------------- �G���[���b�Z�[�W�쐬
                Case Else
                '�d�v�ȗv���Ȃ̂Ŗ��o�^�̓V�X�e����~�Ƃ���
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual, "�v���}�X�^", 0)
                    Exit Function
            End Select
            
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        
            Sendbuf = Text_Create_Proc()
        
        
            MODULE_INSPE_CHECK2_PROC = False

            Exit Function
                        
        
    End Select
            
    MODULE_INSPE_CHECK2_PROC = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Private Sub MODULE_TEXT_PROC(Line1 As String, Line2 As String, Line3 As String, Line4 As String, LINE5 As String, buzzer As String)
'-------------------------------------------------------
'
'   �w���W���[�����i�����@�m�F÷�č쐬�x
'       2014.06.24
'
'-------------------------------------------------------
                        
    '-----------------------------------------------�w�b�_�[
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK

    Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF

    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only

    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"

    Send_Text.FileName = ""                                 '���M�f�[�^�t�@�C����
    ID_KANRI_TBL(ING_No).Send_Text.FileName = ""

'    Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    Send_Text.buzzer = buzzer                               '�u�U�[���@�W��
'    ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
    ID_KANRI_TBL(ING_No).Send_Text.buzzer = buzzer
                        
    '-----------------------------------------------�P�s��
                                                            'BOX����
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, Line1)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, Line1)
                                                            '���l�����\��
    Send_Text.Box_Type(0).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(0).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                            '���͌���
    Send_Text.Box_Type(0).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                            
    Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
    '-----------------------------------------------�Q�s��
                                                            'BOX����
    Send_Text.Box_Type(1).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Line2)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Line2)
                                                            '���l�����\��
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(1).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                            '���͌���
    Send_Text.Box_Type(1).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                            
    Send_Text.Box_Type(1).MENU = ""                         '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '-----------------------------------------------�R�s��
                                                            'BOX����
    Send_Text.Box_Type(2).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '�\�����e

    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Line3)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Line3)
                                                            '���l�����\��
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(2).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                            '���͌���
                                                            
    Send_Text.Box_Type(2).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                            
                                                            
                                                            
    Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
    '-----------------------------------------------�S�s��
                                                            'BOX����
    Send_Text.Box_Type(3).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Line4)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Line4)
                                                            '���l�����\��
    Send_Text.Box_Type(3).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(3).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                            '���͌���
                                                            
    Send_Text.Box_Type(3).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                            
                                                            
                                                            
    Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
    '-----------------------------------------------�T�s��
    
                                                            'BOX����
    Send_Text.Box_Type(4).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LINE5)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LINE5)
                                                            '���l�����\��
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(4).Start_Pos = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                            '���͌���
                                                            
    Send_Text.Box_Type(4).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
    
    Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

End Sub


Public Function Module_In_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���W���[���i�ԓ��ɏ����̃`�F�b�N���X�V�����x
'
'
'       2018.10.03
'-------------------------------------------------------
Dim i           As Integer

Dim Hinban      As String * 20


Dim Tanaban     As String * 8
Dim sts         As Integer

Dim QTY         As Long
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1

Dim MENU_NO     As String * 2

Dim WK_CODE     As String * 5       '2007.05.28
Dim WK_TANKA    As String * 11      '2007.05.28


    Module_In_Proc = True
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                            
                            
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                        '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.SOKO_NO, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")    '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                    Module_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ ���ڃ`�F�b�N
                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")  '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Module_In_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.SOKO_NO, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")    '2017.09.22
                            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                    Module_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
                    
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                    
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")    '2017.09.22
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                
                                Module_In_Proc = False
                                Exit Function
                            End If
                    
                    
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        '------------------ �i�ڃ}�X�^�Ǎ���
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                            '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.SOKO_NO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.SOKO_NO, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")    '2017.09.22
                            
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                            Module_In_Proc = False
                                            Exit Function
                                        Case Else
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                            Exit Function
                                    End Select
                                End If
                            Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                Module_In_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
                    
                    
                    
                    
                        '�݌ɏW�v
'                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
'                        Select Case sts
'                            Case False
'                            Case True           '�����ł͔������Ȃ�
'                                Exit Function
'                            Case SYS_ERR
'                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                Exit Function
'                            Case SYS_CANCEL
'                                Call Err_Send_Proc("�݌Ɏg�p��", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                Module_In_Proc = False
'                                Exit Function
'                        End Select
        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        
                        ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '�I�Ԃ��Z�[�u
                        ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
                        ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
                        ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i�̐���
                                                                        
                        ID_KANRI_TBL(ING_No).RET_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�
                        ID_KANRI_TBL(ING_No).RET_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O
            
            
            
            
                        '���ʕt���̑��M���b�Z�[�W���쐬����
                        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.FileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                                                                                    
                        '-----------------------------------------------�P�s��
                                                                        'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                                                        '���l�����\��
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                        
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                        '���͌���
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(0).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------�Q�s��
                                                                        'BOX����
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
            
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                                Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                                        '���l�����\��
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(1).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                        '���͌���
                        Send_Text.Box_Type(1).Max_Size = "08"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                            
                        Send_Text.Box_Type(1).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------�R�s��
                                                                        'BOX����
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
                                                                         '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                       '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                        '���͌���
                        Send_Text.Box_Type(2).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"
                                                                                            
                        Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                        'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_MI_Suryo)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_MI_Suryo)
                                                                        '���l�����\��
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '���͌���
                        Send_Text.Box_Type(3).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�T�s��
                                                                        'BOX����
                    
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                        '�\�����e
'                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�݌ɐ��F" & Format(MI_QTY + SUMI_QTY, "#0"))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�݌ɐ��F" & Format(MI_QTY + SUMI_QTY, "#0"))
                                                                        
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                        
                                                                        
                                                                        '���l�����\��
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                        '���͌���
                        Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                    
                        Sendbuf = Text_Create_Proc()
                    
                End Select
            
            Next i
            
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���ʁj
            
            For i = 0 To M_Gyo - 1
                
                Select Case i
            
                    
                    Case 3
            
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Module_In_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Module_In_Proc = False
                            Exit Function
                        End If
            
                '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                    
                                                            '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                        
                                        
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                                    
                            MENU_NO = ""
                        End If
                                        
                                        
                       If RET_JGYOBU = SHIZAI Then
                           Call UniCode_Conv(K0_ITEM.JGYOBU, RET_JGYOBU)
                           Call UniCode_Conv(K0_ITEM.NAIGAI, RET_NAIGAI)
                           Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
                           sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                           Select Case sts
                               Case BtNoErr
                               Case Else
                                   Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                   Sendbuf = Text_Create_Proc()
                                   Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                   Exit Function
                           End Select
                       
                           '�i��Ͻ��̍ŐV�d����^�P�����ݒ肳��Ă������́A������̍��ڂ��g�p  2007.05.28
                           If Not IsNumeric(StrConv(ITEMREC.LAST_TANKA, vbUnicode)) Then
                               
                               WK_CODE = StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode)
                               WK_TANKA = StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode)
                           Else
                               WK_CODE = StrConv(ITEMREC.LAST_CODE, vbUnicode)
                               WK_TANKA = StrConv(ITEMREC.LAST_TANKA, vbUnicode)
                           
                           End If
                       
                       
                       
                           sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).RET_JGYOBU, _
                                                   ID_KANRI_TBL(ING_No).RET_NAIGAI, _
                                                   ID_KANRI_TBL(ING_No).Hinban, _
                                                   Format(Now, "YYYYMMDD"), _
                                                   ID_KANRI_TBL(ING_No).Tanaban, _
                                                   (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                   0, _
                                                   QTY, _
                                                   Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                                   ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                   FILE_RETRY, , _
                                                   WK_CODE, _
                                                   WK_TANKA, , _
                                                   MENU_NO)
                    
                           Select Case sts
                               Case False
                               Case True           '���Ɏ��͔������Ȃ�
                               Case SYS_CANCEL
                                   'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")        '2017.09.22
                                   Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "�������f", "", "", "")    '2017.09.22
                                   Sendbuf = Text_Create_Proc()
                                   ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                   
                                   Module_In_Proc = False
                                   GoTo Abort_Tran
                               Case SYS_ERR
                                   Sendbuf = Text_Create_Proc()
                                   Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                   Module_In_Proc = SYS_ERR    '�V�X�e���ُ픭��
                                   
                                   GoTo Abort_Tran
                           End Select
                       
                       
                       
                       
                       
                       
                       Else
                                                           
                                                               
                                                               '���ɍX�V
                           sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).RET_JGYOBU, _
                                                   ID_KANRI_TBL(ING_No).RET_NAIGAI, _
                                                   ID_KANRI_TBL(ING_No).Hinban, _
                                                   Format(Now, "YYYYMMDD"), _
                                                   ID_KANRI_TBL(ING_No).Tanaban, _
                                                   (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                   0, _
                                                   QTY, _
                                                   Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                                   ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                   FILE_RETRY, , , , , MENU_NO)
                           Select Case sts
                               Case False
                               Case True           '���Ɏ��͔������Ȃ�
                               Case SYS_CANCEL
                                   'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")        '2017.09.22
                                   Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, "�������f", "", "", "")    '2017.09.22
                                   Sendbuf = Text_Create_Proc()
                                   ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                   
                                   Module_In_Proc = False
                                   GoTo Abort_Tran
                               Case SYS_ERR
                                   Sendbuf = Text_Create_Proc()
                                   Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                   Module_In_Proc = SYS_ERR    '�V�X�e���ُ픭��
                                   
                                   GoTo Abort_Tran
                           End Select
                       End If
                
                                                            '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                
                                
                                
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        
                        ID_KANRI_TBL(ING_No).Inp_QTY = QTY          '���͐���
                                
                        '�݌ɕt���̑��M���b�Z�[�W���쐬����
                        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.FileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                                
                                
                        '�݌ɏW�v
                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, ID_KANRI_TBL(ING_No).Tanaban, ID_KANRI_TBL(ING_No).RET_JGYOBU, ID_KANRI_TBL(ING_No).RET_NAIGAI, ID_KANRI_TBL(ING_No).Hinban, SUMI_QTY, MI_QTY)
                        Select Case sts
                            Case False
                            Case True           '�����ł͔������Ȃ�
                                Exit Function
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc("�݌Ɏg�p��", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Module_In_Proc = False
                                Exit Function
                        End Select
                                
                                
                                
                                
                        '-----------------------------------------------�P�s��
                                                                        'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                                                        '���l�����\��
                        Send_Text.Box_Type(0).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                                        
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(0).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                        '���͌���
                        Send_Text.Box_Type(0).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(0).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------�Q�s��
                                                                        'BOX����
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, _
                                Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2))
            
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                                Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2))
                                                                        '���l�����\��
                        Send_Text.Box_Type(1).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(1).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                                        '���͌���
                        Send_Text.Box_Type(1).Max_Size = "08"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "08"
                                                                                            
                        Send_Text.Box_Type(1).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
                        '-----------------------------------------------�R�s��
                                                                        'BOX����
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).Hinban)
                                                                         '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                       '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                        '���͌���
                        Send_Text.Box_Type(2).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"
                                                                                            
                        Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                        'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�����i�F" & QTY)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "�����i�F" & QTY)
                                                                        '���l�����\��
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '���͌���
                        Send_Text.Box_Type(3).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�T�s��
                                                                        'BOX����
                    
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�݌ɐ��F" & Format(MI_QTY + SUMI_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�݌ɐ��F" & Format(MI_QTY + SUMI_QTY, "#0"))
                                                                        '���l�����\��
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                        '���͌���
                        Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        
                        
                        
                        Sendbuf = Text_Create_Proc()
                    
                    
                    
                    End Select
    
    
                Next i
    
    
    
            Case Step_Sagyo3_RES        '�R��ڂ̎�M�iENT�j
                    
                '���̍�Ɨv��
                Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
                Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                Select Case sts
                    Case BtNoErr
                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                    Case Else
                    '�d�v�ȗv���Ȃ̂Ŗ��o�^�̓V�X�e����~�Ƃ���
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�v���}�X�^", 0)
                        Exit Function
                End Select
                
                ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                
                If Sagyo_Send_Proc() Then
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                End If
            
                Sendbuf = Text_Create_Proc()
            
            
                Module_In_Proc = False
    
                Exit Function
    
    
    End Select
    
    Module_In_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


