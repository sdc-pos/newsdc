Attribute VB_Name = "mdlProc3"
Option Explicit


Public Function NYUKO_KENPIN_OSAKA_S_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���o�b�@���ތ������Ɂx
'
'   2012.03.06
'
'-------------------------------------------------------
Dim i               As Integer


Dim Hinban          As String * 20
Dim Tanaban         As String * 8
Dim QTY             As Long

Dim sts             As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2


Dim wkORDER_QTY     As Long
Dim wkNYUKO_QTY     As Long
Dim wkQty           As Long

Dim ST_TANABAN      As String * 11


    NYUKO_KENPIN_OSAKA_S_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁj

            For i = 0 To M_Gyo - 1
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI, , BUZAI)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                NYUKO_KENPIN_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            ST_TANABAN = ""
                        Else
                            ST_TANABAN = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                        
                        
                        
                        '�����d���\��c�W�v
                        If ORDER_ZAN_Proc(RET_JGYOBU, NAIGAI_NAI, Hinban, wkORDER_QTY, wkNYUKO_QTY) Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
                        
                        ID_KANRI_TBL(ING_No).ORDER_QTY = wkORDER_QTY
                        ID_KANRI_TBL(ING_No).NYUKO_QTY = wkNYUKO_QTY
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
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
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>>>    2017.09.22
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>>>    2017.09.22
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                                                                                
                                                                                
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "16"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "16"
                                                                                
                                                                                
                                                                                
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                        'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                        '���l�����\��
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                        '���͌���
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�T�s��
                                                                        'BOX����
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                        '�\�����e
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
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                    
                    
                    

                        Sendbuf = Text_Create_Proc()
                        
                        Exit Function
                
                End Select
            Next i
    
    
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�I�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        '------------------ �q�Ƀ}�X�^�Ǎ���
                        Call UniCode_Conv(K0_SOKO.SOKO_NO, Left(Tanaban, 2))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                    
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                NYUKO_KENPIN_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                Exit Function
                        End Select
                        '------------------ ���ڃ`�F�b�N
                        If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                            If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).S_JGYOBU Or _
                                StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")          '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")      '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                NYUKO_KENPIN_OSAKA_S_Proc = False
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
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")            '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")        '2017.09.22
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                NYUKO_KENPIN_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                Exit Function
                        End Select
                    
                        '------------------ �֎~�I�̃`�F�b�N
                        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")            '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")        '2017.09.22
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                        ID_KANRI_TBL(ING_No).Tanaban = Tanaban
    
    
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
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
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
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
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------�Q�s��
                                                                                
                                                                                'BOX����
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "���c(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "���c(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                        '�\�����e
                        
                        wkQty = ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY
                        If wkQty < 0 Then
                            wkQty = 0
                        End If
                        
                        
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Suryo)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Suryo)
                                                                        '���l�����\��
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '���͌���
                        Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
                        
                        NYUKO_KENPIN_OSAKA_S_Proc = False
                        Exit Function
                End Select
            Next i
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i���ʁj
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_Suryo          '����
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                        If QTY <= (ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY) Then
                        
                        
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
                                                                '���ɍX�V
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
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
                                    
                                    NYUKO_KENPIN_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    NYUKO_KENPIN_OSAKA_S_Proc = SYS_ERR    '�V�X�e���ُ픭��
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        
                                                    
                        Else
                        
                            ID_KANRI_TBL(ING_No).INPUT_QTY = QTY
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
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
                    
                            Send_Text.Buzzer = Buzzer_DOUBLE                        '�u�U�[���@��d�x����
                            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DOUBLE
                            
                            '-----------------------------------------------�P�s��
                                                                                    'BOX����
'                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                                                                                    '�\�����e
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                                                                                    '���l�����\��
'                            Send_Text.Box_Type(0).INIT = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
'                                                                                    '�����J�[�\���ʒu
'                            Send_Text.Box_Type(0).Start_Pos = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
'                                                                                    '���͌���
'                            Send_Text.Box_Type(0).Max_Size = "00"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
'
'                            Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                            '-----------------------------------------------�Q�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�����c:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�����c:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�����ɐ�:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�����ɐ�:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
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
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�I������s���܂����H")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�I������s���܂����H")
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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                                                                                    '���l�����\��
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(4).Start_Pos = "20"                  '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "20"
                                                                                    '���͌���
                             Send_Text.Box_Type(4).Max_Size = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "01"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            Sendbuf = Text_Create_Proc()
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            
                            Exit Function
                        End If
                
                End Select
            Next i
    
    
        Case Step_Sagyo4_RES        '�S��ڂ̎�M�i�����p��(Y/N)�j
            
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_CAN_ANS          '����
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "1" And Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "9" Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 �����", "���ĉ������B", "")          '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 �����", "���ĉ������B", "")      '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) = "1" Then
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
                                                                '���ɍX�V
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                    ID_KANRI_TBL(ING_No).Hinban, _
                                                    Format(Now, "YYYYMMDD"), _
                                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    ID_KANRI_TBL(ING_No).INPUT_QTY, _
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
                                    
                                    NYUKO_KENPIN_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    NYUKO_KENPIN_OSAKA_S_Proc = SYS_ERR    '�V�X�e���ُ픭��
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        Else
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
                            
                            
                            
                            NYUKO_KENPIN_OSAKA_S_Proc = False
                            
                            Exit Function
                        
                        End If
            
                End Select
            Next i
    End Select

End_Tran:
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
    
    
    
    NYUKO_KENPIN_OSAKA_S_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Public Function ORDER_ZAN_Proc(JGYOBU As String, NAIGAI As String, Hinban As String, ORDER_QTY As Long, NYUKO_QTY As Long) As Integer
'-------------------------------------------------------
'
'   �w���o�b�@���ޒ����c�W�v�x
'
'   2012.03.06
'
'-------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
    
    
Dim i           As Integer
    
Dim wkJgyobu    As String * 1
    
    
    ORDER_ZAN_Proc = True
    
    If JGYOBU = BUZAI Then
        wkJgyobu = SHIZAI
    Else
        wkJgyobu = JGYOBU
    End If
    
    Call UniCode_Conv(K7_P_SHORDER.USE_YM, USE_YM)
    Call UniCode_Conv(K7_P_SHORDER.JGYOBU, wkJgyobu)
    Call UniCode_Conv(K7_P_SHORDER.NAIGAI, NAIGAI)
    Call UniCode_Conv(K7_P_SHORDER.HIN_GAI, Hinban)
    Call UniCode_Conv(K7_P_SHORDER.CANCEL_F, P_CANCEL_OFF)
    
    ORDER_QTY = 0
    com = BtOpGetGreaterEqual
    
        
    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(com, P_SHORDER_POS, P_SHORDER_REC, Len(P_SHORDER_REC), K7_P_SHORDER, Len(K7_P_SHORDER), 7)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_SHORDER_REC.USE_YM, vbUnicode) <> USE_YM Or _
                    StrConv(P_SHORDER_REC.JGYOBU, vbUnicode) <> wkJgyobu Or _
                    StrConv(P_SHORDER_REC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    RTrim(StrConv(P_SHORDER_REC.HIN_GAI, vbUnicode)) <> RTrim(Hinban) Or _
                    StrConv(P_SHORDER_REC.CANCEL_F, vbUnicode) <> P_CANCEL_OFF Then
                       
                    Exit Do
                    
                End If
            
            Case BtErrEOF
                Exit Do
            
            '   -------------------------------- �G���[���b�Z�[�W�쐬
            Case Else
                '�d�v�ȗv���Ȃ̂Ŗ��o�^�̓V�X�e����~�Ƃ���
                Call File_Error(sts, BtOpGetEqual, "���ޒ����ް�", 0)
                Exit Function
        End Select
    
        ORDER_QTY = ORDER_QTY + (Val(StrConv(P_SHORDER_REC.ORDER_QTY, vbUnicode)) - Val(StrConv(P_SHORDER_REC.UKEIRE_QTY, vbUnicode)))
    
        com = BtOpGetNext
    
    Loop
    
    Call UniCode_Conv(K1_IDO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
    Call UniCode_Conv(K1_IDO.JITU_DT, BUZAI_DATE_S)
    Call UniCode_Conv(K1_IDO.JITU_TM, "")
    NYUKO_QTY = 0
    com = BtOpGetGreaterEqual
    
    Do
        
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        Select Case sts
            Case BtNoErr
            
                If StrConv(IDOREC.JGYOBU, vbUnicode) <> JGYOBU Or _
                    StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    RTrim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> RTrim(Hinban) Or _
                    StrConv(IDOREC.JITU_DT, vbUnicode) > BUZAI_DATE_E Then
                       
                    Exit Do
                    
                End If
            
            Case BtErrEOF
                Exit Do
            
            '   -------------------------------- �G���[���b�Z�[�W�쐬
            Case Else
                '�d�v�ȗv���Ȃ̂Ŗ��o�^�̓V�X�e����~�Ƃ���
                Call File_Error(sts, BtOpGetEqual, "�݌Ɉړ���", 0)
                Exit Function
        End Select
    
    
    
            
    
        For i = 0 To UBound(IN_TANA_S_OSAKA)
        
            If StrConv(IDOREC.RIRK_ID, vbUnicode) = IN_TANA_S_OSAKA(i) Then
            
                NYUKO_QTY = NYUKO_QTY + (Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                    
            
            End If
        
        Next i
        
        com = BtOpGetNext
    
    Loop
    
    
    
    
    ORDER_ZAN_Proc = False

End Function


Public Function NYUKO_MAEGARI_OSAKA_S_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���o�b�@���ތ������Ɂx
'
'   2012.03.06
'
'-------------------------------------------------------
Dim i               As Integer


Dim Hinban          As String * 20
Dim Tanaban         As String * 8
Dim QTY             As Long

Dim sts             As Integer


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2


Dim wkORDER_QTY     As Long
Dim wkNYUKO_QTY     As Long
Dim wkQty           As Long

Dim ST_TANABAN      As String * 11


    NYUKO_MAEGARI_OSAKA_S_Proc = True


    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁj

            For i = 0 To M_Gyo - 1
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI, , SHIZAI)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                            
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
                        
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                            ST_TANABAN = ""
                        Else
                            ST_TANABAN = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                        End If
                        
                        
                        
                        '�����d���\��c�W�v
                        If ORDER_ZAN_Proc(RET_JGYOBU, NAIGAI_NAI, Hinban, wkORDER_QTY, wkNYUKO_QTY) Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
                        
                        ID_KANRI_TBL(ING_No).ORDER_QTY = wkORDER_QTY
                        ID_KANRI_TBL(ING_No).NYUKO_QTY = wkNYUKO_QTY
                        
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
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
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>   2017.09.22
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>   2017.09.22
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban & " " & ST_TANABAN)
                                                                                
                                                                                
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "16"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "16"
                                                                                
                                                                                
                                                                                
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                        'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                        '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                        '���l�����\��
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(3).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                        '���͌���
                        Send_Text.Box_Type(3).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                            
                        Send_Text.Box_Type(3).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�T�s��
                                                                        'BOX����
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                        '�\�����e
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
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                    
                    
                    

                        Sendbuf = Text_Create_Proc()
                        
                        Exit Function
                
                End Select
            Next i
    
    
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�I�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        '------------------ �q�Ƀ}�X�^�Ǎ���
                        Call UniCode_Conv(K0_SOKO.SOKO_NO, Left(Tanaban, 2))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                    
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                Exit Function
                        End Select
                        '------------------ ���ڃ`�F�b�N
                        If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                            If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).S_JGYOBU Or _
                                StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")          '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")      '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
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
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")    '2017.09.22
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                NYUKO_MAEGARI_OSAKA_S_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                Exit Function
                        End Select
                    
                        '------------------ �֎~�I�̃`�F�b�N
                        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")            '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")        '2017.09.22
                            
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                        ID_KANRI_TBL(ING_No).Tanaban = Tanaban
    
    
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
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
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
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
                                                                                
                        Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                        '-----------------------------------------------�Q�s��
                                                                                
                                                                                'BOX����
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))

                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, _
                            Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "���c(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "���c(" & USE_YM & "):" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                        '�\�����e
                        
                        wkQty = ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY
                        If wkQty < 0 Then
                            wkQty = 0
                        End If
                        
                        
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Suryo)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Suryo)
                                                                        '���l�����\��
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(wkQty, "#0"))) & Format(wkQty, "#0")
                                                                        '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                        '���͌���
                        Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                            
                        Send_Text.Box_Type(4).MENU = ""                     '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
                        
                        NYUKO_MAEGARI_OSAKA_S_Proc = False
                        Exit Function
                End Select
            Next i
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i���ʁj
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_Suryo          '����
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                        If QTY <= (ID_KANRI_TBL(ING_No).ORDER_QTY - ID_KANRI_TBL(ING_No).NYUKO_QTY) Then
                        
                        
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
                                                                '���ɍX�V
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
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
                                    
                                    NYUKO_MAEGARI_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    NYUKO_MAEGARI_OSAKA_S_Proc = SYS_ERR    '�V�X�e���ُ픭��
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        
                                                    
                        Else
                        
                            ID_KANRI_TBL(ING_No).INPUT_QTY = QTY
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
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
                    
                            Send_Text.Buzzer = Buzzer_DOUBLE                        '�u�U�[���@��d�x����
                            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DOUBLE
                            
                            '-----------------------------------------------�P�s��
                                                                                    'BOX����
'                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                                                                                    '�\�����e
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
'                                                                                    '���l�����\��
'                            Send_Text.Box_Type(0).INIT = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
'                                                                                    '�����J�[�\���ʒu
'                            Send_Text.Box_Type(0).Start_Pos = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
'                                                                                    '���͌���
'                            Send_Text.Box_Type(0).Max_Size = "00"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
'
'                            Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
                            '-----------------------------------------------�Q�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).Hinban)
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�����c:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�����c:" & Format(ID_KANRI_TBL(ING_No).ORDER_QTY, "#0"))
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�����ɐ�:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�����ɐ�:" & Format(ID_KANRI_TBL(ING_No).NYUKO_QTY + QTY, "#0"))
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
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�I������s���܂����H")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�I������s���܂����H")
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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_CAN_ANS)
                                                                                    '���l�����\��
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(4).Start_Pos = "20"                  '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "20"
                                                                                    '���͌���
                             Send_Text.Box_Type(4).Max_Size = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "01"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            Sendbuf = Text_Create_Proc()
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            
                            Exit Function
                        End If
                
                End Select
            Next i
    
    
        Case Step_Sagyo4_RES        '�S��ڂ̎�M�i�����p��(Y/N)�j
            
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
            
            
                    Case LCD_CAN_ANS          '����
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "1" And Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) <> "9" Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 �����", "���ĉ������B", "")          '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "1 OR 9 �����", "���ĉ������B", "")      '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            Exit Function
                        End If
                
                
                
                        If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) = "1" Then
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
                                                                '���ɍX�V
                            sts = Nyuko_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                                    ID_KANRI_TBL(ING_No).Hinban, _
                                                    Format(Now, "YYYYMMDD"), _
                                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    ID_KANRI_TBL(ING_No).INPUT_QTY, _
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
                                    
                                    NYUKO_MAEGARI_OSAKA_S_Proc = False
                                    GoTo Abort_Tran
                                Case SYS_ERR
                                    Sendbuf = Text_Create_Proc()
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    NYUKO_MAEGARI_OSAKA_S_Proc = SYS_ERR    '�V�X�e���ُ픭��
                                    
                                    GoTo Abort_Tran
                            End Select
                                        
                            GoTo End_Tran
                        Else
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
                            
                            
                            
                            NYUKO_MAEGARI_OSAKA_S_Proc = False
                            
                            Exit Function
                        
                        End If
            
                End Select
            Next i
    End Select

End_Tran:
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
    
    
    
    NYUKO_MAEGARI_OSAKA_S_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


