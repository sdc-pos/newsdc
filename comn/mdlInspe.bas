Attribute VB_Name = "mdlInspe"
Option Explicit

Public Function Inspe_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 20      '2007.07.21 13-->20
Dim SYUKA_QTY       As Long
Dim JITU_QTY        As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim NAIGAI          As String * 1
Dim HIN_NO          As String * 13


Dim KAN_FLG         As String * 1

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim wkTEXT          As String           '2007.07.21

Dim MENU_NO         As String * 2

Dim wkNull_Check    As String       '2009.04.27

Dim Mod_Return      As Integer      '2010.12.09
Dim iNum            As Integer      '2010.12.09

Dim wkKENPIN_YMD    As String * 8   '���i���t   2017.09.07


    Inspe_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�`�[�h�c�j
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No      '�`�[�h�c
    
    
'>>>>   2017.09.07
'                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
'
'                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���i��ƕs��", "", "")
'                            Sendbuf = Text_Create_Proc()
'                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                            Inspe_Proc = False
'                            Exit Function
'                        End If
'>>>> 2017.09.07
    
    
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                NAIGAI, _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN, _
                                                JITU_QTY, , wkKENPIN_YMD)
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")         '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\�薳��", "", "")     '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")           '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc = False
                                Exit Function
                        End Select
                
'>>>>> 2017.09.07
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���i��ƕs��", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���i��ƕs��", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
'>>>>> 2017.09.07
                
                
                
                        '------------------ ������̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "������G���[", "", "")     '2017.09.22
                            '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "������G���[", "", "") '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�א�G���[", MTS_CODE, "") '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ �����敪�̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            
                            
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�����敪�~�X", "", "")     '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�����敪�G���[", "", "") '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ �o�Ɋ����̃`�F�b�N
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                
                                
                                
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "��Ɩ�����", "", "")   '2017.09.07
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ɖ�����", "", "")    '2017.09.07        2017.09.22 DEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ɖ�����", "", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc = False
                                Exit Function
                            End If
                        End If
                        
                        
                        
                        '------------------ �Č��i�̃`�F�b�N    2017.09.07
                        If Inspection_CHK = 1 Then
                            If Trim(wkKENPIN_YMD) <> "" Then
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "���i�ςł�", "", "")       '2017.09.22
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "���i�ςł�", "", "")
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).MTS_CODE, ID_NO, Hinban, "�o�א�:" & SYUKA_QTY, "���i�ςł�")
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc = False
                                Exit Function
                            End If
                        End If
                        
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT
                        ID_KANRI_TBL(ING_No).SYUKA_QTY = JITU_QTY
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>> 2017.09.22
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>> 2017.09.22
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
                        '>>>>>>>    2017.09.25
                        'Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        '>>>>>>>    2017.09.25
                                                                                
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Hinban)
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
'2007.07.21                        Send_Text.Box_Type(2).Max_Size = "13"
'2007.07.21                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(2).Max_Size = "20"                       '2007.07.21
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"  '2007.07.21
                                                                                
                                                                                
                                                                                
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
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
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
                        Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '���͌���
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
            
                '2007.07.21 ��
                If Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size) < 1 Then
                    wkTEXT = Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                Else
                    wkTEXT = Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                End If
                
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                
                Select Case wkTEXT
                '2007.07.21 ��
                
                                    
                    Case LCD_Hinban     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        ID_KANRI_TBL(ING_No).HINBAN_DAMMY = ""                                  '2017.10.30
                        If Trim(Hinban) = "." Then                                              '2017.10.24
                            Hinban = ID_KANRI_TBL(ING_No).Hinban                                '2017.10.24
                            ID_KANRI_TBL(ING_No).HINBAN_DAMMY = "."                             '2017.10.30
                        End If
                        
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc = False
                                Exit Function
                    
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                
                        End Select
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")      '2017.09.22
                            
                            '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                            '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                
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
                
                        
                        '2009.07.29
                        If ID_KANRI_TBL(ING_No).SYUKA_QTY > 1 Then
                        
                            Send_Text.buzzer = Wel_Inspe_BUZZER                 '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Wel_Inspe_BUZZER
                        
                        Else
                            Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        End If
                        
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>>>>   2017.09.22
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>>>>   2017.09.22
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
                        '>>>>>>>>>>>>>> 2017.09.25
                        'Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        '>>>>>>>>>>>>>> 2017.09.25
                                                                                
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Hinban)
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
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
                        '-----------------------------------------------�S�s��
                                                                                
                        '2009.04.16
                        wkNull_Check = Replace(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode), Chr(0), " ")
                        If Trim(wkNull_Check) = "" Then
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
                            Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                    '���͌���
                             Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""






                            '2010.12.09
                            If IsNumeric(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) Then
                                If Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) <> 0 Then
                                    Mod_Return = ID_KANRI_TBL(ING_No).SYUKA_QTY Mod Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))
                                    iNum = CInt(ToRoundDown(CCur(ID_KANRI_TBL(ING_No).SYUKA_QTY) / CCur(Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))), 0))
                            
                                    
                                    Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                                    ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                            
                            
                                    Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                            '�\�����e
                                    
                                    If Mod_Return <> 0 Then
                                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                    Else
                                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                    End If
                                    
                                                                                            '���l�����\��
                                    Send_Text.Box_Type(4).INIT = ""
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                            '�����J�[�\���ʒu
                                    Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                            '���͌���
                                     Send_Text.Box_Type(4).Max_Size = "00"
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                            
                                    Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                                End If
                            
                                                        
                            End If



                        Else
                                                                                    
                                                                                    
                            Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                                                                    
                                                                                    
                                                                                    'BOX����
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                                                                                    '���l�����\��
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                    '���͌���
                             Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        End If


                        Sendbuf = Text_Create_Proc()
                
                
                
                End Select
            
            Next i
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�iAny Key�j
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                            '�o�ח\��̓ǂݍ���
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_KANRI_TBL(ING_No).ID_NO)   'ID��
    
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc = False
                        GoTo Abort_Tran
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
        
            Loop
    
    
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                            
            '2006.07.20 ���i�S���ҏo�͒ǉ�
            Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
            Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, Format(Now, "HHMMSS"))
                                            
                                            '�o�ח\�菑����
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc = False
                        GoTo Abort_Tran
                
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                        Inspe_Proc = SYS_ERR
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
            Loop
                                        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                                        
            '2004.07.16 ��
            sts = IDOREKI_OUTPUT_PROC("", _
                                        "", _
                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                        ID_KANRI_TBL(ING_No).Hinban, _
                                        "", _
                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                        0, _
                                        0, _
                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        FILE_RETRY, _
                                        , _
                                        "", , , , , MENU_NO, _
                                        ID_KANRI_TBL(ING_No).MTS_CODE, _
                                        ID_KANRI_TBL(ING_No).SS_CODE, _
                                        ID_KANRI_TBL(ING_No).ID_NO, _
                                        , , , , , , , , , , , ID_KANRI_TBL(ING_No).HINBAN_DAMMY)
            Select Case sts
                Case False      '����I��
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Inspe_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
            '2004.07.16 ��
                                        
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
    
    
    End Select

    Inspe_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function









Public Function Inspe_Proc_MTS(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����i�l�s�r�ǂݍ��݂���j�x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 20          '2007.07.21 13-->20
Dim SYUKA_QTY       As Long
Dim JITU_QTY        As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim NAIGAI          As String * 1

'2010.12.07
'Dim HIN_NO          As String * 13
Dim HIN_NO          As String * 20
'2010.12.07


Dim KAN_FLG         As String * 1

Dim i               As Integer

Dim DEN_ID_LOOP     As Integer      '2006.06.01
Dim DEN_ID_JGYOBU   As String * 1   '2006.06.01


Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

Dim wkTEXT          As String


Dim wkKENPIN_YMD    As String * 8   '���i���t   2007.10.10


Dim wkNull_Check    As String       '2009.04.27

Dim Mod_Return      As Integer      '2010.12.09
Dim iNum            As Integer      '2010.12.09


    Inspe_Proc_MTS = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i������j
        
            For i = 0 To M_Gyo - 1
                
'>>>>>>>>>>>    2017.09.07
'                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
'                    Case LCD_MTS    '������
                Select Case i
                    Case 1
'>>>>>>>>>>>    2017.09.07
                        
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) < 16 Then
                                    '������i���Ӑ�j�݂̂Ō�����}�X�^�ǂݍ���
                            Call UniCode_Conv(K2_MTS.MUKE_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
                            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
                            Select Case sts
                                Case BtNoErr
                                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                                    
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "")     '2017.09.22
                                        '>>>>>>>    �G���[���b�Z�[�W�ύX    2017.09.25
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "") '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "������}�X�^�[���o�^", "", "")      '2017.09.22
                                        '>>>>>>>    �G���[���b�Z�[�W�ύX    2017.09.25
                    
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                        Inspe_Proc_MTS = False
                                        Exit Function
                                    
                                    End If
                                
                                Case BtErrKeyNotFound
                                
                                    Call UniCode_Conv(K3_MTS.SS_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
                                                        
                                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "")             '2017.09.21
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "������}�X�^�[���o�^", "", "")      '2017.09.21    2017.09.22 DEL
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "������}�X�^�[���o�^", "", "")      '2017.09.22
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                            Inspe_Proc_MTS = False
                                            Exit Function
                                        
                                        Case Else
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^", 0)
                                            Exit Function
                                    End Select
                        
                            End Select
                        
                            MTS_CODE = StrConv(MTSREC.MUKE_CODE, vbUnicode)
                            SS_CODE = StrConv(MTSREC.SS_CODE, vbUnicode)
                        
                        
                        Else
                            MTS_CODE = Left(ID_KANRI_TBL(ING_No).Recv_text(i), 8)
                            SS_CODE = Right(ID_KANRI_TBL(ING_No).Recv_text(i), 8)
                        
                                                '������}�X�^�ǂݍ���
                            Call UniCode_Conv(K0_MTS.MUKE_CODE, MTS_CODE)
                            Call UniCode_Conv(K0_MTS.SS_CODE, SS_CODE)
                         
                            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, MTS_CODE & SS_CODE, "�o�א�G���[", "", "")            '2017.09.21
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, MTS_CODE & SS_CODE, "������}�X�^�[���o�^", "", "")     '2017.09.21    2017.09.22 DEL
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, MTS_CODE & SS_CODE, "������}�X�^�[���o�^", "", "")     '2017.09.22
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Inspe_Proc_MTS = False
                                    Exit Function
                            
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^", 0)
                                    Exit Function
                            End Select
                        
                        
                        End If
                         
                         
                         
                         
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>    2017.09.22
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>    2017.09.22
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, MTS_CODE & SS_CODE)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, MTS_CODE & SS_CODE)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_ID_No)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_ID_No)
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
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
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
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
                        Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '���͌���
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�`�[�h�c�j
            For i = 0 To M_Gyo - 1
                
                
                '2007.07.21 ��
                If Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size) < 1 Then
                    wkTEXT = Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                Else
                    wkTEXT = Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                End If
                
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                
                Select Case wkTEXT
                '2007.07.21 ��
                
                
                    Case LCD_ID_No      '�`�[�h�c
    
    
    
    
    
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
'''                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
'''                        sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
'''                                                ID_KANRI_TBL(ING_No).JGYOBU, _
'''                                                NAIGAI, _
'''                                                Hinban, _
'''                                                MTS_CODE, _
'''                                                SS_CODE, _
'''                                                CYU_KBN, _
'''                                                Y_SYU_CNT, _
'''                                                ID_NO, _
'''                                                SYUKA_QTY, _
'''                                                DEN_NO, _
'''                                                KAN_KBN, _
'''                                                JITU_QTY)
'''
'''
                        
                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
                        For DEN_ID_LOOP = 0 To UBound(JGYOBU_T) '2006.06.01
                        
                            '2007.10.10 �����@���i���t�ǉ�
                            sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
                                                    JGYOBU_T(DEN_ID_LOOP).CODE, _
                                                    NAIGAI, _
                                                    Hinban, _
                                                    MTS_CODE, _
                                                    SS_CODE, _
                                                    CYU_KBN, _
                                                    Y_SYU_CNT, _
                                                    ID_NO, _
                                                    SYUKA_QTY, _
                                                    DEN_NO, _
                                                    KAN_KBN, _
                                                    JITU_QTY, _
                                                    DEN_ID_JGYOBU, _
                                                    wkKENPIN_YMD)
                            If sts <> False Or Y_SYU_CNT <> 0 Then
                                Exit For
                            End If
                        
                        Next DEN_ID_LOOP
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")     '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\�薳��", "", "") '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_MTS = False
                                    Exit Function
                                End If
                
                
                            Case True
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\��g�p��", "", "")   '2017.09.22
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc_MTS = False
                                Exit Function
                        End Select
                        
                        
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "���i��ƕs��", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "���i��ƕs��", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        
                        
                        '------------------ ������̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "������G���[", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�א�G���[", MTS_CODE, "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ �����敪�̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�����敪�~�X", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�����敪�G���[", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ �o�Ɋ����̃`�F�b�N
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                
                                
                                
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "��Ɩ�����", "", "")       '2017.09.07
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ɖ�����", "", "")        '2017.09.07 2017.09.22 DEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ɖ�����", "", "")        '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_MTS = False
                                Exit Function
                            End If
                        End If
                        
                        '------------------ �Č��i�̃`�F�b�N    2007.10.10
                        If Inspection_CHK = 1 Then
                            If Trim(wkKENPIN_YMD) <> "" Then
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "���i�ςł�", "", "")       '2017.09.22
                                '�G���[���b�Z�[�W�ύX   2017.09.25
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "���i�ςł�", "", "")   '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).MTS_CODE, ID_NO, Hinban, "�o�א�:" & SYUKA_QTY, "���i�ςł�")
                                '�G���[���b�Z�[�W�ύX   2017.09.25
                                
                                
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_MTS = False
                                Exit Function
                            End If
                        End If
                        
                        
                        
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT
                        ID_KANRI_TBL(ING_No).SYUKA_QTY = JITU_QTY
                        ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU = DEN_ID_JGYOBU      '2006.06.01
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>  2017.09.22
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>  2017.09.22
                                                                                
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
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
                        '-----------------------------------------------�Q�s��
                                                                                'BOX����
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>> 2017.09.25
                        'Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        '>>>>>> 2017.09.25
                                                                                
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
                        '-----------------------------------------------�R�s��
                                                                                'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
                                                                                '���l�����\��
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '���͌���
'2007.07.21                        Send_Text.Box_Type(3).Max_Size = "13"
'2007.07.21                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(3).Max_Size = "20"                       '2007.07.21
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"  '2007.07.21
                                                                                
                                                                                
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
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
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
            
                
                '2007.07.21 ��
                If Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size) < 1 Then
                    wkTEXT = Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                Else
                    wkTEXT = Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                End If
                
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                
                Select Case wkTEXT
                '2007.07.21 ��
                    
                    Case LCD_Hinban     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        ID_KANRI_TBL(ING_No).HINBAN_DAMMY = ""                                  '2017.10.30
                        If Trim(Hinban) = "." Then                                              '2017.10.24
                            Hinban = ID_KANRI_TBL(ING_No).Hinban                                '2017.10.24
                            ID_KANRI_TBL(ING_No).HINBAN_DAMMY = "."                             '2017.10.30
                        End If                                                                  '2017.10.24
                    
                        '2006.06.01 ID_KANRI_TBL(ING_No).JGYOBU--> DEN_ID_JGYOBU
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                
                                '�G���[���b�Z�[�W�ύX   2017.09.25
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                '�G���[���b�Z�[�W�ύX   2017.09.25
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_MTS = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                
                        End Select
                        
                        
                        If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                            '------------------------- 2011.07.29   �ǂݑւ����ނōēǂݍ���
                            
                            
''                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")
''                            Sendbuf = Text_Create_Proc()
''                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
''                            Inspe_Proc_MTS = False
''                            Exit Function
                            
                            Call UniCode_Conv(K5_ITEM.JGYOBU, RET_JGYOBU)
                            Call UniCode_Conv(K5_ITEM.NAIGAI, RET_NAIGAI)
                            Call UniCode_Conv(K5_ITEM.HIN_CHANGE, Hinban)
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K5_ITEM, Len(K5_ITEM), 5)
                            Select Case sts
                                Case BtNoErr
                                
                                    If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                                
                                
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")      '2017.09.22
                                        
                                        
                                        '�G���[���b�Z�[�W�ύX   2017.09.25
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                        '�G���[���b�Z�[�W�ύX   2017.09.25
                                        
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_MTS = False
                                        Exit Function
                                        
                                    End If
                                    Hinban = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                Case BtErrKeyNotFound
                                
                                
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")          '2017.09.22
                                    
                                    '�G���[���b�Z�[�W�ύX   2017.09.25
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                    '�G���[���b�Z�[�W�ύX   2017.09.25
                                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_MTS = False
                                    Exit Function
                                
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                            '------------------------- 2011.07.29
                        End If
                
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
                
                        '2009.07.29
                        If ID_KANRI_TBL(ING_No).SYUKA_QTY > 1 Then
                        
                            Send_Text.buzzer = Wel_Inspe_BUZZER                 '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Wel_Inspe_BUZZER
                        
                        Else
                            Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        End If
                        
                        
'>>>>>>>>>> 2017.09.25 ������|�|���^�C�g���@�\���ύX
                        '-----------------------------------------------�P�s��
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
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
'                                                                                'BOX����
'                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                                                                                '�\�����e
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
'                                                                                '���l�����\��
'                        Send_Text.Box_Type(0).INIT = ""
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
'                                                                                '�����J�[�\���ʒu
'                        Send_Text.Box_Type(0).Start_Pos = ""
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
'                                                                                '���͌���
'                        Send_Text.Box_Type(0).Max_Size = "00"
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
'
'                        Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
'>>>>>>>>>> 2017.09.25 ������|�|���^�C�g���@�\���ύX
                        '-----------------------------------------------�R�s��
                                                                                'BOX����
                        Send_Text.Box_Type(1).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>  2017.09.25
'                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        '>>>>>>>>>  2017.09.25
                                                                                
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
                        '-----------------------------------------------�S�s��
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
                        Send_Text.Box_Type(2).Start_Pos = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�T�s��
                                                                                'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
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
                        '-----------------------------------------------�S�s��


                        
                        
                        '2009.04.16
                        wkNull_Check = Replace(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode), Chr(0), " ")
                        If Trim(wkNull_Check) = "" Then
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
                            Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                    '���͌���
                             Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                            
                                                        
                            '2010.12.09
                            If IsNumeric(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) Then
                                If Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) <> 0 Then
                                    Mod_Return = ID_KANRI_TBL(ING_No).SYUKA_QTY Mod Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))
                                    iNum = CInt(ToRoundDown(CCur(ID_KANRI_TBL(ING_No).SYUKA_QTY) / CCur(Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))), 0))
                            
                                    
                                    Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                                    ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                            
                            
                                    Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                            '�\�����e
                                    
                                    If Mod_Return <> 0 Then
                                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                    Else
                                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                    End If
                                    
                                                                                            '���l�����\��
                                    Send_Text.Box_Type(4).INIT = ""
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                            '�����J�[�\���ʒu
                                    Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                            '���͌���
                                     Send_Text.Box_Type(4).Max_Size = "00"
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                            
                                    Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                                End If
                            
                                                        
                            End If
                        Else
                                                                                    
                                                                                    
                            Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                                                                    
                                                                                    
                                                                                    'BOX����
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                                                                                    '���l�����\��
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                    '���͌���
                             Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""



                            



                        End If



                        Sendbuf = Text_Create_Proc()
                
                
                
                End Select
            
            Next i
        Case Step_Sagyo4_RES        '�S��ڂ̎�M�iAny Key�j
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                            '�o�ח\��̓ǂݍ���
            '2006.06.01 ID_KANRI_TBL(ING_No).JGYOBU--> DEN_ID_JGYOBU
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU)     '���ƕ�
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_KANRI_TBL(ING_No).ID_NO)   'ID��
    
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc_MTS = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")          '2017.09.22
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc_MTS = False
                        GoTo Abort_Tran
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
        
            Loop
    
    
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                            
            '2006.07.20 ���i�S���ҏo�͒ǉ�
            Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
            Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, Format(Now, "HHMMSS"))
                                            
                                            
                                            '�o�ח\�菑����
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc_MTS = False
                        GoTo Abort_Tran
                
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                        Inspe_Proc_MTS = SYS_ERR
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        GoTo Abort_Tran
                End Select
            Loop
                                        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                                        
                                        
                                        
            '2004.07.16 ��
            
            '2006.06.01 ID_KANRI_TBL(ING_No).JGYOBU--> DEN_ID_JGYOBU
            
            sts = IDOREKI_OUTPUT_PROC("", _
                                        "", _
                                        ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU, _
                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                        ID_KANRI_TBL(ING_No).Hinban, _
                                        "", _
                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                        0, _
                                        0, _
                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        FILE_RETRY, _
                                        , _
                                        "", , , , , MENU_NO, _
                                        ID_KANRI_TBL(ING_No).MTS_CODE, _
                                        ID_KANRI_TBL(ING_No).SS_CODE, _
                                        ID_KANRI_TBL(ING_No).ID_NO, , , , , , , , , , , , ID_KANRI_TBL(ING_No).HINBAN_DAMMY)
            Select Case sts
                Case False      '����I��
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Inspe_Proc_MTS = SYS_ERR
                    GoTo Abort_Tran
            End Select
            '2004.07.16 ��
                                        
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
                
                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                
                '���i�m�F�I�����́A�`�[�h�c�v���ɖ߂��ׂ̓��ꏈ��
                
                
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
        
                Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                
                
                
                
                '-----------------------------------------------�P�s��
                                                                'BOX����
                Send_Text.Box_Type(0).Box_Type = TYPE_REF
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                '�\�����e
                '>>>>>>>>>> 2017.09.22
'                Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                        
                Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                '>>>>>>>>>> 2017.09.22
                                                                        
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
                Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
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
                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_ID_No)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_ID_No)
                                                                        '���l�����\��
                Send_Text.Box_Type(2).INIT = ""
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                        '�����J�[�\���ʒu
                Send_Text.Box_Type(2).Start_Pos = "01"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                        '���͌���
                Send_Text.Box_Type(2).Max_Size = "12"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "12"
                                                                        
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
                                                                        
                Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                '-----------------------------------------------�S�s��
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
                Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                        '���͌���
                 Send_Text.Box_Type(4).Max_Size = "00"
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                        
                Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                Sendbuf = Text_Create_Proc()


'                If Sagyo_Send_Proc() Then
'                    Sendbuf = Text_Create_Proc()
'                    Exit Function
'                End If
            
                Sendbuf = Text_Create_Proc()
    
    
    End Select

    Inspe_Proc_MTS = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function






Public Function New_Inspe_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 20      '2007.07.21 13-->20
Dim SYUKA_QTY       As Long
Dim JITU_QTY        As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim NAIGAI          As String * 1
Dim HIN_NO          As String * 13


Dim KAN_FLG         As String * 1

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim wkTEXT          As String           '2007.07.21

Dim MENU_NO         As String * 2

Dim wkNull_Check    As String       '2009.04.27

Dim Mod_Return      As Integer      '2010.12.09
Dim iNum            As Integer      '2010.12.09

Dim wkKENPIN_YMD    As String * 8   '���i���t   2017.09.07


    New_Inspe_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�`�[�h�c�j
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No      '�`�[�h�c
    
    
'>>>>   2017.09.07
'                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
'
'                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���i��ƕs��", "", "")
'                            Sendbuf = Text_Create_Proc()
'                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                            Inspe_Proc = False
'                            Exit Function
'                        End If
'>>>> 2017.09.07
    
    
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                NAIGAI, _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN, _
                                                JITU_QTY, , wkKENPIN_YMD)
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")         '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\�薳��", "", "")     '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    New_Inspe_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")           '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                New_Inspe_Proc = False
                                Exit Function
                        End Select
                
'>>>>> 2017.09.07
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���i��ƕs��", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���i��ƕs��", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            New_Inspe_Proc = False
                            Exit Function
                        End If
'>>>>> 2017.09.07
                
                
                
                        '------------------ ������̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "������G���[", "", "")     '2017.09.22
                            '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "������G���[", "", "") '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�א�G���[", MTS_CODE, "") '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            New_Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ �����敪�̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            
                            
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�����敪�~�X", "", "")     '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�����敪�G���[", "", "") '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            New_Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ �o�Ɋ����̃`�F�b�N
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                
                                
                                
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "��Ɩ�����", "", "")   '2017.09.07
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ɖ�����", "", "")    '2017.09.07        2017.09.22 DEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ɖ�����", "", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                New_Inspe_Proc = False
                                Exit Function
                            End If
                        End If
                        
                        
                        
                        '------------------ �Č��i�̃`�F�b�N    2017.09.07
                        If Inspection_CHK = 1 Then
                            If Trim(wkKENPIN_YMD) <> "" Then
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "���i�ςł�", "", "")       '2017.09.22
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "���i�ςł�", "", "")
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).MTS_CODE, ID_NO, Hinban, "�o�א�:" & SYUKA_QTY, "���i�ςł�")
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                New_Inspe_Proc = False
                                Exit Function
                            End If
                        End If
                        
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT
                        ID_KANRI_TBL(ING_No).SYUKA_QTY = JITU_QTY
                        
                        ID_KANRI_TBL(ING_No).ITEM_READ_CNT = 0
                        
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>> 2017.09.22
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>> 2017.09.22
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
                        '>>>>>>>    2017.09.25
                        'Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        '>>>>>>>    2017.09.25
                                                                                
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Hinban)
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
'2007.07.21                        Send_Text.Box_Type(2).Max_Size = "13"
'2007.07.21                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(2).Max_Size = "20"                       '2007.07.21
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "20"  '2007.07.21
                                                                                
                                                                                
                                                                                
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
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
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
                        Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '���͌���
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
            
                '2007.07.21 ��
                If Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size) < 1 Then
                    wkTEXT = Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                Else
                    wkTEXT = Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                End If
                
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                
                Select Case i
                '2007.07.21 ��
                
                                    
                    Case 2     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        
                        If IsNumeric(Hinban) And Val(Hinban) = ID_KANRI_TBL(ING_No).SYUKA_QTY Then
                        
                        
                            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                            '�g�����U�N�V�����J�n
                            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            If sts <> BtNoErr Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                Exit Function
                            End If
                                                            '�o�ח\��̓ǂݍ���
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_KANRI_TBL(ING_No).ID_NO)   'ID��
                    
                            Do
                            
                                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrKeyNotFound
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        New_Inspe_Proc = False
                                        GoTo Abort_Tran
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        New_Inspe_Proc = False
                                        GoTo Abort_Tran
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                        
                            Loop
                    
                    
                            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                                            
                            '2006.07.20 ���i�S���ҏo�͒ǉ�
                            Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
                            Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, Format(Now, "HHMMSS"))
                                                            
                                                            '�o�ח\�菑����
                            Do
                                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        New_Inspe_Proc = False
                                        GoTo Abort_Tran
                                
                                    Case Else
                                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                        New_Inspe_Proc = SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                            Loop
                                                        
                            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                        
                            Else
                                        
                                MENU_NO = ""
                            End If
                                                        
                            '2004.07.16 ��
                            sts = IDOREKI_OUTPUT_PROC("", _
                                                        "", _
                                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                                        ID_KANRI_TBL(ING_No).Hinban, _
                                                        "", _
                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                        0, _
                                                        0, _
                                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                        FILE_RETRY, _
                                                        , _
                                                        "", , , , , MENU_NO, _
                                                        ID_KANRI_TBL(ING_No).MTS_CODE, _
                                                        ID_KANRI_TBL(ING_No).SS_CODE, _
                                                        ID_KANRI_TBL(ING_No).ID_NO, _
                                                        , , , , , , , , , , , ID_KANRI_TBL(ING_No).HINBAN_DAMMY)
                            Select Case sts
                                Case False      '����I��
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    New_Inspe_Proc = SYS_ERR
                                    GoTo Abort_Tran
                            End Select
                            '2004.07.16 ��
                                                        
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
                        
                        
                        
                        Else
                        
                        
                            ID_KANRI_TBL(ING_No).HINBAN_DAMMY = ""                                  '2017.10.30
                            If Trim(Hinban) = "." Then                                              '2017.10.24
                                Hinban = ID_KANRI_TBL(ING_No).Hinban                                '2017.10.24
                                ID_KANRI_TBL(ING_No).HINBAN_DAMMY = "."                             '2017.10.30
                            End If
                            
                            
                            sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                            Select Case sts
                                Case BtNoErr
                        
                                Case BtErrKeyNotFound
                                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                    
                                    '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                    '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                        
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    New_Inspe_Proc = False
                                    Exit Function
                        
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                    Exit Function
                    
                            End Select
                            
                            If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")      '2017.09.22
                                
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                '>>>>>>>>>>>>>>>    �G���[���b�Z�[�W�ύX�@2017.09.25
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                New_Inspe_Proc = False
                                Exit Function
                            End If
                    
                            ID_KANRI_TBL(ING_No).ITEM_READ_CNT = ID_KANRI_TBL(ING_No).ITEM_READ_CNT + 1
                    
                    
                    
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
                    
                            
                            '2009.07.29
                            If ID_KANRI_TBL(ING_No).SYUKA_QTY > 1 Then
                            
                                Send_Text.buzzer = Wel_Inspe_BUZZER                 '�u�U�[���@�W��
                                ID_KANRI_TBL(ING_No).Send_Text.buzzer = Wel_Inspe_BUZZER
                            
                            Else
                                Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                                ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                            End If
                            
                            
                            '-----------------------------------------------�P�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            '>>>>>>>>>>>>   2017.09.22
    '                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
    '                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                            
                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                            '>>>>>>>>>>>>   2017.09.22
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
                            '>>>>>>>>>>>>>> 2017.09.25
                            'Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                            'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                            '>>>>>>>>>>>>>> 2017.09.25
                                                                                    
                                                                                    
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
                                                                                    
                            Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                            '-----------------------------------------------�S�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(3).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide) & "�^" & StrConv(Format(ID_KANRI_TBL(ING_No).ITEM_READ_CNT, "#0"), vbWide))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide) & "�^" & StrConv(Format(ID_KANRI_TBL(ING_No).ITEM_READ_CNT, "#0"), vbWide))
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
                            '-----------------------------------------------�S�s��
                                                                                    
                            '2009.04.16
                            wkNull_Check = Replace(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode), Chr(0), " ")
                            If Trim(wkNull_Check) = "" Then
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
                                Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                        '���͌���
                                 Send_Text.Box_Type(4).Max_Size = "00"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                        
                                Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
    
    
    
    
    
                                '2010.12.09
                                If IsNumeric(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) Then
                                    If Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) <> 0 Then
                                        Mod_Return = ID_KANRI_TBL(ING_No).SYUKA_QTY Mod Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))
                                        iNum = CInt(ToRoundDown(CCur(ID_KANRI_TBL(ING_No).SYUKA_QTY) / CCur(Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))), 0))
                                
                                        
                                        Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                
                                
                                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                                '�\�����e
                                        
                                        If Mod_Return <> 0 Then
                                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                        Else
                                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                        End If
                                        
                                                                                                '���l�����\��
                                        Send_Text.Box_Type(4).INIT = ""
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                                '�����J�[�\���ʒu
                                        Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                                '���͌���
                                         Send_Text.Box_Type(4).Max_Size = "00"
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                                
                                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                                    End If
                                
                                                            
                                End If
    
    
    
                            Else
                                                                                        
                                                                                        
                                Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                                ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                                                                        
                                                                                        
                                                                                        'BOX����
                                Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                                                                                        '���l�����\��
                                Send_Text.Box_Type(4).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                        '�����J�[�\���ʒu
                                Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                        '���͌���
                                 Send_Text.Box_Type(4).Max_Size = "00"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                        
                                Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            End If
                        
                        
                            Sendbuf = Text_Create_Proc()
                        
                        
                        End If

                
                
                
                End Select
            
            Next i
    
    End Select

    New_Inspe_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If




End Function









Public Function New_Inspe_Proc_MTS(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����i�l�s�r�ǂݍ��݂���j�x
'
'   2018.11.05�@�i�ԕ����Ǎ���
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 20          '2007.07.21 13-->20
Dim SYUKA_QTY       As Long
Dim JITU_QTY        As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim NAIGAI          As String * 1

'2010.12.07
'Dim HIN_NO          As String * 13
Dim HIN_NO          As String * 20
'2010.12.07


Dim KAN_FLG         As String * 1

Dim i               As Integer
Dim j               As Integer

Dim DEN_ID_LOOP     As Integer      '2006.06.01
Dim DEN_ID_JGYOBU   As String * 1   '2006.06.01


Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

Dim wkTEXT          As String


Dim wkKENPIN_YMD    As String * 8   '���i���t   2007.10.10


Dim wkNull_Check    As String       '2009.04.27

Dim Mod_Return      As Integer      '2010.12.09
Dim iNum            As Integer      '2010.12.09


    New_Inspe_Proc_MTS = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i������j
        
            For i = 0 To M_Gyo - 1
                
'>>>>>>>>>>>    2017.09.07
'                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
'                    Case LCD_MTS    '������
                Select Case i
                    Case 1
'>>>>>>>>>>>    2017.09.07
                        
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) < 16 Then
                                    '������i���Ӑ�j�݂̂Ō�����}�X�^�ǂݍ���
                            Call UniCode_Conv(K2_MTS.MUKE_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
                            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
                            Select Case sts
                                Case BtNoErr
                                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                                    
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "")     '2017.09.22
                                        '>>>>>>>    �G���[���b�Z�[�W�ύX    2017.09.25
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "") '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "������}�X�^�[���o�^", "", "")      '2017.09.22
                                        '>>>>>>>    �G���[���b�Z�[�W�ύX    2017.09.25
                    
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                        New_Inspe_Proc_MTS = False
                                        Exit Function
                                    
                                    End If
                                
                                Case BtErrKeyNotFound
                                
                                    Call UniCode_Conv(K3_MTS.SS_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
                                                        
                                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                        Case BtErrKeyNotFound
                                        
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "")             '2017.09.21
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "������}�X�^�[���o�^", "", "")      '2017.09.21    2017.09.22 DEL
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "������}�X�^�[���o�^", "", "")      '2017.09.22
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                            New_Inspe_Proc_MTS = False
                                            Exit Function
                                        
                                        Case Else
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^", 0)
                                            Exit Function
                                    End Select
                        
                            End Select
                        
                            MTS_CODE = StrConv(MTSREC.MUKE_CODE, vbUnicode)
                            SS_CODE = StrConv(MTSREC.SS_CODE, vbUnicode)
                        
                        
                        Else
                            MTS_CODE = Left(ID_KANRI_TBL(ING_No).Recv_text(i), 8)
                            SS_CODE = Right(ID_KANRI_TBL(ING_No).Recv_text(i), 8)
                        
                                                '������}�X�^�ǂݍ���
                            Call UniCode_Conv(K0_MTS.MUKE_CODE, MTS_CODE)
                            Call UniCode_Conv(K0_MTS.SS_CODE, SS_CODE)
                         
                            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, MTS_CODE & SS_CODE, "�o�א�G���[", "", "")            '2017.09.21
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, MTS_CODE & SS_CODE, "������}�X�^�[���o�^", "", "")     '2017.09.21    2017.09.22 DEL
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, MTS_CODE & SS_CODE, "������}�X�^�[���o�^", "", "")     '2017.09.22
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    New_Inspe_Proc_MTS = False
                                    Exit Function
                            
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^", 0)
                                    Exit Function
                            End Select
                        
                        
                        End If
                         
                         
                         
                         
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>    2017.09.22
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>    2017.09.22
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, MTS_CODE & SS_CODE)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, MTS_CODE & SS_CODE)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_ID_No)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_ID_No)
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
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
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
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
                        Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                '���͌���
                         Send_Text.Box_Type(4).Max_Size = "00"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�`�[�h�c�j
            For i = 0 To M_Gyo - 1
                
                
                '2007.07.21 ��
                If Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size) < 1 Then
                    wkTEXT = Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                Else
                    wkTEXT = Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                End If
                
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                
                Select Case wkTEXT
                '2007.07.21 ��
                
                
                    Case LCD_ID_No      '�`�[�h�c
    
    
    
    
    
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
'''                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
'''                        sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
'''                                                ID_KANRI_TBL(ING_No).JGYOBU, _
'''                                                NAIGAI, _
'''                                                Hinban, _
'''                                                MTS_CODE, _
'''                                                SS_CODE, _
'''                                                CYU_KBN, _
'''                                                Y_SYU_CNT, _
'''                                                ID_NO, _
'''                                                SYUKA_QTY, _
'''                                                DEN_NO, _
'''                                                KAN_KBN, _
'''                                                JITU_QTY)
'''
'''
                        
                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
                        For DEN_ID_LOOP = 0 To UBound(JGYOBU_T) '2006.06.01
                        
                            '2007.10.10 �����@���i���t�ǉ�
                            sts = Y_Syuka_Chek_Proc(KAN_KBN_FIN, _
                                                    JGYOBU_T(DEN_ID_LOOP).CODE, _
                                                    NAIGAI, _
                                                    Hinban, _
                                                    MTS_CODE, _
                                                    SS_CODE, _
                                                    CYU_KBN, _
                                                    Y_SYU_CNT, _
                                                    ID_NO, _
                                                    SYUKA_QTY, _
                                                    DEN_NO, _
                                                    KAN_KBN, _
                                                    JITU_QTY, _
                                                    DEN_ID_JGYOBU, _
                                                    wkKENPIN_YMD)
                            If sts <> False Or Y_SYU_CNT <> 0 Then
                                Exit For
                            End If
                        
                        Next DEN_ID_LOOP
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")     '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\�薳��", "", "") '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    New_Inspe_Proc_MTS = False
                                    Exit Function
                                End If
                
                
                            Case True
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\��g�p��", "", "")   '2017.09.22
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                New_Inspe_Proc_MTS = False
                                Exit Function
                        End Select
                        
                        
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "���i��ƕs��", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "���i��ƕs��", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            New_Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        
                        
                        '------------------ ������̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "������G���[", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�א�G���[", MTS_CODE, "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            New_Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ �����敪�̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                
                                Case SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                            End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                            
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�����敪�~�X", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�����敪�G���[", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            New_Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ �o�Ɋ����̃`�F�b�N
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                
                                
                                
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "��Ɩ�����", "", "")       '2017.09.07
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ɖ�����", "", "")        '2017.09.07 2017.09.22 DEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ɖ�����", "", "")        '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                New_Inspe_Proc_MTS = False
                                Exit Function
                            End If
                        End If
                        
                        '------------------ �Č��i�̃`�F�b�N    2007.10.10
                        If Inspection_CHK = 1 Then
                            If Trim(wkKENPIN_YMD) <> "" Then
                                
                                
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                sts = Data_Clear_Proc(0, Sendbuf)
                                Select Case sts
                                    Case SYS_CANCEL
                                        Call Err_Send_Proc("�f�[�^�g�p��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    
                                    Case SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                End Select
'---------------------------    �f�[�^�����ǉ�    2011.07.29
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "���i�ςł�", "", "")       '2017.09.22
                                '�G���[���b�Z�[�W�ύX   2017.09.25
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "���i�ςł�", "", "")   '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).MTS_CODE, ID_NO, Hinban, "�o�א�:" & SYUKA_QTY, "���i�ςł�")
                                '�G���[���b�Z�[�W�ύX   2017.09.25
                                
                                
                                
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                New_Inspe_Proc_MTS = False
                                Exit Function
                            End If
                        End If
                        
                        
                        
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT
                        ID_KANRI_TBL(ING_No).SYUKA_QTY = JITU_QTY
                        ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU = DEN_ID_JGYOBU      '2006.06.01
                        
                        
                        ID_KANRI_TBL(ING_No).ITEM_READ_CNT = 0
                        
                        
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
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>  2017.09.22
'                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>  2017.09.22
                                                                                
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
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
                        '-----------------------------------------------�Q�s��
                                                                                'BOX����
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>> 2017.09.25
                        'Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                        '>>>>>> 2017.09.25
                                                                                
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
                        '-----------------------------------------------�R�s��
                                                                                'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
                                                                                '���l�����\��
                        Send_Text.Box_Type(3).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(3).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(3).Max_Size = "20"                       '2007.07.21
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"  '2007.07.21
                                                                                
                                                                                
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
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
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
            
                
                '2007.07.21 ��
'                If Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size) < 1 Then
'                    wkTEXT = Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
'                Else
'                    wkTEXT = Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
'                End If
                
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                
                If ID_KANRI_TBL(ING_No).ITEM_READ_CNT >= 1 Then
                    j = 2
                Else
                    j = 3
                End If
                Select Case i
                '2007.07.21 ��
                    
                    Case j          '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        If IsNumeric(Hinban) And Val(Hinban) = ID_KANRI_TBL(ING_No).SYUKA_QTY Then
                        
                        
                                                    
                    
                            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                        '�g�����U�N�V�����J�n
                            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            If sts <> BtNoErr Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                                Exit Function
                            End If
                                                            '�o�ח\��̓ǂݍ���
                            '2006.06.01 ID_KANRI_TBL(ING_No).JGYOBU--> DEN_ID_JGYOBU
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU)     '���ƕ�
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_KANRI_TBL(ING_No).ID_NO)   'ID��
                    
                            Do
                            
                                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrKeyNotFound
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        New_Inspe_Proc_MTS = False
                                        GoTo Abort_Tran
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")          '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        New_Inspe_Proc_MTS = False
                                        GoTo Abort_Tran
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                        
                            Loop
                    
                    
                            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                            Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                                            
                            '2006.07.20 ���i�S���ҏo�͒ǉ�
                            Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
                            Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, Format(Now, "HHMMSS"))
                                                            
                                                            
                                                            '�o�ח\�菑����
                            Do
                                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        New_Inspe_Proc_MTS = False
                                        GoTo Abort_Tran
                                
                                    Case Else
                                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                        New_Inspe_Proc_MTS = SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                            Loop
                                                        
                            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                        
                            Else
                                        
                                MENU_NO = ""
                            End If
                                                        
                                                        
                                                        
                            '2004.07.16 ��
                            
                            '2006.06.01 ID_KANRI_TBL(ING_No).JGYOBU--> DEN_ID_JGYOBU
                            
                            sts = IDOREKI_OUTPUT_PROC("", _
                                                        "", _
                                                        ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU, _
                                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                                        ID_KANRI_TBL(ING_No).Hinban, _
                                                        "", _
                                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                        0, _
                                                        0, _
                                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                        FILE_RETRY, _
                                                        , _
                                                        "", , , , , MENU_NO, _
                                                        ID_KANRI_TBL(ING_No).MTS_CODE, _
                                                        ID_KANRI_TBL(ING_No).SS_CODE, _
                                                        ID_KANRI_TBL(ING_No).ID_NO, , , , , , , , , , , , ID_KANRI_TBL(ING_No).HINBAN_DAMMY)
                            Select Case sts
                                Case False      '����I��
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    New_Inspe_Proc_MTS = SYS_ERR
                                    GoTo Abort_Tran
                            End Select
                            '2004.07.16 ��
                                                        
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
                                
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                                
                                '���i�m�F�I�����́A�`�[�h�c�v���ɖ߂��ׂ̓��ꏈ��
                                
                                
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
                        
                                Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                                ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                                
                                
                                
                                
                                '-----------------------------------------------�P�s��
                                                                                'BOX����
                                Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                                '>>>>>>>>>> 2017.09.22
                '                Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                '                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                                '>>>>>>>>>> 2017.09.22
                                                                                        
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
                                Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
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
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_ID_No)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_ID_No)
                                                                                        '���l�����\��
                                Send_Text.Box_Type(2).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                        '�����J�[�\���ʒu
                                Send_Text.Box_Type(2).Start_Pos = "01"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                        '���͌���
                                Send_Text.Box_Type(2).Max_Size = "12"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "12"
                                                                                        
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
                                                                                        
                                Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                '-----------------------------------------------�S�s��
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
                                Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                        '���͌���
                                 Send_Text.Box_Type(4).Max_Size = "00"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                        
                                Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                
                                Sendbuf = Text_Create_Proc()
                
                
                '                If Sagyo_Send_Proc() Then
                '                    Sendbuf = Text_Create_Proc()
                '                    Exit Function
                '                End If
                            
                                Sendbuf = Text_Create_Proc()
                        
                                                    
                            Else
                        
                        
                                ID_KANRI_TBL(ING_No).HINBAN_DAMMY = ""                                  '2017.10.30
                                If Trim(Hinban) = "." Then                                              '2017.10.24
                                    Hinban = ID_KANRI_TBL(ING_No).Hinban                                '2017.10.24
                                    ID_KANRI_TBL(ING_No).HINBAN_DAMMY = "."                             '2017.10.30
                                End If                                                                  '2017.10.24
                            
                                '2006.06.01 ID_KANRI_TBL(ING_No).JGYOBU--> DEN_ID_JGYOBU
                                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).DEN_ID_JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                                Select Case sts
                                    Case BtNoErr
                            
                                    Case BtErrKeyNotFound
                                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                        
                                        '�G���[���b�Z�[�W�ύX   2017.09.25
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                        '�G���[���b�Z�[�W�ύX   2017.09.25
                            
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        New_Inspe_Proc_MTS = False
                                        Exit Function
                                
                                    Case Else
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                        Exit Function
                        
                                End Select
                                
                                
                                If Trim(Hinban) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                                    '------------------------- 2011.07.29   �ǂݑւ����ނōēǂݍ���
                                    
                                    
        ''                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")
        ''                            Sendbuf = Text_Create_Proc()
        ''                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
        ''                            Inspe_Proc_MTS = False
        ''                            Exit Function
                                    
                                    Call UniCode_Conv(K5_ITEM.JGYOBU, RET_JGYOBU)
                                    Call UniCode_Conv(K5_ITEM.NAIGAI, RET_NAIGAI)
                                    Call UniCode_Conv(K5_ITEM.HIN_CHANGE, Hinban)
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K5_ITEM, Len(K5_ITEM), 5)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            If Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) <> Trim(ID_KANRI_TBL(ING_No).Hinban) Then
                                        
                                        
                                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")      '2017.09.22
                                                
                                                
                                                '�G���[���b�Z�[�W�ύX   2017.09.25
                                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                                '�G���[���b�Z�[�W�ύX   2017.09.25
                                                
                                                Sendbuf = Text_Create_Proc()
                                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                                New_Inspe_Proc_MTS = False
                                                Exit Function
                                                
                                            End If
                                            Hinban = Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                        Case BtErrKeyNotFound
                                        
                                        
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")          '2017.09.22
                                            
                                            '�G���[���b�Z�[�W�ύX   2017.09.25
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", Hinban, "")  '2017.09.22
                                            '�G���[���b�Z�[�W�ύX   2017.09.25
                                            
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            New_Inspe_Proc_MTS = False
                                            Exit Function
                                        
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                    '------------------------- 2011.07.29
                                End If
                        
                        
                                ID_KANRI_TBL(ING_No).ITEM_READ_CNT = ID_KANRI_TBL(ING_No).ITEM_READ_CNT + 1
                        
                        
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
                        
                                '2009.07.29
                                If ID_KANRI_TBL(ING_No).SYUKA_QTY > 1 Then
                                
                                    Send_Text.buzzer = Wel_Inspe_BUZZER                 '�u�U�[���@�W��
                                    ID_KANRI_TBL(ING_No).Send_Text.buzzer = Wel_Inspe_BUZZER
                                
                                Else
                                    Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                                    ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                                End If
                                
                                
        '>>>>>>>>>> 2017.09.25 ������|�|���^�C�g���@�\���ύX
                                '-----------------------------------------------�P�s��
                                Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
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
        '                                                                                'BOX����
        '                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
        '                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
        '                                                                                '�\�����e
        '                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
        '                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
        '                                                                                '���l�����\��
        '                        Send_Text.Box_Type(0).INIT = ""
        '                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
        '                                                                                '�����J�[�\���ʒu
        '                        Send_Text.Box_Type(0).Start_Pos = ""
        '                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
        '                                                                                '���͌���
        '                        Send_Text.Box_Type(0).Max_Size = "00"
        '                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
        '
        '                        Send_Text.Box_Type(0).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
        '>>>>>>>>>> 2017.09.25 ������|�|���^�C�g���@�\���ύX
                                '-----------------------------------------------�R�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                '>>>>>>>>>  2017.09.25
        '                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
        '                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).ID_NO)
                                '>>>>>>>>>  2017.09.25
                                                                                        
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
                                '-----------------------------------------------�S�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
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
                                                                                        
                                Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                                '-----------------------------------------------�T�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide) & "�^" & StrConv(Format(ID_KANRI_TBL(ING_No).ITEM_READ_CNT, "#0"), vbWide))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide) & "�^" & StrConv(Format(ID_KANRI_TBL(ING_No).ITEM_READ_CNT, "#0"), vbWide))
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
                                '-----------------------------------------------�S�s��
        
        
                                
                                
                                '2009.04.16
                                wkNull_Check = Replace(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode), Chr(0), " ")
                                If Trim(wkNull_Check) = "" Then
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
                                    Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                            '���͌���
                                     Send_Text.Box_Type(4).Max_Size = "00"
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                            
                                    Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                                    
                                                                
                                    '2010.12.09
                                    If IsNumeric(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) Then
                                        If Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode)) <> 0 Then
                                            Mod_Return = ID_KANRI_TBL(ING_No).SYUKA_QTY Mod Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))
                                            iNum = CInt(ToRoundDown(CCur(ID_KANRI_TBL(ING_No).SYUKA_QTY) / CCur(Val(StrConv(ITEMREC.GAISO_IRI_QTY, vbUnicode))), 0))
                                    
                                            
                                            Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                    
                                    
                                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                                    '�\�����e
                                            
                                            If Mod_Return <> 0 Then
                                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "���{" & StrConv(Format(Mod_Return, "#0"), vbWide) & "�]")
                                            Else
                                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�@�O���F" & StrConv(Format(iNum, "#0"), vbWide) & "��")
                                            End If
                                            
                                                                                                    '���l�����\��
                                            Send_Text.Box_Type(4).INIT = ""
                                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                                    '�����J�[�\���ʒu
                                            Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                                    '���͌���
                                             Send_Text.Box_Type(4).Max_Size = "00"
                                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                                    
                                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                                        End If
                                    
                                                                
                                    End If
                                Else
                                                                                            
                                                                                            
                                    Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                                    ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                                                                            
                                                                                            
                                                                                            'BOX����
                                    Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                            '�\�����e
                                    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Trim(StrConv(ITEMREC.INSP_MESSAGE, vbUnicode)))
                                                                                            '���l�����\��
                                    Send_Text.Box_Type(4).INIT = ""
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                            '�����J�[�\���ʒu
                                    Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                                            '���͌���
                                     Send_Text.Box_Type(4).Max_Size = "00"
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                                            
                                    Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
        
                                    
        
        
        
                                End If
        
        
        
                                Sendbuf = Text_Create_Proc()
                
                        End If
                
                End Select
            
            Next i
    
    
    End Select

    New_Inspe_Proc_MTS = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function



