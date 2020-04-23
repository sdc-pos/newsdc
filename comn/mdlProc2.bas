Attribute VB_Name = "mdlProc2"
Option Explicit

'[2014/02/10 - M.MATSUYAMA �ړ�(Ver2.0.0)] F1100101����ړ�

Public Function Tanto_Check_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   �w�S���҃R�[�h�̃`�F�b�N�x
'
'-------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Tanto_Check_Proc = True

    For i = 0 To M_Gyo
        
        If ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = TYPE_REF Then
        Else
                                '�S���҃}�X�^�ǂݍ���
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    ID_KANRI_TBL(ING_No).TANTO_CODE = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                    Exit For
                Case BtErrKeyNotFound
                    
                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                    Call Err_Send_Proc("�S���Җ��o�^", "", "", "", "")
                    
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                    
                    Tanto_Check_Proc = False
                    Exit Function
                Case Else
                    Sendbuf = Text_Create_Proc()
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^", 0)
                    Exit Function
            End Select
        
        End If
    
    Next i

    If i > M_Gyo Then                       '���ۂ͂��肦�Ȃ��i�S���҂������́j
        ID_KANRI_TBL(ING_No).Step = Step_Start
        '   -------------------------------- �G���[���b�Z�[�W�쐬
        Call Err_Send_Proc("�S���Җ��o�^", "", "", "", "")
                    
        Sendbuf = Text_Create_Proc()
        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                    
        Tanto_Check_Proc = False
        Exit Function
    End If

'----------------------------------------------- '��p���j���[�l��
    If Menu_Type = 1 Then
                        '���ʃ��j���[
    Else
        Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
        sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
        Select Case sts
            Case BtNoErr
'                ID_KANRI_TBL(ING_No).MENU_GRP = StrConv(P_TMENUREC.TANTO_CODE, vbUnicode)
            Case BtErrKeyNotFound
                    
'                ID_KANRI_TBL(ING_No).MENU_GRP = ""
                '   -------------------------------- �G���[���b�Z�[�W�쐬
                Call Err_Send_Proc("�S���҃��j���[", "���o�^", "", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
            
                Tanto_Check_Proc = False
                Exit Function
            Case Else
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "�S���҃��j���[", 0)
                Exit Function
        End Select
            
    
    
    End If
'----------------------------------------------- '���j���[��񁕍�Ə��̏�����
    
    If UBound(JGYOBU_T) = 0 Then
                                                '�P���ƕ��Œ�
        ID_KANRI_TBL(ING_No).JGYOBU = JGYOBU_T(0).CODE
    Else
        ID_KANRI_TBL(ING_No).JGYOBU = ""
    End If
    
    If UBound(NAIGAI) = 0 Then
                                                '�����O�Œ�
        ID_KANRI_TBL(ING_No).NAIGAI = NAIGAI(0).CODE
    Else
        ID_KANRI_TBL(ING_No).NAIGAI = ""
    End If
    
    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'    ID_KANRI_TBL(ING_No).MENU_LV3 = ""

    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'    ID_KANRI_TBL(ING_No).PageNo_LV3 = 0

    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ""
    ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = ""
    ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = ""

'---------------------------------------------- '���j���[���M
    If Menu_Send_Proc(Sendbuf) Then
        Exit Function
    End If

    Tanto_Check_Proc = False

End Function


'[2016/05/14 -  mdlProc����ړ�

Public Function Menu_Send_Proc(Optional Sendbuf As String) As Integer

'-------------------------------------------------------
'
'   �w���j���[�e�L�X�g�쐬�x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Menu_Tbl()  As Menu_Tbl_tag
Dim Menu_Cnt    As Integer
Dim Max_Page    As Integer


Dim PageNo      As Integer

Dim Gyo_Suu     As Integer
Dim Start_Gyo   As Integer
Dim End_Gyo     As Integer


Dim WK_LV1      As String * 3
Dim WK_LV2      As String * 3



Dim wkHex       As String   '2017.09.07


    Menu_Send_Proc = True
'----------------------------------------------- '���ƕ��I������
    If Trim(ID_KANRI_TBL(ING_No).JGYOBU) = "" Then
        Call JGYOBU_MENU_SET

        Sendbuf = Text_Create_Proc()


        Menu_Send_Proc = False
        Exit Function
    End If
'----------------------------------------------- '�����O�I������
    If ID_KANRI_TBL(ING_No).NAIGAI = " " Then

        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        ID_KANRI_TBL(ING_No).MENU_LV2 = ""

        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0


        Call NAIGAI_MENU_SET
        Sendbuf = Text_Create_Proc
        Menu_Send_Proc = False
        Exit Function
    Else
        '2010.03.30
        If Trim(ID_KANRI_TBL(ING_No).MENU_LV1) = "" Then
            If ID_KANRI_TBL(ING_No).Step < Step_MENU1_REQ Then
                ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
            End If
'            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
        
'            ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
        
'            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ""
'            ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = ""
'            ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = ""
        End If
        '2010.03.30
    
    
    End If
    '   -------------------------------- ���x���P�@�g�b�v���j���[�̊Ǘ�
    If Trim(ID_KANRI_TBL(ING_No).MENU_LV1) = "" Then
'        ST_LOG_OUT_F = True '2008.08.08
        
        
        MENU_UP_F = False   '2008.08.08
        
        '�ƭ���ٰ��
        Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
        Erase Menu_Tbl

        sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
        Select Case sts
            
            Case BtNoErr
            Case BtErrKeyNotFound
            
                        
                Call Err_Send_Proc("���j���[���o�^", "", "", "", "")
                Sendbuf = Text_Create_Proc()
'                If UBound(NAIGAI) = 0 Then
'                    ID_KANRI_TBL(ING_No).Step = Step_Start
'                Else
'                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'                End If
                
                
                
                
                
                
                
                If UBound(NAIGAI) = 0 Then
                    
                    
                    
                    If UBound(JGYOBU_T) = 0 Then
                    
                    
                        ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                
    '                    Call Menu_Send_Proc(Sendbuf)
                
                
                    Else
                
                        '���ƕ��v���Ń��[�v����
                        ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                        ID_KANRI_TBL(ING_No).JGYOBU = ""
                        
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
            
    '                    Call Menu_Send_Proc(Sendbuf)
                    End If
                
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                    ID_KANRI_TBL(ING_No).NAIGAI = ""
    '                Call Menu_Send_Proc(Sendbuf)
                End If
                
                Menu_Send_Proc = False
                
                Exit Function
            
            
            Case Else

                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, com, "�S���ҕ��ƭ�", 0)
                Exit Function
        End Select


        Menu_Cnt = -1
        For i = 0 To 179     '29--->179 2006.10.11
            If Trim(StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)) = "" Then
                Exit For
            End If
        
            If StrConv(P_TMENUREC.MENU_T(i).JGYOBU, vbUnicode) = ID_KANRI_TBL(ING_No).JGYOBU And _
                StrConv(P_TMENUREC.MENU_T(i).NAIGAI, vbUnicode) = ID_KANRI_TBL(ING_No).NAIGAI Then
        
                Menu_Cnt = Menu_Cnt + 1
                ReDim Preserve Menu_Tbl(Menu_Cnt)
            
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)
                
                Call UniCode_Conv(K0_P_MENU.JGYOBU, StrConv(P_TMENUREC.MENU_T(i).JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_MENU.NAIGAI, StrConv(P_TMENUREC.MENU_T(i).NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_MENU.MENU_NO, StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                Select Case sts
                    
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                                
                        Call Err_Send_Proc("���j���[�ُ�", "", "", "", "")
'                        Sendbuf = Text_Create_Proc()
'                        If UBound(NAIGAI) = 0 Then
'                            ID_KANRI_TBL(ING_No).Step = Step_Start
'                        Else
'                            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'                        End If
'                        Menu_Send_Proc = False
'                        Exit Function
                    
                    
                        If UBound(NAIGAI) = 0 Then
                            
                            
                            
                            If UBound(JGYOBU_T) = 0 Then
                            
                            
                                ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        
            '                    Call Menu_Send_Proc(Sendbuf)
                        
                        
                            Else
                        
                                '���ƕ��v���Ń��[�v����
                                ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                                ID_KANRI_TBL(ING_No).JGYOBU = ""
                                
                                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                                ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                    
            '                    Call Menu_Send_Proc(Sendbuf)
                            End If
                        
                        Else
                            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                            ID_KANRI_TBL(ING_No).NAIGAI = ""
            '                Call Menu_Send_Proc(Sendbuf)
                        End If
                        
                        Menu_Send_Proc = False
                        
                        Exit Function
                    
                    
                    
                    
                    Case Else
        
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, com, "���j���[�Ǘ��}�X�^", 0)
                        Exit Function
                End Select
                
                
                
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)
                Menu_Tbl(Menu_Cnt).Disp = StrConv(P_MENUREC.MENU_DSP, vbUnicode)
        
        
            End If
        
        Next i


        If Menu_Cnt = -1 Then
        '   -------------------------------- �G���[���b�Z�[�W�쐬
            Call Err_Send_Proc("���j���[���o�^", "", "", "", "")
            Sendbuf = Text_Create_Proc()
'            If UBound(NAIGAI) = 0 Then
'                ID_KANRI_TBL(ING_No).Step = Step_Start
'            Else
'                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'            End If
            
            
            
            
            
            If UBound(NAIGAI) = 0 Then
                
                
                
                If UBound(JGYOBU_T) = 0 Then
                
                
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            
'                    Call Menu_Send_Proc(Sendbuf)
            
            
                Else
            
                    '���ƕ��v���Ń��[�v����
                    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                    ID_KANRI_TBL(ING_No).JGYOBU = ""
                    
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
        
'                    Call Menu_Send_Proc(Sendbuf)
                End If
            
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
'                Call Menu_Send_Proc(Sendbuf)
            End If
            
'''''            Menu_Send_Proc = False
            
            Exit Function
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        End If


        Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
        PageNo = ID_KANRI_TBL(ING_No).PageNo_LV1

        Start_Gyo = PageNo * M_Gyo
        End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)


        Send_Text.sts = Sts_OK                                      '�X�e�[�^�X�@OK
        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK

        Send_Text.Display_Flg = Display_MENU                        '�\����ʃt���O ���j���[���
        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
        If Max_Page = 1 Then
            Send_Text.End_Menu = Menu_Only          '�P��ʂ̂�
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
        Else
            If (Max_Page - 1) = PageNo Then
                Send_Text.End_Menu = MENU_END       '�ŏI�y�[�W
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = MENU_END
            Else
                If PageNo = 0 Then
                    Send_Text.End_Menu = Menu_Head  '�擪�y�[�W
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
                Else
                    Send_Text.End_Menu = Menu_Mid   '�r���y�[�W
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
                End If
            End If
        End If
        Send_Text.FileName = ""                                         '���M�f�[�^�t�@�C����
        ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
        Send_Text.buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
'---------------------------------------------------------------
        Gyo_Suu = 0
        j = -1
        For i = Start_Gyo To End_Gyo
            j = j + 1
            If i > UBound(Menu_Tbl) Then
                Send_Text.Box_Type(j).Box_Type = ""                 'BOX����
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
                Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '�\�����e
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
                Send_Text.Box_Type(j).INIT = ""                     '���l�����l
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                Send_Text.Box_Type(j).Start_Pos = ""                '�����J�[�\���ʒu
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                Send_Text.Box_Type(j).Max_Size = "00"               '���͌���
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                Send_Text.Box_Type(j).MENU = ""                     '���j���\�ԍ�
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
            Else
                Gyo_Suu = Gyo_Suu + 1
                Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX����
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
                Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                Send_Text.Box_Type(j).INIT = ""                     '���l�����l
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                Send_Text.Box_Type(j).Start_Pos = ""                '�����J�[�\���ʒu
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                Send_Text.Box_Type(j).Max_Size = "00"               '���͌���
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO
                
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO
            End If
        Next i
        
        Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      '���j���[���ڐ�
        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
        ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M
        Sendbuf = Text_Create_Proc()
        
    Else
        '   -------------------------------- ���x���Q�@��ƃ��j���[�̊Ǘ�
        If Trim(ID_KANRI_TBL(ING_No).MENU_LV2) = "" Then
            
            
'2008.08.12            If ST_LOG_OUT_F Then            '2008.08.08

'                ST_LOG_OUT_F = False        '2008.08.08

                
'                If Not MENU_UP_F Then       '2008.08.08
'
'                    If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
'                                                Format(ID_KANRI_TBL(ING_No).ID, "000"), _
'                                                ID_KANRI_TBL(ING_No).JGYOBU, _
'                                                ID_KANRI_TBL(ING_No).NAIGAI, _
'                                                ID_KANRI_TBL(ING_No).MENU_LV1, _
'                                                "ST", , , , , , , , , FILE_RETRY) Then
'
'                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                        Sendbuf = Text_Create_Proc()
'                        Exit Function
'
'                    End If
'
'                End If


 '2008.08.12           End If
            
            
            
            
            
            
            
            
            
            '�ƭ���ٰ��
            Call UniCode_Conv(K0_P_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
            Call UniCode_Conv(K0_P_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
            Call UniCode_Conv(K0_P_MENU.MENU_NO, ID_KANRI_TBL(ING_No).MENU_LV1)
            
            
            Erase Menu_Tbl
    
            sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
            Select Case sts
                
                Case BtNoErr
                Case BtErrKeyNotFound
                
                            
                    Call Err_Send_Proc("���j���[���o�^", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    If UBound(NAIGAI) = 0 Then
                        ID_KANRI_TBL(ING_No).Step = Step_Start
                    Else
                        ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                    End If
                    Menu_Send_Proc = False
                    Exit Function
                
                
                Case Else
    
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, com, "�S���ҕ��ƭ�", 0)
                    Exit Function
            End Select
    
    
            Menu_Cnt = -1
            For i = 0 To 19
                If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = "" Then
                    Exit For
                End If
            
            
                Menu_Cnt = Menu_Cnt + 1
                ReDim Preserve Menu_Tbl(Menu_Cnt)
                
                    
                    
                Menu_Tbl(Menu_Cnt).MENU_NO = StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)
                Menu_Tbl(Menu_Cnt).PARAM = StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode)
                Menu_Tbl(Menu_Cnt).Disp = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
            
                Menu_Tbl(Menu_Cnt).LOG_OUT = StrConv(P_MENUREC.SAGYO(i).LOG_OUT, vbUnicode)
            
            Next i
    
    
            If Menu_Cnt = -1 Then
            '   -------------------------------- �G���[���b�Z�[�W�쐬
                Call Err_Send_Proc("���j���[���o�^", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                If UBound(NAIGAI) = 0 Then
                    ID_KANRI_TBL(ING_No).Step = Step_Start
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                End If
                Menu_Send_Proc = False
                Exit Function
            End If
    
    
            Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV2
    
            Start_Gyo = PageNo * M_Gyo
            End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)
    
    
            Send_Text.sts = Sts_OK                                      '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_MENU                        '�\����ʃt���O ���j���[���
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
            If Max_Page = 1 Then
                Send_Text.End_Menu = Menu_Only          '�P��ʂ̂�
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
            Else
                If (Max_Page - 1) = PageNo Then
                    Send_Text.End_Menu = MENU_END       '�ŏI�y�[�W
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = MENU_END
                Else
                    If PageNo = 0 Then
                        Send_Text.End_Menu = Menu_Head  '�擪�y�[�W
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
                    Else
                        Send_Text.End_Menu = Menu_Mid   '�r���y�[�W
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
                    End If
                End If
            End If
            Send_Text.FileName = ""                                         '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.FileName = ""
            Send_Text.buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
    '---------------------------------------------------------------
            Gyo_Suu = 0
            j = -1
            For i = Start_Gyo To End_Gyo
                j = j + 1
                If i > UBound(Menu_Tbl) Then
                    Send_Text.Box_Type(j).Box_Type = ""                 'BOX����
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
                    Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '�\�����e
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
                    Send_Text.Box_Type(j).INIT = ""                     '���l�����l
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                    Send_Text.Box_Type(j).Start_Pos = ""                '�����J�[�\���ʒu
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                    Send_Text.Box_Type(j).Max_Size = "00"               '���͌���
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                    Send_Text.Box_Type(j).MENU = ""                     '���j���\�ԍ�
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
                
                    Send_Text.Box_Type(j).MENU18 = ""                     '���j���\�ԍ� 2017.09.07
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU18 = ""              '2017.09.07
                
                
                Else
                    Gyo_Suu = Gyo_Suu + 1
                    Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX����
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
                    Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
                    Send_Text.Box_Type(j).INIT = ""                     '���l�����l
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
                    Send_Text.Box_Type(j).Start_Pos = ""                '�����J�[�\���ʒu
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
                    Send_Text.Box_Type(j).Max_Size = "00"               '���͌���
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
                                                                        
                                                                        
'>>>>>>>>>>>>>>>>>>>    2017.09.07
                                                                        '���j���\�ԍ� & ���Ұ�
'                     Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM                              '2017.09.07
'                     ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM         '2017.09.07
                
                     
                     
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MTS_FLG = ""         '2017.10.27
                     
                    wkHex = f10sinTo16sin(Menu_Tbl(i).PARAM)
                     
                     
                     
                    If Trim(wkHex) = "" Then
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MTS_FLG = "1"    '2017.10.27
                    
                        wkHex = Menu_Tbl(i).PARAM
                    End If
                    Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & wkHex                                           '2017.09.07
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & wkHex                      '2017.09.07

                    
'>>>>>>>>>>>>>>>>>>>    2017.09.07
                
                
                    Send_Text.Box_Type(j).MENU18 = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM                          '���j���\�ԍ� 2017.09.07
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU18 = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM     '2017.09.07
                
                End If
            Next i
            
            Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      '���j���[���ڐ�
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
            ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M
            Sendbuf = Text_Create_Proc()
            
        
        End If
    End If
    
    Menu_Send_Proc = False




End Function

'[2016/05/14 -  mdlProc����ړ�

Public Function Menu_Recv_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   �w�Q�K�w�ȏ�̃��j���[���M�x
'
'-------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    
Dim MTS     As String * 8
Dim SS      As String * 8
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_MENU1_RES
    '   -------------------------------- �����ق�
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> BEF_Page And Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> NEXT_Page Then

                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
            End If
            
            If Menu_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
        Case Else
    
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = BEF_Page Or Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) = NEXT_Page Then
                If Menu_Send_Proc() Then
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                End If
            Else
    
    '   -------------------------------- ���j���[�Ǘ��}�X�^�Ǎ���
                Call UniCode_Conv(K0_P_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_P_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_P_MENU.MENU_NO, ID_KANRI_TBL(ING_No).MENU_LV1)
        
                sts = BTRV(BtOpGetEqual, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
                Select Case sts
                    Case BtNoErr
    
                        For i = 0 To 19
                        
                            If Trim(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode)) = Trim(ID_KANRI_TBL(ING_No).MENU_LV2) And _
                               ((StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode) = (ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)) Or _
                                (Trim(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode)) = (Trim(Format(ID_KANRI_TBL(ING_No).MTS_CODE, "#0") & ID_KANRI_TBL(ING_No).SS_CODE)))) Then
                                                                
                                '>>>>>>>>>  2017.11.30
                                If (Trim(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode)) = (Trim(Format(ID_KANRI_TBL(ING_No).MTS_CODE, "#0") & ID_KANRI_TBL(ING_No).SS_CODE))) Then
                                    ID_KANRI_TBL(ING_No).MTS_CODE = Trim(Format(ID_KANRI_TBL(ING_No).MTS_CODE, "#0"))
                                End If
                                '>>>>>>>>>  2017.11.30
                                
                                Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                
                                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                                Select Case sts
                                
                                    Case BtNoErr
                                        '���ŕ\������
                                        ID_KANRI_TBL(ING_No).YOIN_DNAME = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
                                
                                
                                        '2010.09.15
                                        ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
                                
                                        
                                        ID_KANRI_TBL(ING_No).SAGYO_LOG = StrConv(P_MENUREC.SAGYO(i).LOG_OUT, vbUnicode)
                                                                    
                                        If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then   '������Ȃ�i�o�ׁj
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
'2006.01.30                                            ID_KANRI_TBL(ING_No).MTS_CODE = Left(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
'2006.01.30                                            ID_KANRI_TBL(ING_No).SS_CODE = Right(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
                                        End If
                                                                                            '���i�i������w��j�Ȃ�
                                        If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_MTS Then
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        End If
                                        
                                                                                            '���i�i�����w��j�Ȃ�   2016.10.14
                                        If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_Drct Then
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        End If
                                        
                                        
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(YOINREC.CODE_TYPE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(YOINREC.SOKO_NO, vbUnicode)
                                
                                    Case BtErrKeyNotFound
    
                                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                        Call Err_Send_Proc("�v���}�X�^", "���o�^", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                                        Menu_Recv_Proc = False
                                        Exit Function
                                    Case Else
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Call File_Error(sts, BtOpGetEqual, "���j���[�Ǘ�", 0)
                                        Exit Function
                                End Select
                                
                                Exit For
                            End If
                        
                        Next i
    
                        If i > 19 Then
                
    
                            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc("���j���[�Ǘ��}�X�^", "�ݒ�~�X", Trim(ID_KANRI_TBL(ING_No).MENU_LV1), "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                            Menu_Recv_Proc = False
                            Exit Function
                        
                        End If
                
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
        
        
                    Case BtErrKeyNotFound
        
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc("���j���[�Ǘ��}�X�^", "���o�^", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
                        Menu_Recv_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "���j���[�Ǘ�", 0)
                        Exit Function
        
                End Select
            End If
        End Select
        Sendbuf = Text_Create_Proc()
    
    Menu_Recv_Proc = False

End Function
'2006.01.30Private Function Menu_Recv_Proc(Sendbuf As String) As Integer
'2006.01.30'-------------------------------------------------------
'2006.01.30'
'2006.01.30'   �w�Q�K�w�ȏ�̃��j���[���M�x
'2006.01.30'
'2006.01.30'-------------------------------------------------------
'2006.01.30Dim sts As Integer
'2006.01.30
'2006.01.30    Menu_Recv_Proc = True
'2006.01.30                                        '���j���Ǘ��}�X�^�̓ǂݍ���
'2006.01.30    Call UniCode_Conv(K1_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30    Call UniCode_Conv(K1_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_GRP_NO, ID_KANRI_TBL(ING_No).MENU_GRP)
'2006.01.30
'2006.01.30    If ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ Then
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, "000")
'2006.01.30    Else
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, ID_KANRI_TBL(ING_No).MENU_LV1)
'2006.01.30    End If
'2006.01.30
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_LV2, ID_KANRI_TBL(ING_No).MENU_LV2)
'2006.01.30    Call UniCode_Conv(K1_MENU.MENU_LV3, ID_KANRI_TBL(ING_No).MENU_LV3)
'2006.01.30    sts = BTRV(BtOpGetEqual, MENU_POS, MENUREC, Len(MENUREC), K1_MENU, Len(K1_MENU), 1)
'2006.01.30    Select Case sts
'2006.01.30        Case BtNoErr
'2006.01.30        Case BtErrKeyNotFound
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30            '   -------------------------------- �G���[���b�Z�[�W�쐬
'2006.01.30            Call Err_Send_Proc("���j���[�Ǘ�", "���o�^", "", "", "")
'2006.01.30
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'2006.01.30
'2006.01.30            Menu_Recv_Proc = False
'2006.01.30            Exit Function
'2006.01.30        Case Else
'2006.01.30            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Call File_Error(sts, BtOpGetEqual, "���j���[�Ǘ�", 0)
'2006.01.30            Exit Function
'2006.01.30    End Select
'2006.01.30
'2006.01.30    If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> BEF_Page And Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> NEXT_Page And _
'2006.01.30     StrConv(MENUREC.MENU_KBN, vbUnicode) = "1" Then
'2006.01.30
'2006.01.30
'2006.01.30                                            '�v���̓Ǎ���
'2006.01.30        Call UniCode_Conv(K0_YOIN.CODE_TYPE, StrConv(MENUREC.CODE_TYPE, vbUnicode))
'2006.01.30        Call UniCode_Conv(K0_YOIN.YOIN_CODE, StrConv(MENUREC.YOIN_CODE, vbUnicode))
'2006.01.30        sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30                ID_KANRI_TBL(ING_No).YOIN_DNAME = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
'2006.01.30
'2006.01.30                If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then   '������Ȃ�i�o�ׁj
'2006.01.30                    ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(MENUREC.YOIN_CODE, vbUnicode)
'2006.01.30                    ID_KANRI_TBL(ING_No).MTS_CODE = Left(StrConv(MENUREC.PARAM, vbUnicode), 8)
'2006.01.30                    ID_KANRI_TBL(ING_No).SS_CODE = Right(StrConv(MENUREC.PARAM, vbUnicode), 8)
'2006.01.30                End If
'2006.01.30                                                                    '���i�i������w��j�Ȃ�
'2006.01.30                If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_MTS Then
'2006.01.30                    ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(MENUREC.YOIN_CODE, vbUnicode)
'2006.01.30                End If
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(YOINREC.CODE_TYPE, vbUnicode)
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(YOINREC.YOIN_CODE, vbUnicode)
'2006.01.30                ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(MENUREC.PARAM, vbUnicode)
'2006.01.30
'2006.01.30            Case BtErrKeyNotFound
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30                '   -------------------------------- �G���[���b�Z�[�W�쐬
'2006.01.30                Call Err_Send_Proc("�v���}�X�^", "���o�^", "", "", "")
'2006.01.30
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'2006.01.30
'2006.01.30                Menu_Recv_Proc = False
'2006.01.30                Exit Function
'2006.01.30            Case Else
'2006.01.30                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, BtOpGetEqual, "���j���[�Ǘ�", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30        If Sagyo_Send_Proc() Then
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Exit Function
'2006.01.30        End If
'2006.01.30    Else
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
'2006.01.30
'2006.01.30        If Menu_Send_Proc() Then
'2006.01.30            Sendbuf = Text_Create_Proc()
'2006.01.30            Exit Function
'2006.01.30        End If
'2006.01.30    End If
'2006.01.30
'2006.01.30    Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Recv_Proc = False
'2006.01.30
'2006.01.30End Function


'2006.01.30Private Function Menu_Send_Proc(Optional Sendbuf As String) As Integer
'2006.01.30'-------------------------------------------------------
'2006.01.30'
'2006.01.30'   �w���j���[�e�L�X�g�쐬�x
'2006.01.30'
'2006.01.30'-------------------------------------------------------
'2006.01.30Dim sts         As Integer
'2006.01.30Dim com         As Integer
'2006.01.30
'2006.01.30Dim i           As Integer
'2006.01.30Dim j           As Integer
'2006.01.30
'2006.01.30Dim Menu_Tbl()  As Menu_Tbl_tag
'2006.01.30Dim Menu_Cnt    As Integer
'2006.01.30Dim Max_Page    As Integer
'2006.01.30
'2006.01.30
'2006.01.30Dim PageNo      As Integer
'2006.01.30
'2006.01.30Dim Gyo_Suu     As Integer
'2006.01.30Dim Start_Gyo   As Integer
'2006.01.30Dim End_Gyo     As Integer
'2006.01.30
'2006.01.30
'2006.01.30Dim WK_LV1      As String * 3
'2006.01.30Dim WK_LV2      As String * 3
'2006.01.30Dim WK_LV3      As String * 3
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Send_Proc = True
'2006.01.30'----------------------------------------------- '���ƕ��I������
'2006.01.30    If ID_KANRI_TBL(ING_No).JGYOBU = " " Then
'2006.01.30        Call JGYOBU_MENU_SET
'2006.01.30
'2006.01.30        Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30'----------------------------------------------- '�����O�I������
'2006.01.30    If ID_KANRI_TBL(ING_No).NAIGAI = " " Then
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        Call NAIGAI_MENU_SET
'2006.01.30        Sendbuf = Text_Create_Proc
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30'----------------------------------------------- '���j���[�Ǘ��̓ǂݍ���
'2006.01.30    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_GRP)) = 0 Then
'2006.01.30                                    '�����Ŗ��m��Ȃ̂͋��ʃ��j���[������
'2006.01.30        Call UniCode_Conv(K1_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30        Call UniCode_Conv(K1_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_GRP_NO, ALL_MENU_GRP)
'2006.01.30
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV1, "")
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV2, "")
'2006.01.30        Call UniCode_Conv(K1_MENU.MENU_LV3, "")
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        sts = BTRV(BtOpGetGreaterEqual, MENU_POS, MENUREC, Len(MENUREC), K1_MENU, Len(K1_MENU), 1)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30            Case BtErrEOF
'2006.01.30
'2006.01.30            '   -------------------------------- �G���[���b�Z�[�W�쐬
'2006.01.30                Call Err_Send_Proc("���j���[���o�^", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                If UBound(NAIGAI) = 0 Then
'2006.01.30                    ID_KANRI_TBL(ING_No).Step = Step_Start
'2006.01.30                Else
'2006.01.30                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'2006.01.30                End If
'2006.01.30                Menu_Send_Proc = False
'2006.01.30                Exit Function
'2006.01.30            Case Else
'2006.01.30                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, BtOpGetEqual, "���j���[�Ǘ��}�X�^", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30        ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
'2006.01.30
'2006.01.30
'2006.01.30    End If
'2006.01.30    '   -------------------------------- ���j���[�Ǘ��}�X�^�Ǎ���
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_GRP_NO, ID_KANRI_TBL(ING_No).MENU_GRP)
'2006.01.30    Call UniCode_Conv(K0_MENU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'2006.01.30    Call UniCode_Conv(K0_MENU.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV1, ID_KANRI_TBL(ING_No).MENU_LV1)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV2, ID_KANRI_TBL(ING_No).MENU_LV2)
'2006.01.30    Call UniCode_Conv(K0_MENU.MENU_LV3, ID_KANRI_TBL(ING_No).MENU_LV3)
'2006.01.30
'2006.01.30    Erase Menu_Tbl
'2006.01.30
'2006.01.30    com = BtOpGetGreater
'2006.01.30
'2006.01.30    Menu_Cnt = -1
'2006.01.30    Do
'2006.01.30        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
'2006.01.30        Select Case sts
'2006.01.30            Case BtNoErr
'2006.01.30                If ID_KANRI_TBL(ING_No).MENU_GRP <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Or _
'2006.01.30                    ID_KANRI_TBL(ING_No).JGYOBU <> StrConv(MENUREC.JGYOBU, vbUnicode) Or _
'2006.01.30                    ID_KANRI_TBL(ING_No).NAIGAI <> StrConv(MENUREC.NAIGAI, vbUnicode) Then
'2006.01.30                    Exit Do
'2006.01.30                End If
'2006.01.30
'2006.01.30                WK_LV1 = ID_KANRI_TBL(ING_No).MENU_LV1
'2006.01.30                WK_LV2 = ID_KANRI_TBL(ING_No).MENU_LV2
'2006.01.30                WK_LV3 = ID_KANRI_TBL(ING_No).MENU_LV3
'2006.01.30
'2006.01.30
'2006.01.30                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV1)) = 0 Then
'2006.01.30                    WK_LV1 = StrConv(MENUREC.MENU_LV1, vbUnicode)
'2006.01.30                    WK_LV2 = ""
'2006.01.30                    WK_LV3 = ""
'2006.01.30
'2006.01.30
'2006.01.30'                    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
'2006.01.30                    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30                    ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30                Else
'2006.01.30                    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) = 0 Then
'2006.01.30                        WK_LV2 = StrConv(MENUREC.MENU_LV2, vbUnicode)
'2006.01.30                        WK_LV3 = ""
'2006.01.30
'2006.01.30'                        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.30                        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30
'2006.01.30                    Else
'2006.01.30                        If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) = 0 Then
'2006.01.30                            WK_LV3 = StrConv(MENUREC.MENU_LV3, vbUnicode)
'2006.01.30
'2006.01.30'                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30
'2006.01.30                        End If
'2006.01.30                    End If
'2006.01.30                End If
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30            Case BtErrEOF
'2006.01.30                Exit Do
'2006.01.30            Case Else
'2006.01.30
'2006.01.30                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'2006.01.30                Sendbuf = Text_Create_Proc()
'2006.01.30                Call File_Error(sts, com, "���j���[�Ǘ��}�X�^", 0)
'2006.01.30                Exit Function
'2006.01.30        End Select
'2006.01.30
'2006.01.30
'2006.01.30    '   -------------------------------- ���j���[����ۑ�
'2006.01.30        If WK_LV1 <> StrConv(MENUREC.MENU_LV1, vbUnicode) Or _
'2006.01.30            WK_LV2 <> StrConv(MENUREC.MENU_LV2, vbUnicode) Or _
'2006.01.30            WK_LV3 <> StrConv(MENUREC.MENU_LV3, vbUnicode) Then
'2006.01.30        Else
'2006.01.30
'2006.01.30
'2006.01.30            Menu_Cnt = Menu_Cnt + 1
'2006.01.30
'2006.01.30            ReDim Preserve Menu_Tbl(Menu_Cnt)
'2006.01.30
'2006.01.30            If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV1)) = 0 Then
'2006.01.30                Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV1, vbUnicode)
'2006.01.30            Else
'2006.01.30                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) = 0 Then
'2006.01.30                    Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV2, vbUnicode)
'2006.01.30                Else
'2006.01.30                    If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) = 0 Then
'2006.01.30                        Menu_Tbl(Menu_Cnt).CODE = StrConv(MENUREC.MENU_LV3, vbUnicode)
'2006.01.30                    End If
'2006.01.30                End If
'2006.01.30            End If
'2006.01.30
'2006.01.30
'2006.01.30            Menu_Tbl(Menu_Cnt).Disp = StrConv(MENUREC.DISPLAY_ITEM, vbUnicode)
'2006.01.30        End If
'2006.01.30
'2006.01.30        com = BtOpGetNext
'2006.01.30
'2006.01.30    Loop
'2006.01.30
'2006.01.30    If Menu_Cnt = -1 Then
'2006.01.30            '   -------------------------------- �G���[���b�Z�[�W�쐬
'2006.01.30        Call Err_Send_Proc("���j���[���o�^", "", "", "", "")
'2006.01.30        Sendbuf = Text_Create_Proc()
'2006.01.30        If UBound(NAIGAI) = 0 Then
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_Start
'2006.01.30        Else
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
'2006.01.30        End If
'2006.01.30        Menu_Send_Proc = False
'2006.01.30        Exit Function
'2006.01.30    End If
'2006.01.30
'2006.01.30'----------------------------------------------- '���j���[���M�e�L�X�g�쐬
'2006.01.30'''    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1   '���j���[���M
'2006.01.30    '---------------------------------------------------------------
'2006.01.30    Max_Page = Int(CDbl((Menu_Cnt + 1) / M_Gyo + 0.9))
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30'    Start_Gyo = ID_KANRI_TBL(ING_No).PageNo * M_Gyo
'2006.01.30'    End_Gyo = (ID_KANRI_TBL(ING_No).PageNo * M_Gyo) + (M_Gyo - 1)
'2006.01.30
'2006.01.30
'2006.01.30    Select Case ID_KANRI_TBL(ING_No).Step
'2006.01.30        Case Step_MENU1_REQ, Step_MENU1_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV1
'2006.01.30        Case Step_MENU2_REQ, Step_MENU2_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV2
'2006.01.30        Case Step_MENU3_REQ, Step_MENU3_RES
'2006.01.30            PageNo = ID_KANRI_TBL(ING_No).PageNo_LV3
'2006.01.30    End Select
'2006.01.30
'2006.01.30
'2006.01.30    Start_Gyo = PageNo * M_Gyo
'2006.01.30    End_Gyo = (PageNo * M_Gyo) + (M_Gyo - 1)
'2006.01.30
'2006.01.30
'2006.01.30    Send_Text.sts = Sts_OK                                      '�X�e�[�^�X�@OK
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
'2006.01.30
'2006.01.30    Send_Text.Display_Flg = Display_MENU                        '�\����ʃt���O ���j���[���
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
'2006.01.30                                                                '�ŏI���j���[�t���O
'2006.01.30    If Max_Page = 1 Then
'2006.01.30        Send_Text.End_Menu = Menu_Only          '�P��ʂ̂�
'2006.01.30        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
'2006.01.30    Else
'2006.01.30        If (Max_Page - 1) = PageNo Then
'2006.01.30            Send_Text.End_Menu = Menu_End       '�ŏI�y�[�W
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_End
'2006.01.30        Else
'2006.01.30            If PageNo = 0 Then
'2006.01.30                Send_Text.End_Menu = Menu_Head  '�擪�y�[�W
'2006.01.30                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Head
'2006.01.30            Else
'2006.01.30                Send_Text.End_Menu = Menu_Mid   '�r���y�[�W
'2006.01.30                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Mid
'2006.01.30            End If
'2006.01.30        End If
'2006.01.30    End If
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Send_Text.fileName = ""                                         '���M�f�[�^�t�@�C����
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
'2006.01.30
'2006.01.30    Send_Text.Buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
'2006.01.30    '---------------------------------------------------------------
'2006.01.30    Gyo_Suu = 0
'2006.01.30    j = -1
'2006.01.30    For i = Start_Gyo To End_Gyo
'2006.01.30        j = j + 1
'2006.01.30        If i > UBound(Menu_Tbl) Then
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Box_Type = ""                 'BOX����
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = ""
'2006.01.30
'2006.01.30            Call UniCode_Conv(Send_Text.Box_Type(j).LCD, "")    '�\�����e
'2006.01.30            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, "")
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).INIT = ""                     '���l�����l
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Start_Pos = ""                '�����J�[�\���ʒu
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Max_Size = "00"               '���͌���
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).MENU = ""                     '���j���\�ԍ�
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = ""
'2006.01.30
'2006.01.30
'2006.01.30        Else
'2006.01.30            Gyo_Suu = Gyo_Suu + 1
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Box_Type = TYPE_MENU          'BOX����
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Box_Type = TYPE_MENU
'2006.01.30                                                                '�\�����e
'2006.01.30            Call UniCode_Conv(Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
'2006.01.30            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).LCD, Menu_Tbl(i).Disp)
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).INIT = ""                     '���l�����l
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).INIT = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Start_Pos = ""                '�����J�[�\���ʒu
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Start_Pos = ""
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).Max_Size = "00"               '���͌���
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).Max_Size = "00"
'2006.01.30
'2006.01.30            Send_Text.Box_Type(j).MENU = Menu_Tbl(i).CODE       '���j���\�ԍ�
'2006.01.30            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).CODE
'2006.01.30
'2006.01.30        End If
'2006.01.30
'2006.01.30    Next i
'2006.01.30
'2006.01.30    Send_Text.Menu_Suu = Format(Gyo_Suu, "00")      '���j���[���ڐ�
'2006.01.30    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(Gyo_Suu, "00")
'2006.01.30
'2006.01.30
'2006.01.30    ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M
'2006.01.30
'2006.01.30    Sendbuf = Text_Create_Proc()
'2006.01.30
'2006.01.30
'2006.01.30
'2006.01.30    Menu_Send_Proc = False
'2006.01.30
'2006.01.30End Function


'[2016/05/14 -  mdlProc����ړ�
Public Function Sagyo_Main_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   �w��Ǝ�M���̃��C�������x
'
'-------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Found_Flg   As Boolean
    
    
    Sagyo_Main_Proc = True
    
    
    Found_Flg = False
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '�ŏ��͂Q���Ō���
            If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = _
                WEL_Para_Tbl(i, j).Action Then
                Found_Flg = True
                Exit For
            End If
        
        Next j
            
        If Found_Flg Then
            Exit For
        End If
    
    Next i
    
    If Not Found_Flg Then
        
        For i = 0 To UBound(WEL_Para_Tbl, 1)
            For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '�Ō�͂P���Ō���
               If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = Left(WEL_Para_Tbl(i, j).Action, 1) Then
                    Found_Flg = True
                    Exit For
                End If
        
            Next j
            
            If Found_Flg Then
                Exit For
            End If
        
        
        Next i
            
    End If


    If Not Found_Flg Then
                        
                        '���肦�Ȃ��ُ�i�Y����ƃp�����[�^�Ȃ��j
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_GRP = ""
                
        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                
        '   -------------------------------- �G���[���b�Z�[�W�쐬
        Call Err_Send_Proc("��ƃp�����[�^�iINI�j", "���o�^", "", "", "")
                    
        Sendbuf = Text_Create_Proc()
        ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
            
        Sagyo_Main_Proc = False
        Exit Function
    
    End If

    Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE
        Case ACT_ZAITEI_IN          '�ݒ��{
        
        
            If Zaitei_In_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_ZAITEI_OUT         '�ݒ��|
            
            '2007.10.02
            If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = _
                Wel_S_SHOUHI Then
                        
                '2007.10.02 ���ޏ����p
                If S_SHOUHI_Out_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
                        
                        
            Else
                
                '2015.02.21 ���ޏ���(�V)
                If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = _
                    Wel_S_SHOUHI2 Then
                    If S_SHOUHI_Out2_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Else
                
                
                
                    If Zaitei_Out_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                End If
            End If
        
        
        Case ACT_NYUKA              '����
    
        Case ACT_SYUKA_KEI          '�o��(�o�ח\��L��)��������錾
        
            
            If MTS_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_SYUKA_HYO          '�o��(�o�ɕ\)
        
            If SYUKO_HYO_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        Case ACT_SYUKA_GAI          '�o��(�o�ח\�薳��)
        
            If Out_Plan_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        
        Case ACT_IDO_IN             '�ړ�����
        
'-----------------------------------------  2012.03.06
'            If Ido_In_Proc(Sendbuf, i, j) Then
'                Exit Function
'            End If
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
        
                Case Wel_IDO_IN_OSAKA          '2012.03.15
            
                    If Ido_In_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Else
            
                    If Ido_In_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            End Select
        
        
        
        Case ACT_IDO_OUT            '�ړ��o��
        
            
            
'-----------------------------------------  2012.03.06
'            '2011.06.01
'            If (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE) = _
'                Wel_HIN_FURIKAE_MAINA Then
'                '2011.06.01
'                If Ido_Out_Hin_Furikae_Proc(Sendbuf, i, j) Then
'                    Exit Function
'                End If
'
'            Else
'                If Ido_Out_Proc(Sendbuf, i, j) Then
'                    Exit Function
'                End If
'
'            End If
        
        Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
            Case Wel_HIN_FURIKAE_MAINA
        
                If Ido_Out_Hin_Furikae_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
        
        
            Case Wel_IDO_OUT_OSAKA          '2012.03.10
            
                If Ido_Out_OSAKA_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
            
            Case Wel_IDO_OUT_OSAKA2         '2014.11.07
            
                If Ido_Out_OSAKA_NEW_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
            
            
            Case Wel_IDO_OUT_OSAKA3         '2016.05.11
            
                If Ido_Out_OSAKA_NEW2_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
            
            
            
            
            Case Else
                If Ido_Out_Proc(Sendbuf, i, j) Then
                    Exit Function
                End If
        
        End Select
'-----------------------------------------  2012.03.06
        
        Case ACT_DENPYO_ID          '�`�[�h�c
        
            If DEN_ID_Dec_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_DENPYO_ID2          '*�`�[�h�c
        
            If DEN_ID_Dec2_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_KENPIN             '���i
        
            If Inspe_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_WEL_ETC            'WEL��p�i�Ɖ�j
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
            
                Case Wel_TANAOROSI      '�uWEL �I�����v�̗v��
                
                    If Tanaorosi_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_TANAHYOJI      '�uWEL �I�ԕ\���v�̗v��
                
                    If Tanahyoji_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                
                Case Wel_HIN_SHOGO      '�uWEL �i�ԕʏƍ��v�̗v��
                    
                    If Hin_Shogo_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_AVE_SYUKA      '�uWEL �����Ϗo�א��v�̗v��
                
                    If Ave_Syuka_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_HOST_ZAIKO     '�uWEL �z�X�g�݌ɏƉ�v�̗v��
                    
                    If Host_Zaiko_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                            
                Case Wel_ST_TANABAN     '�uWEL �W���I�Ԑݒ�v�̗v��
                    
                    If St_Tanaban_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                            
            
                Case Wel_RIREKI         '�uWEL �����o�ɗ����v�̗v��
                
                    If Rireki_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_SUII           '�uWEL �o�א��ځv�̗v��

                    If Suii_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_TANA_HIN_SHOGO '�uWEL �I�ԁE�i�ԕʏƍ��v�̗v��

                    If Tana_Hin_Shogo_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_TANAHYOJI_KASO '�uWEL �I�ԕ\��(���z�D��)�v�̗v��
                
                    If Tanahyoji_Kaso_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_COMPO          '�uWEL �\���\���v�̗v�� 2006.10.15
                
                    If COMPO_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_KEN_RIREKI     '�uWEL ���i���сv�̗v�� 2006.10.15
                
                    If KEN_Rireki_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_RIREKI2        '�uWEL �����o�ɗ����v�̗v�� 2009.01.09
                
                    If Rireki2_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_AVE_SYUKA_ID   '�uWEL �����ρ@�h�c�ǂݍ��݁v�̗v�� 2009.03.19
                
                    If Ave_Syuka_ID_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_KEN_ZAN_ID   '�uWEL �W�񍫕�c�v�̗v�� 2010.02.15
            
            
                    If KEN_ZAN_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '2010.12.13
                Case Wel_AKI_LOC                '�uWEL �󂫃��P�[�V�����̌����v�̗v��
            
            
                        
                    If AKI_LOC_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '2011.07.05
                Case Wel_S_AVE_SYUKA
            
                    If S_Ave_Syuka_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            End Select
    
        Case ACT_KENPIN_MTS             '���i�i�l�s�r�ǂݍ��݂���j
        
            If Inspe_Proc_MTS(Sendbuf, i, j) Then
                Exit Function
            End If
    
    
        Case ACT_GOODS_ONFF             '���i�^�����i�؂�ւ�
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
    
                Case Wel_GOODS_ONOFF_ONO        '�uWEL ���i/�����i�؂�ւ��@����v�̗v��
                
                    If GOODS_ONOFF_Ono_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_GOODS_ONOFF_SIGA       '�uWEL ���i/�����i�؂�ւ��@����v�̗v��
            
                    If GOODS_ONOFF_Siga_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            End Select
    
    
        Case ACT_SPECIAL_PROCESS    '���ꏈ��
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
    
                Case Wel_RETURNED_GOODS         '�u�Ǖi�ԕi�v�̗v��
                
                    If RETURNED_GOODS_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                
                Case Wel_LOCATION_MOVE         '�u�I�ړ��v�̗v��
                
                
                    If Location_Move_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
            
                Case Wel_RETURNED_GOODS_OSAKA   '�u���o�b�@�Ǖi�ԕi�v�̗v��   2007.09.12
                
                    If RETURNED_GOODS_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_SYUKA_CENCEL           '�u�o�׷�ݾفv�̗v��    2007.11.02
            
                    If SYUKA_CANCEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_ZAIKO_SEISA            '�u�݌ɐ����v�̗v��    2008.11.20
            
                    If Zaiko_Seisa_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_SYUKO_CANCEL           '�u�o��CANCEL�v�̗v��    2008.12.05
            
                    If SYUKO_CANCEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                                                '�u�ڊǏ����v�̗v��    2009.02.26
                Case Wel_IKAN_1, Wel_IKAN_2, Wel_IKAN_3
            
                    
                    Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                        
                        Case Wel_IKAN_1
                    
                            If LSBU_IKAN_Proc(Sendbuf, i, j, 1) Then
                                Exit Function
                            End If
            
                        Case Wel_IKAN_2
                    
                            If LSBU_IKAN_Proc(Sendbuf, i, j, 2) Then
                                Exit Function
                            End If
            
                        Case Wel_IKAN_3
                    
                            If LSBU_IKAN_Proc(Sendbuf, i, j, 3) Then
                                Exit Function
                            End If
            
            
                    End Select
            
            
            
            
                                                '�u�ڊǍ폜�����v�̗v��    2009.03.09
                Case Wel_IKAN_DEL
            
                    
                
                    If LSBU_IKAN_DEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '�u�ڊǕ\�������v�̗v��    2009.03.09
                Case Wel_IKAN_DSP
            
                    
                
                    If LSBU_IKAN_DSP_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '�u�I�Ԓǉ������v�̗v��    2009.03.17
                Case Wel_TANA_INS
            
                    
                
                    If TANA_INS_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                '�u�o�׍폜�����v�̗v��    2009.03.17
                Case Wel_SYUKA_DEL
            
                    
                
                    If SYUKA_DEL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
                                                '�u�ː��^���������v�̗v��    2010.03.09
                Case Wel_SAI_SU
            
                    
                
                    If SAI_SU_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            
                                                '2010.12.13
                Case Wel_TANA_USE               '�uWEL �I�g�p�󋵁v�̗v��
                                                
                                                
                                                
                    If TANA_USE_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                                                
                                                
                                                    
                                                
                                                '2011.03.05
                Case Wel_LABEL_PRINT            '�uWEL ���ٔ��s�v�̗v��
            
            
                    
                    
            
            
            
                    If LABEL_Print_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                                                '2011.08.05
                Case Wel_JAN_SET                '�uWEL �i��(JAN)�o�^�v�̗v��
            
                    If JAN_SET_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            
                Case Wel_T_back                 '2015.01.22 ���揈��(�L��)
            
            
                    If T_back_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
                                                '2015.10.06
                Case Wel_LABEL_PRINT_CNT    '�uWEL ���ٔ��s �����w��v�̗v��
            
            
                    
                    
            
            
            
                    If LABEL_Print_Cnt_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            End Select
        
        Case ACT_KENPIN_DEN             '���i�i���o�b�����j   2006.12.07
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
'                Case Wel_Inspe_DEN, Wel_Inspe_DEN2              '2009.06.03
                Case Wel_Inspe_DEN                              '2009.06.03
        
        
        
                    If Inspe_Proc_DEN(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
        
                Case Wel_Inspe_E_BAG     '2010.01.21
                    
                    If Inspe_Proc_E_BAG(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
                Case Wel_KYOSEI_END
    
                    If KYOSEI_END_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
                Case Wel_LABEL_REPRINT  '2010.01.21
    
                    If LABEL_RePrint_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                
                Case Wel_Inspe_LOGISTIC '2010.01.25
                    
                    If Inspe_Proc_LOGISTIC(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
                Case Wel_SEK_PACKING    '2011.04.25
                    
                    If SEK_PACKING_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
    
                Case Wel_SEK_Inspe      '2011.05.09
            
                    If Inspe_Proc_SEK(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_SEK_SYUGO_PACKING  '2011.05.12
                    
                    If SEK_SYUGO_PACKING_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_SEK_KYOSEI_END     '2011.06.28
            
                    If SEK_KYOSEI_END_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
            
'------------------------------------------------------------------ �S�����i�Ή�    2012.03.21
                Case Wel_Inspe_DEN_ALL                  '�uWEL ��㌟�i�v�̗v��
                    If Inspe_Proc_DEN_ALL(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_Inspe_LOGISTIC_ALL             '�uWEL ۼ޽ï���v�̗v��
                    If Inspe_Proc_LOGISTIC_ALL(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_SEK_PACKING_ALL                '�uWEL �ϐ��@�ʍ�����v�̗v��
                    If SEK_PACKING_ALL_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
'------------------------------------------------------------------ �S�����i�Ή�    2012.03.21
            
            
                Case Wel_Inspe_E_BAG_ALL                                            '2012.06.20
                    If Inspe_Proc_E_BAG_ALL(Sendbuf, i, j) Then                     '2012.06.20
                        Exit Function                                               '2012.06.20
                    End If                                                          '2012.06.20
            
            
            End Select
    
        Case ACT_SYUKA_HYO_OSAKA        '�o�ɕ\�o�Ɂi���o�b�����j   2007.03.16
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_SYUKO_HYO_OSAKA
        
                    If SYUKO_HYO_Dec_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
    
            End Select
    
    
    
    
        Case ACT_IN_KENPIN_OSAKA        '���Ɍ��i�i���o�b�����j   2007.06.07
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_IN_KENPIN_OSAKA
        
                    If NYUKO_KENPIN_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
    
                Case WEL_IN_TANA_S_OSAKA            '���ތ������� 2012.03.01
    
                    If NYUKO_KENPIN_OSAKA_S_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
                Case WEL_MAEGARI_TANA_S_OSAKA       '���ޑO�ؓ��� 2016.05.30
    
                    If NYUKO_MAEGARI_OSAKA_S_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
    
            End Select
    
        Case ACT_IN_TANA_OSAKA        '���Ɍ��i�i���o�b�����j   2007.06.07
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_IN_TANA_OSAKA
        
                    If NYUKO_TANA_OSAKA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
    
            End Select
    
    
        Case ACT_FURIKAE                '���ސU�֏���   2007.06.28
    
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_FURIKAE
        
                    If Furikae_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
        
        
                Case Wel_HIN_FURIKAE_PLUS   '�i�ԐU�֏���   2011.06.01
        
                    If Hin_Furikae_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
        
        
                Case Wel_HIN_FURIKA_S    '���ރZ���^�[�U�֏o��
                    
                    If Hin_Furikae_S_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
    
    
            End Select
    
    
        Case ACT_BINNO                  '�և�����(�ڊǗp) 2009.03.11
    
            If LSBU_IKAN_BinNo_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
    
    
        Case ACT_KENPIN_GAI             '�C�O�������i   2009.08.07
        
            
            
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE   '2014.03.05
                Case WEL_KENPIN_GAI                                                                             '2014.03.05
                    If Inspe_Proc_GAI(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case WEL_KENPIN_Su                          '���ʌ��i   2014.03.05
                    If Inspe_Proc_Su(Sendbuf, i, j) Then    '           2014.03.05
                        Exit Function                       '           2014.03.05
                    End If                                  '           2014.03.05
    
            End Select                                                                                          '2014.03.05
    
'        Case ACT_SAI_SU                 '�ː��^����   2010.01.21
'
'            If SAI_SU_Proc(Sendbuf, i, j) Then
'                Exit Function
'            End If
    
    
        
        Case ACT_SHOUHINKA              '���i������     2010.09.03
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                
                Case Wel_SHOUHINKA_CHECK
                
                    If SHOUHINKA_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_SHOUHINKA_GAI_CHECK                           '2015.11.07
                                                                        '2015.11.07
                    If SHOUHINKA_CHECK_GAI_PROC(Sendbuf, i, j) Then     '2015.11.07
                        Exit Function                                   '2015.11.07
                    End If                                              '2015.11.07
            
            
            
            
                Case Wel_HINBAN_CHECK    '�i������   2010.09.10
                
                    If HINBAN_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_COMPO_CHECK    '�\������   2011.03.02
                
                    If COMPO_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_COMPO_OSAKA_CHECK  '���o�b�@���޾����\������   2012.03.16
                
                    If COMPO_OSAKA_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_KENPIN_OSAKA       '���o�b�@���i   2012.03.16
                
                    If KENPIN_OSAKA_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
                Case Wel_KENPIN_OSAKA_NEW   '���o�b�@���i   2016.05.20
                
                    If KENPIN_OSAKA_NEW_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_KENPIN_OSAKA_NEW2  '���o�b�@���i(�G���[�\���L��)   2016.06.27
                
                    If KENPIN_OSAKA_NEW2_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            

            
            End Select


'-----------------------------------------------    ���g�[�@������  2013.06.06
        Case ACT_LotNo
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                
                Case Wel_LotNo_IN_CHECK
                
                    If LOTNO_IN_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_LotNo_OUT_CHECK
                
                    If LOTNO_OUT_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_LotNo_OUT_CANCEL
                
                    If LOTNO_OUT_CANCEL_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If

                Case Wel_LotNo_LABEL_PRINT
                
                    If LOTNO_LABEL_PRINT_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
            
                Case Wel_InvNo_OUT_CHECK                                '2014.07.01
                                                                        '2014.07.01
                    If INVNO_OUT_CHECK_PROC(Sendbuf, i, j) Then         '2014.07.01
                        Exit Function                                   '2014.07.01
                    End If                                              '2014.07.01
                                    
                Case Wel_InvNo_OUT_CANCEL                               '2014.07.01
                                                                        '2014.07.01
                    If INVNO_OUT_CANCEL_PROC(Sendbuf, i, j) Then        '2014.07.01
                        Exit Function                                   '2014.07.01
                    End If                                              '2014.07.01
            
            
            End Select
'-----------------------------------------------    ���g�[�@������  2013.06.06


'-----------------------------------------------    ���W���[�� 2014.06.24
        Case ACT_MODULE
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                
                Case Wel_MODULE_INSPE
                
                    If MODULE_INSPE_CHECK_PROC(Sendbuf, i, j) Then
                        Exit Function
                    End If
            
            
                Case Wel_MODULE_INSPE2                                  '2015.02.19
                
                    If MODULE_INSPE_CHECK2_PROC(Sendbuf, i, j) Then     '2015.02.19
                        Exit Function                                   '2015.02.19
                    End If                                              '2015.02.19
            
            End Select
'-----------------------------------------------    ���W���[�� 2014.06.24



'-----------------------------------------------    (�V)�������i    2016.10.14
        Case ACT_KENPIN_Drct
        
            If Inspe_Proc_Drct(Sendbuf, i, j) Then
                Exit Function
            End If
'-----------------------------------------------    (�V)�������i    2016.10.14

'-----------------------------------------------   �ް���ވ�  2017.04.10
        Case ACT_BCR_PRINT

            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
                Case Wel_BCR_DAKUTO
                    If BCR_DAKUTO_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_BCR_JAN
                    If BCR_JAN_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_BCR_Inspe
                    If BCR_Inspe_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
                Case Wel_BCR_TANA
                    If BCR_TANA_Proc(Sendbuf, i, j) Then
                        Exit Function
                    End If
            End Select
            

            






'-----------------------------------------------   �ް���ވ�  2017.04.10






    End Select

    Sagyo_Main_Proc = False

End Function

'[2016/05/14 -  mdlProc����ړ�

Public Function Cancel_Proc(Sendbuf As String, Optional Mode As Integer = 0, Optional Para As String = "  ") As Integer
'-------------------------------------------------------
'
'   �w�L�����Z�������i�O��ʌ����j�x
'
'-------------------------------------------------------
    
    Cancel_Proc = True
        
    
    Select Case ID_KANRI_TBL(ING_No).Step
    
        Case Step_Start         '�q�@�d���n�m
        Case Step_TANTO_REQ     '�S���җv��
            
            Call Re_Send_Proc(Sendbuf)
                        
        Case Step_JGYOBU_REQ    '���ƕ��v��

'            ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'            ID_KANRI_TBL(ING_No).JGYOBU = ""
'            Call Start_Proc(Sendbuf)

            '���ƕ��v���Ń��[�v����
            ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
            ID_KANRI_TBL(ING_No).JGYOBU = ""
            
            
            If UBound(NAIGAI) = 0 Then
            Else
                ID_KANRI_TBL(ING_No).NAIGAI = ""
            End If
            
'            ID_KANRI_TBL(ING_No).MENU_GRP = ""
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            
            
            '2010.04.23
            ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
            '2010.04.23
            
            

            Call Menu_Send_Proc(Sendbuf)


        Case Step_NAIGAI_REQ    '�����O�v��
            
            
            If UBound(JGYOBU_T) = 0 Then
            
                        
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
                Call Menu_Send_Proc(Sendbuf)
            
            Else
            
                ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                ID_KANRI_TBL(ING_No).JGYOBU = ""
                ID_KANRI_TBL(ING_No).NAIGAI = ""
            
'               ID_KANRI_TBL(ING_No).MENU_GRP = ""
                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'               ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            
                '2010.04.23
                ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
                '2010.04.23
            
            
            
                Call Menu_Send_Proc(Sendbuf)
            End If

        Case Step_MENU1_REQ     '���j���[�P�v��
        
            
            If Mode = 0 Then            '2008.08.08
'                ST_LOG_OUT_F = False    '2008.08.08
            End If                      '2008.08.08
            
            If UBound(NAIGAI) = 0 Then
                
                
                
                If UBound(JGYOBU_T) = 0 Then
                
                
                
                '�����O�̐؂蕪���Ȃ�
    '                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
    '                ID_KANRI_TBL(ING_No).JGYOBU = ""
    '                Call Start_Proc(Sendbuf)
                
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    
                    '2010.04.23
                    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                    '2010.04.23
            
                
                    Call Menu_Send_Proc(Sendbuf)
            
            
                Else
            
                    '���ƕ��v���Ń��[�v����
                    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
                    ID_KANRI_TBL(ING_No).JGYOBU = ""
        '            ID_KANRI_TBL(ING_No).NAIGAI = ""
                    
        '            ID_KANRI_TBL(ING_No).MENU_GRP = ""
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
        '            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                    
                    
                    '2010.04.23
                    ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                    ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
                    '2010.04.23
                    
                    
        
                    Call Menu_Send_Proc(Sendbuf)
                End If
            
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
                Call Menu_Send_Proc(Sendbuf)
            End If
        
        
        
        Case Step_MENU2_REQ     '���j���[�Q�v��
        
            
            If Not CANCEL_OPE Then      '2008.09.01
                
                If Para <> "EN" Then
                
                
                
                    '�O�񂪃G���[���M
                    Call Re_Send_Proc(Sendbuf)
                    
                    
                    
                    Cancel_Proc = False
                    Exit Function
            
                End If
            End If
            
            
            
            
            
            ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        
            If Mode = 0 Then            '2008.08.08
'                ST_LOG_OUT_F = False    '2008.08.08
            End If                      '2008.08.08
        
        
            Call Menu_Send_Proc(Sendbuf)
        
        
'2006.01.30        Case Step_MENU3_REQ     '���j���[�R�v��
'2006.01.30
'2006.01.30            ID_KANRI_TBL(ING_No).Step = Step_MENU2_REQ
'2006.01.30            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30
'2006.01.30            Call Menu_Send_Proc(Sendbuf)

        Case Step_Sagyo1_REQ    '��ƂP�v��
'2006.01.30            If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV3)) <> 0 Then
'2006.01.30                ID_KANRI_TBL(ING_No).Step = Step_MENU3_REQ
'2006.01.30                ID_KANRI_TBL(ING_No).MENU_LV3 = ""
'2006.01.30                Call Menu_Send_Proc(Sendbuf)
'2006.01.30            Else
                
                
            If Mode = 0 Then            '2008.08.08
'                ST_LOG_OUT_F = False    '2008.08.08
            End If                      '2008.08.08
                
                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) <> 0 Then
                    MENU_UP_F = True   '2008.08.08
                    ID_KANRI_TBL(ING_No).Step = Step_MENU2_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                    Call Menu_Send_Proc(Sendbuf)
                Else
                    MENU_UP_F = True   '2008.08.08
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    Call Menu_Send_Proc(Sendbuf)
                End If
'2006.01.30            End If
                                                    
                                                    
                                                    
                                                    '��ƂQ�^��ƂR�^��ƂS�^��ƂT�v��
        Case Step_Sagyo2_REQ, Step_Sagyo3_REQ, Step_Sagyo4_REQ, Step_Sagyo5_REQ, Step_PRINT_REQ
        
        
        
            Select Case ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE
        
                Case Wel_SHOUHINKA_CHECK
        
        
                    Select Case ID_KANRI_TBL(ING_No).Step
                    
                        Case Step_Sagyo2_REQ
                    
                    
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                            If Sagyo_Send_Proc() Then
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                            
                            Sendbuf = Text_Create_Proc()
                    
                    
                        Case Step_Sagyo3_REQ
                    
                    
                    
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
                            
                            '>>>>>>>>>  2017.09.22
                            'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                            'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                            
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
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
                            
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                            
                            
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & _
                                                                                    Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                                                                                    '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    '���͌���
                            Send_Text.Box_Type(3).Max_Size = "13"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�S�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            
                            If ID_KANRI_TBL(ING_No).GENPIN_CNT = 0 Then
                            
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & _
                                                                                        Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
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
        
                            Sendbuf = Text_Create_Proc()
                    
                    
                                            
                    
                    
                        Case Step_Sagyo4_REQ
                    
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))) & Format(ID_KANRI_TBL(ING_No).SHIJI_QTY, "#0"))
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
    
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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                            If ID_KANRI_TBL(ING_No).GENPIN_CNT = 0 Then
                            
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 13) & Space(7 - Len(Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))) & _
                                                                                        Format(ID_KANRI_TBL(ING_No).GENPIN_CNT, "#0"))
                            End If
                                                                                    '���l�����\��
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(4).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '���͌���
                            Send_Text.Box_Type(4).Max_Size = "13"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "13"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
                            Sendbuf = Text_Create_Proc()
                    
                    
                    End Select
        
        
        
        
                Case Wel_SHOUHINKA_GAI_CHECK        '2016.04.05
        
                    Select Case ID_KANRI_TBL(ING_No).Step
                    
                        Case Step_Sagyo1_REQ
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                            If Sagyo_Send_Proc() Then
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                            
                            Sendbuf = Text_Create_Proc()
                        Case Step_Sagyo2_REQ
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                            If Sagyo_Send_Proc() Then
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            End If
                            
                            Sendbuf = Text_Create_Proc()
                        
                        
                        Case Step_Sagyo3_REQ
                        
                            ID_KANRI_TBL(ING_No).LABEL_CNT = 0
                            ID_KANRI_TBL(ING_No).GENPIN_CNT = 0
                            ID_KANRI_TBL(ING_No).GAISOU_CNT = 0
                        
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
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
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_L_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_L_HIN_CNT)
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
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                                                                                    '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    '���͌���
                            Send_Text.Box_Type(3).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
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
        
                            Exit Function
                        
                        
                        Case Step_Sagyo4_REQ
                    
                    
                    
                    
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�w�}�[��" & ID_KANRI_TBL(ING_No).SHIJI_No)
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
                            
                            
                            
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & _
                                                                                    Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Left(ID_KANRI_TBL(ING_No).Hinban, 16) & _
                                                                                    Space(4 - Len(Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))) & Format(ID_KANRI_TBL(ING_No).LABEL_CNT, "#0"))
                            
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_GAI_HIN_CNT)
                                                                                    '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = "01"                  '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    
                            Send_Text.Box_Type(3).Max_Size = "20"                           '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"      '2011.04.11 13-->20
                            
                            
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_G_HIN_CNT)
                                                                                    '���l�����\��
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(4).Start_Pos = "01"                  '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    
                            Send_Text.Box_Type(4).Max_Size = "20"                           '2011.04.11 13-->20
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"      '2011.04.11 13-->20
                                                                                    
                                                                                    
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
                            Sendbuf = Text_Create_Proc()
    
                            Exit Function
                    
                    
                    End Select
                Case Else
        
            
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
                    If Sagyo_Send_Proc() Then
                        Sendbuf = Text_Create_Proc()
                        Exit Function
                    End If
                    
                    Sendbuf = Text_Create_Proc()
                
                    If Mode = 0 Then            '2008.08.08
        '                ST_LOG_OUT_F = False    '2008.08.08
                    End If                      '2008.08.08
    
            End Select
    
    End Select
    
    Cancel_Proc = False


End Function

'[2016/05/14 -  mdlProc����ړ�
Public Function tmpZaiko_Clear_Proc() As Integer
'-------------------------------------------------------
'
'   �w�݌Ƀf�[�^�i�ꎞ�f�[�^�j�̏����x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim RETRY_CNT   As Integer
    
    
    tmpZaiko_Clear_Proc = True
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        tmpZaiko_Clear_Proc = SYS_ERR
        Exit Function
    End If
    
    com = BtOpGetFirst

    Do
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                                
                Case BtNoErr
                    
                    Exit Do
                
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                        tmpZaiko_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                    tmpZaiko_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        RETRY_CNT = 0
        
        Do
        
            sts = BTRV(BtOpDelete, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                        tmpZaiko_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
'                    DoEvents
                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                        DoEvents                                                    '2016.01.26
                    End If                                                          '2016.01.26
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                    tmpZaiko_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop

End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        tmpZaiko_Clear_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    tmpZaiko_Clear_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If


End Function


