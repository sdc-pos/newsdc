Attribute VB_Name = "mdlSKE"
Option Explicit


Public Sub SEK_HINB_CD_CNT_Proc(HINB_CD_CNT As Long)

'-------------------------------------------------------
'
'   �w�ϐ��n�E�X�o�׍���@�i�ԏ�Ԃ̃J�E���g�x
'               2011.04.25
'
'-------------------------------------------------------
Dim i   As Integer

Dim k   As Integer
Dim l   As Integer

    HINB_CD_CNT = 0
    For i = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_TBL)
    
        If ID_KANRI_TBL(ING_No).SEK_TBL(i).SEK_KONPO_F Then
            HINB_CD_CNT = HINB_CD_CNT + 1
        End If
    
    Next i
    '------------------------------------------------------------   2011.06.27
    For k = 0 To UBound(ID_KANRI_TBL)
        If k = ING_No Then
        Else
            If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = _
                ID_KANRI_TBL(k).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(k).Sagyo_Code.YOIN_CODE Then
                
                
                If ID_KANRI_TBL(ING_No).SEK_TEI_LABELID = ID_KANRI_TBL(k).SEK_TEI_LABELID Then
                    For l = 0 To UBound(ID_KANRI_TBL(k).SEK_TBL)
                        
                        
                        If l <= UBound(ID_KANRI_TBL(ING_No).SEK_TBL) Then
                            If ID_KANRI_TBL(ING_No).SEK_TBL(l).SEK_KONPO_F Then
                            Else
                                If ID_KANRI_TBL(k).SEK_TBL(l).SEK_KONPO_F Then
                                    HINB_CD_CNT = HINB_CD_CNT + 1
                                End If
                            End If
                        Else
                            If ID_KANRI_TBL(k).SEK_TBL(l).SEK_KONPO_F Then
                                HINB_CD_CNT = HINB_CD_CNT + 1
                            End If
                        End If
                    
                    Next l
                End If
            End If
        End If
    Next k
    '------------------------------------------------------------   2011.06.27



End Sub






Public Sub SEK_Kenpin_Count_Proc(Sumi_CNT As Integer)
'-------------------------------------------------------
'
'   �w�ϐ��n�E�X�����o�ח\�� ���i�ς݂̃J�E���g�x
'   ���o�b�p  2011.05.09
'-------------------------------------------------------
Dim i   As Integer

    Sumi_CNT = 0

    For i = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL)

        If ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_KENPIN_F Then
            Sumi_CNT = Sumi_CNT + 1
        End If

    Next i

End Sub
Public Function SEK_SYUGO_PACKING_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�ϐ��n�E�X�@�@�ʏW��������x
'
'   2011.05.12
'-------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim i               As Integer
Dim j               As Integer


Dim KONPO_ID        As String * 20

Dim ID_NO           As String * 13
Dim Found_Flg       As Integer
Dim SYUGO_SUMI_CNT  As Long

Dim IN_WORD         As String * 50
Dim OUT_WORD        As String * 20


Dim SAI_SU          As Double

Dim MENU_NO         As String * 2


    SEK_SYUGO_PACKING_Proc = True





    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�w�}�h�c�j
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_SEK_SYUGO_ID_NO      '�w�}�h�c
                        KONPO_ID = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        sts = Y_SYUKA_TEI_SYUGO_Check_Proc(KONPO_ID, Found_Flg)
                        Select Case sts
                            Case BtNoErr          '����
                
                                '------------------ ����ς݂̃`�F�b�N
                                If Not Found_Flg Then
                                    
                                    If ID_KANRI_TBL(ING_No).SEK_SYUGO_SAI_SU >= 999999 Then
                                    
                                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            
                                        ID_KANRI_TBL(ING_No).SEK_KONPO_ID = KONPO_ID
                            
                            
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
                                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
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
                                        IN_WORD = Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME)
                                        Call Moji_Cut_Proc(IN_WORD, OUT_WORD, 20)
                                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Trim(OUT_WORD))
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Trim(OUT_WORD))
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
                                                                                                
                                        '�W�������Ԃ̃J�E���g
                                        Call SEK_SYUGO_CNT_Proc(SYUGO_SUMI_CNT)
                                                                                                
                                                                                                
                                                                                                'BOX����
                                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                                '�\�����e
                                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                                    " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                                    " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
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
                                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_KUTI_SU_S & " 1")
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_KUTI_SU_S & " 1")
                                                                                        '���l�����\��
                                        Send_Text.Box_Type(3).INIT = ""
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                        '�����J�[�\���ʒu
                                        Send_Text.Box_Type(3).Start_Pos = ""
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                                '���͌���
                                        Send_Text.Box_Type(3).Max_Size = "00"
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                
                                        '-----------------------------------------------�T�s��
                                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                        '�\�����e
                                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_SAI_SU_S)
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_SAI_SU_S)
                                                                                        '���l�����\��
                                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Trim(Format(0, "#0.00")))) & Trim(Format(0, "#0.00"))
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Trim(Format(0, "#0.00")))) & Trim(Format(0, "#0.00"))
                                                                                        '�����J�[�\���ʒu
                                        Send_Text.Box_Type(4).Start_Pos = "12"
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "12"
                                                                                '���͌���
                                        Send_Text.Box_Type(4).Max_Size = "05"
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                        
                            
                                        Sendbuf = Text_Create_Proc()
                                    
                                        SEK_SYUGO_PACKING_Proc = False
                                        Exit Function
                                    
                                    Else
'2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, KONPO_ID, "�S������ς݂ł�", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        SEK_SYUGO_PACKING_Proc = False
                                        Exit Function
                                    End If
                                End If
                            Case BtErrEOF
'2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, KONPO_ID, "�����ް��Ȃ�", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SEK_SYUGO_PACKING_Proc = False
                                Exit Function
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                
                        ID_KANRI_TBL(ING_No).SEK_KONPO_ID = KONPO_ID
                        
                        
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
'2017.09.22
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
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
                        
                        
                        IN_WORD = Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME)
                        Call Moji_Cut_Proc(IN_WORD, OUT_WORD, 20)
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(OUT_WORD))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(OUT_WORD))
                        
                        
'                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME))
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME))
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
                                                                                
                        '�W�������Ԃ̃J�E���g
                        Call SEK_SYUGO_CNT_Proc(SYUGO_SUMI_CNT)
                                                                                
                                                                                
                                                                                'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                    " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                    " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
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
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_SEK_ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_SEK_ID_NO)
                                                                                '���l�����\��
                        Send_Text.Box_Type(4).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = "01"                  '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                '���͌���
                         Send_Text.Box_Type(4).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�w�����j
        
            For i = 0 To M_Gyo - 1
                Select Case Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), 3)
                    
                    Case LCD_SEK_ID_NO     '�w����
                        ID_NO = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        '�Y���i�ԗL��������
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL)
                            If Trim(ID_NO) = Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(j).SEK_TEI_LABELID) Then
                                Exit For
                            End If
                        Next j
                        
                        If j > UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) Then
'2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).SEK_KONPO_ID, ID_NO, "�w�����G���[", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SEK_SYUGO_PACKING_Proc = False
                            Exit Function
                        End If
                        
                
                        If ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(j).SEK_SYUGO_F Then
'2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).SEK_KONPO_ID, ID_NO, "����ς݁I", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SEK_SYUGO_PACKING_Proc = False
                            Exit Function
                        End If
                
                        '2011.05.27
                        If SEK_KONPO_F Then
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL)
                                If Trim(ID_NO) = Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(j).SEK_TEI_LABELID) Then
                                    If Not ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(j).SEK_KONPO_F Then
                                        Exit For
                                    End If
                                End If
                            Next j

                            If j <= UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) Then
'2011.06.22                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).SEK_KONPO_ID, ID_NO, "������L��I�I", "")
'2017.09.22
                                
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).SEK_KONPO_ID, ID_NO, "������ł��I�I", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SEK_SYUGO_PACKING_Proc = False
                                Exit Function
                            End If
                        End If
                        '2011.05.27
                        
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL)
                            If Trim(ID_NO) = Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(j).SEK_TEI_LABELID) Then
                                Exit For
                            End If
                        Next j
                        
                        ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(j).SEK_SYUGO_F = True
                        ID_KANRI_TBL(ING_No).SEK_KEN_TEI_LABELID = ID_NO
    
                        '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                        '------------------------------------   �����ް��̏���
                        
                        com = BtOpGetGreaterEqual
                        Call UniCode_Conv(K1_Y_SYU_TEI.TEI_LABELID, ID_KANRI_TBL(ING_No).SEK_KEN_TEI_LABELID)
            
                        Do
                            
'                            DoEvents
                            If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                DoEvents                                                    '2016.01.26
                            End If                                                          '2016.01.26
                        
                            Do
                            
                                sts = BTRV(com + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K1_Y_SYU_TEI, Len(K1_Y_SYU_TEI), 1)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrEOF
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�@�ʒ����f�[�^�g�p��", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        SEK_SYUGO_PACKING_Proc = False
                                        GoTo Abort_Tran
                                    Case Else
                                        Call File_Error(sts, com + BtSNoWait, "�@�ʒ����f�[�^", 0)
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                            Loop
                                    
                            If sts = BtErrEOF Then
                                Exit Do
                            End If
            
                            If Trim(StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode)) <> Trim(ID_KANRI_TBL(ING_No).SEK_KEN_TEI_LABELID) Then
                                Exit Do
                            End If
                        
                
                            Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, ID_KANRI_TBL(ING_No).TANTO_CODE)
                            Call UniCode_Conv(Y_SYU_TEI_REC.SYUGO_KONPO_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                
                
                            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                            Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                
                
                
                            Do
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                                sts = BTRV(BtOpUpdate, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K1_Y_SYU_TEI, Len(K1_Y_SYU_TEI), 1)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrKeyNotFound
'2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�@�ʒ����f�[�^�s��", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        SEK_SYUGO_PACKING_Proc = False
                                        GoTo Abort_Tran
                                     Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�@�ʒ����f�[�^�g�p��", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        SEK_SYUGO_PACKING_Proc = False
                                        GoTo Abort_Tran
                                    Case Else
                                        Call File_Error(sts, BtOpUpdate, "�@�ʒ����f�[�^", 0)
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                            Loop
                
                
                            com = BtOpGetNext
                
                
                        Loop
                                            
                        '------------------------------------   �݌Ɉړ������̏���
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                            MENU_NO = ""
                        End If
                                                        
                                                        
                        sts = IDOREKI_OUTPUT_PROC("", _
                                                    "", _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    "", _
                                                    "", _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    0, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    FILE_RETRY, _
                                                    , _
                                                    ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME, _
                                                    , , , , MENU_NO, _
                                                    , _
                                                    "", _
                                                    , , , , 1, , , , , , , ID_KANRI_TBL(ING_No).SEK_KEN_TEI_LABELID)
                        Select Case sts
                            Case False      '����I��
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                SEK_SYUGO_PACKING_Proc = SYS_ERR
                                GoTo Abort_Tran
                        End Select
            
                        '------------------------------------   ��ƃ��O�̏���
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
                                                                 , , , , , _
                                                                 ID_KANRI_TBL(ING_No).SEK_KEN_TEI_LABELID) Then
                                SEK_SYUGO_PACKING_Proc = SYS_ERR
                                Exit Function
                            End If
                        End If
                                            '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                    
                        '�W�������Ԃ̃J�E���g
                        Call SEK_SYUGO_CNT_Proc(SYUGO_SUMI_CNT)
        
                        If SYUGO_SUMI_CNT <> UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1 Then
                            
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
'2017.09.22
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
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
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
                            
                            
                            IN_WORD = Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME)
                            Call Moji_Cut_Proc(IN_WORD, OUT_WORD, 20)
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(OUT_WORD))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(OUT_WORD))
                            
                            
                            
'                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME))
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME))
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
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                        " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                        " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
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
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_SEK_ID_NO)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_SEK_ID_NO)
                                                                                    '���l�����\��
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(4).Start_Pos = "01"                    '���l�͂T���Œ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                                                    '���͌���
                             Send_Text.Box_Type(4).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "20"
                                                                                    
                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    
                            Sendbuf = Text_Create_Proc()
            
                        Else
                        '�����
            
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
'                                                                                    'BOX����
'                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
'                                                                                    '�\�����e
'                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                            '-----------------------------------------------�P�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(0).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
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
                            
                            
                            IN_WORD = Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME)
                            Call Moji_Cut_Proc(IN_WORD, OUT_WORD, 20)
                            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Trim(OUT_WORD))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Trim(OUT_WORD))
                            
                            
 '                           Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME))
 '                           Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME))
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
                                                                                    
                            '�W�������Ԃ̃J�E���g
                            Call SEK_SYUGO_CNT_Proc(SYUGO_SUMI_CNT)
                                                                                    
                                                                                    
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                        " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 1, 4) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 5, 2) & "/" & Mid(ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD, 7, 2) & _
                                                                        " (" & Format(SYUGO_SUMI_CNT, "#0") & "/" & Format(UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) + 1, "#0") & ")")
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
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_KUTI_SU_S & " 1")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_KUTI_SU_S & " 1")
                                                                            '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                            '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                    '���͌���
                            Send_Text.Box_Type(3).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                    
                            '-----------------------------------------------�T�s��
                                    
                            If ID_KANRI_TBL(ING_No).SEK_SYUGO_SAI_SU >= 999999 Then
                            
                            
                            
                                Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                '�\�����e
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_SAI_SU_S)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_SAI_SU_S)
                                                                                '���l�����\��
                                Send_Text.Box_Type(4).INIT = Space(10 - Len(Trim(Format(0, "#0.00")))) & Trim(Format(0, "#0.00"))
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Trim(Format(0, "#0.00")))) & Trim(Format(0, "#0.00"))
                                                                                '�����J�[�\���ʒu
                                Send_Text.Box_Type(4).Start_Pos = "12"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "12"
                                                                        '���͌���
                                Send_Text.Box_Type(4).Max_Size = "05"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                            
                            
                            Else
                            
                            
                            
                                Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                                                '�\�����e
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_KUTI_SU_S & Format(ID_KANRI_TBL(ING_No).SEK_SAI_SU, "#0.00"))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_KUTI_SU_S & Format(ID_KANRI_TBL(ING_No).SEK_SAI_SU, "#0.00"))
                                                                                '���l�����\��
                                Send_Text.Box_Type(3).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                '�����J�[�\���ʒu
                                Send_Text.Box_Type(3).Start_Pos = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
                                                                        '���͌���
                                Send_Text.Box_Type(3).Max_Size = "00"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                            
                            End If
                            
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
'                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
'                                                                                    '�\�����e
'                            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "��ENT")
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "��ENT")
'                                                                                    '���l�����\��
'                            Send_Text.Box_Type(4).INIT = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
'                                                                                    '�����J�[�\���ʒu
'                            Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
'                                                                                    '���͌���
'                             Send_Text.Box_Type(4).Max_Size = "00"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
'
'                            Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                
                            Sendbuf = Text_Create_Proc()
                        End If
                End Select
            Next i
                    
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�iAny Key�j
            
            
            
            If ID_KANRI_TBL(ING_No).SEK_SYUGO_SAI_SU >= 999999 Then
            
            
                For i = 0 To M_Gyo - 1
            
                
                
                    Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                        '�ː�
                        Case LCD_SAI_SU_S
                    
                            If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SEK_SYUGO_PACKING_Proc = False
                                Exit Function
                            
                            End If
                    
                            SAI_SU = Val(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                    
                    
                            If SAI_SU = 0 Then
'2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SEK_SYUGO_PACKING_Proc = False
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
                            '------------------------------------   �����ް��̏���
                            Call UniCode_Conv(K3_Y_SYU_TEI.KONPO_ID, ID_KANRI_TBL(ING_No).SEK_KONPO_ID)
                            com = BtOpGetGreaterEqual
            
                            Do
                        
'                                DoEvents
                                If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                    DoEvents                                                    '2016.01.26
                                End If                                                          '2016.01.26
                    
                                Do
                        
                                    sts = BTRV(com + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K3_Y_SYU_TEI, Len(K3_Y_SYU_TEI), 3)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrEOF
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�@�ʒ����f�[�^�g�p��", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            SEK_SYUGO_PACKING_Proc = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, com + BtSNoWait, "�@�ʒ����f�[�^", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                        
                            
                                If sts = BtErrEOF Then
                                    Exit Do
                                End If
                    
                                If Trim(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode)) <> Trim(ID_KANRI_TBL(ING_No).SEK_KONPO_ID) Then
                                    Exit Do
                                End If
                        
                
                                Call UniCode_Conv(Y_SYU_TEI_REC.KUTI_SU, "0001")
                                Call UniCode_Conv(Y_SYU_TEI_REC.SAI_SU, Format(SAI_SU, "000.00"))
                        
                                Call UniCode_Conv(Y_SYU_TEI_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                Call UniCode_Conv(Y_SYU_TEI_REC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                
                                
                                Do
'                                    DoEvents
                                    If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
                                        DoEvents                                                    '2016.01.26
                                    End If                                                          '2016.01.26
                                    sts = BTRV(BtOpUpdate, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K3_Y_SYU_TEI, Len(K3_Y_SYU_TEI), 3)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
'2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�@�ʒ����f�[�^�s��", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            SEK_SYUGO_PACKING_Proc = False
                                            GoTo Abort_Tran
                                         Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�@�ʒ����f�[�^�g�p��", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            SEK_SYUGO_PACKING_Proc = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�@�ʒ����f�[�^", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                
                
                                com = BtOpGetNext
                            
                    
                            Loop
                                    '�g�����U�N�V�����I��
                            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            If sts <> BtNoErr Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                            End If
            
                    End Select
                Next i
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

    SEK_SYUGO_PACKING_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

    SEK_SYUGO_PACKING_Proc = False

End Function


Public Function Y_SYUKA_TEI_SYUGO_Check_Proc(ID_NO As String, Found_Flg As Integer) As Integer
'-------------------------------------------------------
'
'   �w�ϐ��n�E�X�o�׏W������@�����f�[�^�̃`�F�b�N�x
'               2011.05.12
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer



    Y_SYUKA_TEI_SYUGO_Check_Proc = True
    
    
    Call UniCode_Conv(K3_Y_SYU_TEI.KONPO_ID, ID_NO)
    com = BtOpGetGreaterEqual
    
    
    Found_Flg = False
    
    i = -1
    Erase ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL
    
    
    Do
    
'        DoEvents
        If StrComp(App.EXEName, "F110010", vbTextCompare) <> 0 Then     '2016.01.26
            DoEvents                                                    '2016.01.26
        End If                                                          '2016.01.26
    
        sts = BTRV(com, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K3_Y_SYU_TEI, Len(K3_Y_SYU_TEI), 3)
        Y_SYUKA_TEI_SYUGO_Check_Proc = sts
        Select Case sts
            Case BtNoErr
                If Trim(ID_NO) <> Trim(StrConv(Y_SYU_TEI_REC.KONPO_ID, vbUnicode)) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Y_SYUKA_TEI_SYUGO_Check_Proc = SYS_ERR
                Exit Function
        End Select
    
    
        If Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode)) = "" Then
            Found_Flg = True
        End If
        
        If i = -1 Then
            
            
            ID_KANRI_TBL(ING_No).SEK_SYUGO_SYU_YMD = StrConv(Y_SYU_TEI_REC.SYU_YMD, vbUnicode)
            ID_KANRI_TBL(ING_No).SEK_SYUGO_L_TOK_NAME = StrConv(Y_SYU_TEI_REC.L_TOK_NAME, vbUnicode)
            ID_KANRI_TBL(ING_No).SEK_SYUGO_SAI_SU = StrConv(Y_SYU_TEI_REC.SAI_SU, vbUnicode)
            
            i = i + 1
                
            ReDim Preserve ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(0 To i)
            ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_TEI_LABELID = StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode)
            
            
            If Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode)) = "" Then
                ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_SYUGO_F = False
                Found_Flg = True
            Else
                ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_SYUGO_F = True
            End If
            
            
            If Trim(StrConv(Y_SYU_TEI_REC.KONPO_TANTO, vbUnicode)) = "" Then
                ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_KONPO_F = False
                Found_Flg = True
            Else
                ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_KONPO_F = True
            End If
            
            
            
            
        Else
            
            For i = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL)
                
                If Trim(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_TEI_LABELID) = Trim(StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode)) Then
                    Exit For
                End If
            
            Next i
            
            
            If i > UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL) Then
            
                ReDim Preserve ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(0 To i)
                ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_TEI_LABELID = StrConv(Y_SYU_TEI_REC.TEI_LABELID, vbUnicode)
                
                
                If Trim(StrConv(Y_SYU_TEI_REC.SYUGO_KONPO_TANTO, vbUnicode)) = "" Then
                    ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_SYUGO_F = False
                    Found_Flg = True
                Else
                    ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_SYUGO_F = True
                End If
                
                
                If Trim(StrConv(Y_SYU_TEI_REC.KONPO_TANTO, vbUnicode)) = "" Then
                    ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_KONPO_F = False
                    Found_Flg = True
                Else
                    ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_KONPO_F = True
                End If
            End If
        
        End If
            
            
        
        
            
        
        com = BtOpGetNext
    
    Loop
    
    Y_SYUKA_TEI_SYUGO_Check_Proc = False
    If i = -1 Then
        Y_SYUKA_TEI_SYUGO_Check_Proc = BtErrEOF
    End If

End Function

Public Sub SEK_SYUGO_CNT_Proc(SEK_SYUGO_CNT As Long)

'-------------------------------------------------------
'
'   �w�ϐ��n�E�X�o�׏W������@�����Ԃ̃J�E���g�x
'               2011.05.12
'
'-------------------------------------------------------
Dim i   As Integer


    SEK_SYUGO_CNT = 0
    For i = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL)
    
        If ID_KANRI_TBL(ING_No).SEK_SYUGO_TBL(i).SEK_SYUGO_F Then
            SEK_SYUGO_CNT = SEK_SYUGO_CNT + 1
        End If
    
    Next i



End Sub

Public Function SEK_Label_File_Make_Proc(FileName As String, Start_Page_No As Integer, Mai_Su As Integer, TOTAL_Mai_Su As Integer) As Integer
'-------------------------------------------------------
'
'
'   �w�ϐ������@�׎D�p�f�[�^�t�@�C���쐬�x
'       2011.05.09
'-------------------------------------------------------
Dim sts         As Integer


Dim FileNo      As Long

Dim FullPath    As String


Dim LCnt        As Integer

Dim OKURI_NO    As Integer


Dim Work        As String * 20


    
Dim c           As String * 128
Dim DEF_FLG     As Boolean

    
    
''''''''''''''''''  �@�ʌ��i���A�P���ڂ���׎D���s      2011.05.25
'    If ID_KANRI_TBL(ING_No).SEK_PAGE_NO < SEK_LABEL_PAGE Then
'        ID_KANRI_TBL(ING_No).SEK_PAGE_NO = ID_KANRI_TBL(ING_No).SEK_PAGE_NO + 1
'        SEK_Label_File_Make_Proc = False
'        Exit Function
'    End If
''''''''''''''''''  �@�ʌ��i���A�P���ڂ���׎D���s      2011.05.25
    
'sts = MsgBox("�׎D�쐬", vbYesNo, "�e�X�g�p")
'If sts = vbNo Then
'    SEK_Label_File_Make_Proc = False
'    Exit Function
'End If
    
''''''''''''''''''''''''''''''' 2011.06.29
    If Start_Page_No < SEK_LABEL_PAGE Then
        Mai_Su = Mai_Su - Start_Page_No
    End If
''''''''''''''''''''''''''''''' 2011.06.29
    
    
    
    SEK_Label_File_Make_Proc = True
        
    Start_Page_No = Start_Page_No + 1

    If Right(F1100101.CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = F1100101.CtrsWsk1.SendFolder & "\" & F0_SendFile & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"
    Else
        FullPath = F1100101.CtrsWsk1.SendFolder & F0_SendFile & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"
    End If


    If GetIni("LABEL", "LABEL_DEF", "F110010LABEL", c) Then
        DEF_FLG = False
    Else
        DEF_FLG = True
    End If


    On Error Resume Next
    Kill (FullPath)             '���M�p�t�@�C���폜
    On Error GoTo 0
        
    FileNo = FreeFile           '���M�p�t�@�C���n�o�d�m
    Open FullPath For Output As #FileNo

    Print #FileNo, "#"
    Print #FileNo, "JOB"
    
    
    If DEF_FLG Then
        Print #FileNo, Trim(c)
    Else
        Print #FileNo, "DEF MK=1,DK=8,MD=3,PW=384,PH=344,XO=8,UM=24,BM=0,AF=1"
    End If
    
    Print #FileNo, "START"
    OKURI_NO = Len(Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)) + 1
    '�����(BC)
    If Trim(Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)) = Trim(KEN_CHARTER_CD) Or _
        Trim(Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)) = Trim(KEN_AKABOU_CD) Or _
        Trim(Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)) = Trim(KEN_LOGISTIC_CD) Then
        Print #FileNo, ""
    Else
        
        
        Print #FileNo, "BCD TP=6,X=0,Y=0,RA=1,HT=80,HR=1,MG=0,NS=" & Format(OKURI_NO, "#0") & ",NE=2,NZ=0"
        Print #FileNo, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) & Format(Start_Page_No, "00")
    End If
    
    Print #FileNo, "#FONT TP=3,CS=0"
    
    '�����
    If Trim(Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)) = Trim(KEN_CHARTER_CD) Or _
        Trim(Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)) = Trim(KEN_AKABOU_CD) Or _
        Trim(Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)) = Trim(KEN_LOGISTIC_CD) Then
        Print #FileNo, ""
    Else
        Print #FileNo, "#TEXT X=33,Y=60,L=1,NS=12,NS=" & Format(OKURI_NO, "#0") & ",NE=2,NZ=0"
        Print #FileNo, "#" & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) & Format(Start_Page_No, "00")
    End If
    
    
    '�������
    Print #FileNo, "FONT TP=7,CS=0,LG=72,WD=36,LS=0"
    Print #FileNo, "TEXT X=214,Y=245,L=1,NS=1,NE=3,NZ=1"
    Print #FileNo, Format(Start_Page_No, "000") & "/" & Format(TOTAL_Mai_Su, "#")
    
    Print #FileNo, "FONT TP=7,CS=0,LG=36,WD=18,LS=0"
    Print #FileNo, "TEXT X=0,Y=65,L=7"
    
    Print #FileNo, ""
    
    
    
    
    
    
    '�Z��1
    Print #FileNo, Trim(Mid(ID_KANRI_TBL(ING_No).KEN_JYUSHO, 1, 20))
    
    '�Z��2
    Print #FileNo, Trim(Mid(ID_KANRI_TBL(ING_No).KEN_JYUSHO, 21, 20))
    
    '�����
    Print #FileNo, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
    
    Print #FileNo, ""
    
    '���t
    Print #FileNo, Mid(ID_KANRI_TBL(ING_No).KEN_SYUKA_YMD, 1, 4) & _
                                               "�N" & Mid(ID_KANRI_TBL(ING_No).KEN_SYUKA_YMD, 5, 2) & _
                                               "��" & Mid(ID_KANRI_TBL(ING_No).KEN_SYUKA_YMD, 7, 2) & "��"
    


    
    '���茳
    Print #FileNo, "�r�c�b���@���ޗ��ʂb"
    
    
    
    Print #FileNo, "QTY P=" & Format(Mai_Su, "#0")
    Print #FileNo, "END"


    
    Print #FileNo, "JOBE"
    
    FileName = F0_SendFile & Format(ID_KANRI_TBL(ING_No).ID, "000") & ".txt"










    Close #FileNo



    SEK_Label_File_Make_Proc = False


End Function


Public Sub SEK_KUTI_SU_Count_Proc(SAI_SU As Double, KUTI_SU As Integer)
'-------------------------------------------------------
'
'   �w�ϐ��n�E�X�����o�ח\�� �����^�ː��̃J�E���g�x
'   ���o�b�p  2011.05.09
'-------------------------------------------------------
Dim i   As Integer

    KUTI_SU = 0
    SAI_SU = 0

    For i = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL)



'        KUTI_SU = KUTI_SU + ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_KUTI_SU
'        If ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_SAI_SU > 999 Or ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_SAI_SU = 0 Then
        If ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_SAI_SU >= 999999 Then
            SAI_SU = ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_SAI_SU
            Exit For
        Else
            SAI_SU = SAI_SU + ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_SAI_SU
        End If


    Next i


    For i = 0 To UBound(ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL)

        KUTI_SU = KUTI_SU + ID_KANRI_TBL(ING_No).SEK_KENPIN_TBL(i).SEK_KUTI_SU


Debug.Print KUTI_SU

    Next i


End Sub


