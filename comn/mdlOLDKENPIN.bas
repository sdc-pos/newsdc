Attribute VB_Name = "mdlOLDKENPIN"
Option Explicit
Public Function Inspe_Proc_DEN(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����i��z�`�[�ǂݍ��� ���o�b�����j�x
'
'       2006.12.07
'
'
'-------------------------------------------------------
Dim sts             As Integer

'2010.12.07
'Dim Hinban          As String * 13
Dim Hinban          As String * 20
'2010.12.07


Dim SYUKA_QTY       As Long
Dim MTS_CODE        As String * 8

'2010.12.07
'Dim HIN_NO          As String * 13
Dim HIN_NO          As String * 20
'2010.12.07

Dim OKURI_NO        As String
Dim KUTI_SU         As Long

Dim SAI_SU          As Double

Dim UNSOU_KAISHA    As String
 
Dim SYUKA_YMD       As String
Dim JYUSHO          As String
Dim BIKOU           As String

Dim SURYO           As String

Dim Y_SYU_TBL()     As KEN_DEN_TBL_Tag

Dim KAN_FLG         As String * 1

Dim i               As Integer
Dim j               As Integer
Dim k               As Integer

Dim DEN_ID_LOOP     As Integer
Dim DEN_ID_JGYOBU   As String * 1

Dim Y_SYU_CNT       As Integer
Dim Sumi_CNT        As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 7
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

Dim KENPIN_END      As Boolean

Dim OKURI_SAKI      As String

Dim CANCEL_F        As Boolean

Dim FAST_F          As Boolean
Dim Found_F         As Boolean

'2010.01.21
Dim KONPOU_ON           As Integer




Dim KONPOU_ON_SUMI      As Integer
Dim KONPOU_OFF          As Integer
Dim KONPOU_OFF_SUMI     As Integer



Dim PRINT_OFF           As Boolean
Dim Start_Page_No       As Long
Dim PRINT_TOTAL_SU      As Long
Dim PRINT_MAISU         As Long
Dim FileName            As String
Dim ID_SEQ              As Integer
Dim DISP_SAI_SU         As Double

Dim wkKUTI_SU           As String
Dim wkKONPO_F           As String * 1

Dim TOTAL_KUTI_SU       As Integer
Dim TOTAL_SAI_SU        As Double
Dim MUKE_NAME           As String
Dim OKURI_NO_MAX        As Integer
Dim KUTI_SU_INPUT_F     As Boolean

Dim KEN_TEL_NO          As String * 20

Dim KEN_TYAKUTEN        As String * 3       '2017.04.06

Dim OKURI_NO_F          As Boolean
'2010.01.21


Dim FUKUYAMA_CHK_F      As Boolean

    Inspe_Proc_DEN = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�`�[�h�c�j
        
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No  '�`�[�h�c
                                
                        '�e�`���j�d�d�o
                        ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                         
                        Erase Y_SYU_TBL
                                        
                        sts = Y_Syuka_H_Chek_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                                MTS_CODE, _
                                                Y_SYU_CNT, _
                                                Sumi_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                Y_SYU_TBL(), _
                                                OKURI_NO, _
                                                UNSOU_KAISHA, _
                                                OKURI_SAKI, _
                                                Found_F, _
                                                SYUKA_YMD, _
                                                JYUSHO, _
                                                BIKOU, _
                                                Start_Page_No, _
                                                KUTI_SU, _
                                                MUKE_NAME, _
                                                OKURI_NO_MAX, , , _
                                                KEN_TEL_NO, , , _
                                                KEN_TYAKUTEN)
                        
                        
                        '���[���Ŏg�p�� 2011.04.07
                        If sts = SYS_CANCEL Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\��g�p��", "", "")   '2107.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        End If
                        '���[���Ŏg�p�� 2011.04.07
                                
                        
                        
                        
                        '��ݾٓ`�[������
                        If Found_F Then
                        
                            CANCEL_F = True
                                                     
                            
                            For j = 0 To UBound(Y_SYU_TBL)
                            
                                If Not Y_SYU_TBL(j).CANCEL_F Then
                                    CANCEL_F = False
                                    Exit For
                                End If
                                                    
                            Next j
                                                     
                                                     
                            If CANCEL_F Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�L�����Z���`�[�ł��B", "", "")         '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�L�����Z���`�[�ł��B", "", "")     '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            End If
                        End If
                        
                        
                        If Y_SYU_CNT = 0 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\�薳��", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        End If
                                                 
                        If Sumi_CNT = Y_SYU_CNT And Start_Page_No <> 0 Then
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "���i�����ρI", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "���i�����ρI", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        
                        End If
                                                             
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                                                 
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        
                        Erase ID_KANRI_TBL(ING_No).KEN_DEN_TBL
                        For j = 0 To UBound(Y_SYU_TBL)
                        
                            ReDim Preserve ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j)
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO = Y_SYU_TBL(j).SEQ_NO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO = Y_SYU_TBL(j).HIN_NO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO = Y_SYU_TBL(j).SURYO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI = Y_SYU_TBL(j).SUMI
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F = Y_SYU_TBL(j).CANCEL_F
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KAN_KBN = Y_SYU_TBL(j).KAN_KBN      '2007.05.14
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F = Y_SYU_TBL(j).KONPOU_F        '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_CND = Y_SYU_TBL(j).KONPOU_CND    '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).TOTAL_SU = Y_SYU_TBL(j).TOTAL_SU        '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SAI_SU = Y_SYU_TBL(j).SAI_SU            '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU = Y_SYU_TBL(j).KUTI_SU          '2010.01.21
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = Y_SYU_TBL(j).PRINT_SUMI    '2010.01.21
                        
                                                                                                        '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).OKURI_NO_SEQ = Y_SYU_TBL(j).OKURI_NO_SEQ
                        
                        Next j
                        
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO = OKURI_NO
                        
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI = OKURI_SAKI
                        
                        '�^�����
                        ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA = UNSOU_KAISHA
                        
                        '�o�ד�
                        ID_KANRI_TBL(ING_No).KEN_SYUKA_YMD = SYUKA_YMD
                        '�Z��
                        ID_KANRI_TBL(ING_No).KEN_JYUSHO = JYUSHO
                        
                        '���l
                        ID_KANRI_TBL(ING_No).KEN_BIKOU = BIKOU
                        '������
                        ID_KANRI_TBL(ING_No).KEN_MUKE_NAME = MUKE_NAME
                        
                        '�}��
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO_MAX = OKURI_NO_MAX
                        '�d�b�ԍ�
                        ID_KANRI_TBL(ING_No).KEN_TEL_NO = KEN_TEL_NO
                        '���X�R�[�h
                        ID_KANRI_TBL(ING_No).KEN_TYAKUTEN = KEN_TYAKUTEN    '2017.04.06
                        
                        
                        
                        '���يJ�n�y�[�W��
                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                        
                        ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = KUTI_SU
                        If Label_Print_Total_Su_Proc(KUTI_SU, PRINT_TOTAL_SU) Then
                        
                    
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        
                        
                        
                        End If
                        ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = PRINT_TOTAL_SU
                        
                        
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                        
                        
                        If KONPOU_ON <> 0 Then
                            If KONPOU_ON = KONPOU_ON_SUMI Then
                                                    
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                End If
                                                    
                            End If
                        End If
                        
                        
                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = "" Then
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                            
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, OKURI_SAKI)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, OKURI_SAKI)
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                            If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = "" Then
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
                            End If
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
    
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                
                
                
                        Else
                
                
                
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                                                                                    'BOX����
                                                                                    
                                                                                    
                            '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                            Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
        
    ''' �i�ԒP�ʂł̊ۂߏ���
KONPOU_ON = KONPOU_ON - KONPOU_ON_SUMI          '2011.03.17
                            
                            
Select Case KONPOU_ON                           '2011.03.17
                    
''''''''''''Select Case (KONPOU_ON - KONPOU_ON_SUMI)     '2011.03.17
                            
                                Case 0
                                '�W������Ȃ�
                                
                                    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                                    If KONPOU_ON_SUMI <> 0 And KONPOU_OFF_SUMI = 0 Then
                                        If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                        
                                        
                                        
                                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                            
                                            '-----------------------------------------------�w�b�_�[
                                            Call Wel_Head_Text_Proc
                                            '-----------------------------------------------�P�s��
                                            Call Wel_DETAIL_0_Text_Proc
                                            '-----------------------------------------------�Q�s��
                                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                            '-----------------------------------------------�R�s��
                                                                                                    'BOX����
                                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                                    '�\�����e
                                                                                                    
                                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                                    
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
                                        
                                        
                                        
                                        
                                        
                                        
                                            TOTAL_KUTI_SU = 1
                                            KUTI_SU_INPUT_F = True
                                            TOTAL_SAI_SU = 0#
                                                
                                                
                                                
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                            End If
                                                
                                                
                                    
                                            Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                            ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                            ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                    
                                    
                                    
                                            Sendbuf = Text_Create_Proc()
                                        
                                            Inspe_Proc_DEN = False
                                            Exit Function
                                        
                                        
                                        End If
                                    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                                    
                                    
                                    
                                    '-----------------------------------------------�{�f�B
                                    Call Wel_Hin_No_Req_Text_Proc(Sumi_CNT, Y_SYU_CNT)
            
                                    Sendbuf = Text_Create_Proc()
                                Case Else
                                '�W�������
                                    '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                                    Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                            
                                    '-----------------------------------------------�Q�s��
                                    Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                    '-----------------------------------------------�R�s��
                                                                                            'BOX����
                                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                            '�\�����e
                                                                                            
                                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                                    Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                            '�\�����e
                                                                                        '�\�����e
                                    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                            "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                            "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
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
                                    '-----------------------------------------------�T�s��
                                    Call Wel_Clear_Text_Proc
            
                                    Sendbuf = Text_Create_Proc()
                            
                            End Select
                        End If
                
                
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i����󇂁j
                
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    '�����
                    Case Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO, _
                                LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
                        
                        If Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) = LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) Then
                            
                            If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) > Len(LCD_OKURI_NO_S) Then
                                If Left(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), Len(LCD_OKURI_NO_S)) = LCD_OKURI_NO_S Then
                                    OKURI_NO = Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
                                Else
                                    OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                                End If
                            Else
                                OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                            End If
                        Else
                            OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        End If
                        
                        If Trim(OKURI_NO) = Trim(KEN_CHARTER_CD) Or Trim(OKURI_NO) = Trim(KEN_AKABOU_CD) Or Trim(OKURI_NO) = Trim(KEN_LOGISTIC_CD) Then
                        
                        '�`���[�^�[��   2010.01.21
                        
                        Else
'2009.10.14                         If Len(OKURI_NO) < 11 Or Len(OKURI_NO) > 13 Then
'                            If Len(OKURI_NO) < 10 Or Len(OKURI_NO) > 13 Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
'
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                Inspe_Proc_DEN = False
'                                Exit Function
'                            End If
                        
                            If Not IsNumeric(OKURI_NO) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "����󇂃G���[", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            End If
                        
                        
                        
                            If OKURI_NO_CHECK_PROC(OKURI_NO, OKURI_NO_F, FUKUYAMA_CHK_F) Then
                            End If
                            
                            
                            
                            
                            
                            If Not OKURI_NO_F Then
                            
                        
                        
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "����󇂃G���[", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                                                
                            End If
                        
                        
                        
                            '2009.04.28
                            If FUKUYAMA_CHK_F Then
                            
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "���R �����޼ޯĴװ", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "���R �����޼ޯĴװ", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            
                            End If
                            '2009.04.28
                        
                        
                        
                        
                        
                        
'                            Select Case Len(Trim(OKURI_NO))
'
'                                Case FUKUYAMA_Length
'                                Case SEIBU_Length
'                                Case KURUME_Length
'
'                                    For k = 0 To UBound(KURUME_CODE)
'
'                                        If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
'                                            Exit For
'                                        End If
'                                    Next k
'
'                                    If k > UBound(KURUME_CODE) Then
'
'
'                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
'
'                                        Sendbuf = Text_Create_Proc()
'                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                        Inspe_Proc_DEN = False
'                                        Exit Function
'
'                                    End If
'
'                                Case SAGAWA_Length, YAMATO_Length
'
'                                    For k = 0 To UBound(KURUME_CODE)
'
'                                        If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
'                                            Exit For
'                                        End If
'                                    Next k
'
'                                    If k > UBound(SAGAWA_CODE) Then
'
'                                        For k = 0 To UBound(YAMATO_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
'                                                Exit For
'                                            End If
'
'                                        Next k
'
'                                        If k > UBound(YAMATO_CODE) Then
'
'                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
'
'                                            Sendbuf = Text_Create_Proc()
'                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                            Inspe_Proc_DEN = False
'                                            Exit Function
'
'
'
'                                        End If
'
'                                    End If
'
'
'                            End Select
                        
                        End If
                    
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO = OKURI_NO

                
                        '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        '------------------------------------   �o�ח\��̏���
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                'ID��
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
            
                            Do
                            
                                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrKeyNotFound
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_DEN = False
                                        GoTo Abort_Tran
                                     Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_DEN = False
                                        GoTo Abort_Tran
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                        
                            Loop
    
                            '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                    
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")         '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                            
                                Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, OKURI_NO)
                                            
                                '�^����Еϊ�
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, 3) = UNSOU_KAISHA_CODE Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU_KAISHA_NAME)
'                                End If
'                                '�V�^����Еϊ� 2007.01.09
'
'                                If KURUME_F Then        '�v����
'                                    For k = 1 To UBound(KURUME)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(KURUME(k))) = KURUME(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
'
'                                If FUKUYAMA_F Then      '���R
'                                    For k = 1 To UBound(FUKUYAMA)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(FUKUYAMA(k))) = FUKUYAMA(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
'
'                                If SAGAWA_F Then        '����
'                                    For k = 1 To UBound(SAGAWA)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(SAGAWA(k))) = SAGAWA(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
                                                    
                                                    
                                                    
                                                    
                                                    




                                                    
                                                    
                                                    
                                                    
'                                Select Case Len(Trim(OKURI_NO))
'
'                                    Case FUKUYAMA_Length
'                                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA_Name)
'                                    Case SEIBU_Length
'                                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEIBU_Name)
'
'                                    Case KURUME_Length
'
'                                        For k = 0 To UBound(KURUME_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
'                                                Exit For
'                                            End If
'                                        Next k
'
'                                        If k > UBound(KURUME_CODE) Then
'                                        Else
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME_Name)
'                                        End If
'
'                                    Case SAGAWA_Length, YAMATO_Length
'
'                                        For k = 0 To UBound(KURUME_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
'                                                Exit For
'                                            End If
'                                        Next k
'
'                                        If k > UBound(SAGAWA_CODE) Then
'
'                                            For k = 0 To UBound(YAMATO_CODE)
'
'                                                If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
'                                                    Exit For
'                                                End If
'
'                                            Next k
'
'                                            If k > UBound(YAMATO_CODE) Then
'                                            Else
'
'                                                Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, YAMATO_Name)
'                                            End If
'
'                                        Else
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA_Name)
'                                        End If
'
'
'                                End Select
                                                    
                                                    
                                                    
                                                    
                                Call OKURI_NO_SET_PROC(OKURI_NO)
                                                    
                                
                                
                                
                                
                                
                                
                                
                                
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                
                                
                                
                                
                                                    
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_DEN = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
                                            '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        '-----------------------------------------------�w�b�_�[
                        Call Wel_Head_Text_Proc
                        
                        '-----------------------------------------------�P�s��
                        Call Wel_DETAIL_0_Text_Proc
                                                                                'BOX����
                                                                                
                                                                                
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                        
KONPOU_ON = KONPOU_ON - KONPOU_ON_SUMI              '2011.03.17
        
    ''' �i�ԒP�ʂł̊ۂߏ���
Select Case KONPOU_ON                               '2011.03.17
'''''Select Case (KONPOU_ON - KONPOU_ON_SUMI)     '2011.03.17



                        
                            Case 0
                            '�W������Ȃ�
                            
                                
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                                If KONPOU_ON_SUMI <> 0 And KONPOU_OFF_SUMI = 0 Then
                                    If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                    
                                    
                                    
                                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                        
                                        '-----------------------------------------------�w�b�_�[
                                        Call Wel_Head_Text_Proc
                                        '-----------------------------------------------�P�s��
                                        Call Wel_DETAIL_0_Text_Proc
                                        '-----------------------------------------------�Q�s��
                                        Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                        '-----------------------------------------------�R�s��
                                                                                                'BOX����
                                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                                '�\�����e
                                                                                                
                                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                                
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
                                    
                                    
                                    
                                    
                                    
                                    
                                        TOTAL_KUTI_SU = 1
                                        KUTI_SU_INPUT_F = True
                                        TOTAL_SAI_SU = 0#
                                            
                                            
                                            
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                        End If
                                            
                                            
                                
                                        Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                        ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                        ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                
                                
                                
                                        Sendbuf = Text_Create_Proc()
                                    
                                        Inspe_Proc_DEN = False
                                        Exit Function
                                    
                                    
                                    End If
                                End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                                
                                
                                
                                '-----------------------------------------------�{�f�B
                                Call Wel_Hin_No_Req_Text_Proc(Sumi_CNT, Y_SYU_CNT)
        
                                Sendbuf = Text_Create_Proc()
                            Case Else
                            '�W�������
                        
                        
                                '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g 2011.03.04
                                Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                        
                        
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                '-----------------------------------------------�R�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                                Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                        '�\�����e
                                                                                    '�\�����e
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                        "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                        "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                                                                        '���l�����\��
                                Send_Text.Box_Type(3).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                        '�����J�[�\���ʒu
                                Send_Text.Box_Type(3).Start_Pos = "01"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                        '���͌���
                                '2010.12.07
'                                Send_Text.Box_Type(3).Max_Size = "13"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                Send_Text.Box_Type(3).Max_Size = "20"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                '2010.12.07
                                                                                        
                                Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                '-----------------------------------------------�T�s��
                                Call Wel_Clear_Text_Proc
        
                                Sendbuf = Text_Create_Proc()
                        
                        End Select
                    End Select
                Next i
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                Select Case Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), 2)
                    
                    Case LCD_Hinban     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                
                        End Select
                '�W������L�莞�̕i�ԃ`�F�b�N
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                        '�Y���i�ԗL��������
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                Exit For
                            End If
                        Next j
                        
                        If j > UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")      '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")  '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        End If
                        
                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND <> "1" Then
                        
                            If KONPOU_ON <> KONPOU_ON_SUMI Then
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F = "1" Then
                            
                            
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")      '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")  '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_DEN = False
                                    Exit Function
                            
                                End If
                            End If
                        
                        End If
                        '��ݾق�����
                        CANCEL_F = True
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                    CANCEL_F = False
                                    Exit For
                                End If
                            
                            End If
                        Next j
                        
                        If CANCEL_F Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�L�����Z���i�Ԃł��B", "")        '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�L�����Z���i�Ԃł��B", "")    '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        
                        
                        End If
                        
                        
                        
                                
                        '���i�ς݂�����
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                Exit For
                            End If
                        Next j
                
                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���i�ς݁I", "")          '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���i�ς݁I", "")      '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        End If
                
                        '���o�ɂ�����   2007.05.14
                        If Inspection_Flg = 0 Then
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KAN_KBN <> KAN_KBN_FIN Then
                                        Exit For
                                    End If
                                End If
                            Next j
                                            
                            If j <= UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���o�ɕ��L��I�I", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���o�ɕ��L��I�I", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            End If
                        End If
                        '���o�ɂ�����   2007.05.14
                        ID_KANRI_TBL(ING_No).KEN_HINBAN = Hinban
                
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
                        
                        
                        '-----------------------------------------------�w�b�_�[
                        Call Wel_Head_Text_Proc
                        '-----------------------------------------------�P�s��
                        Call Wel_DETAIL_0_Text_Proc

''' �i�ԒP�ʂł̊ۂߏ���
                        
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
''' �i�ԒP�ʂł̊ۂߏ���
                        '-----------------------------------------------�Q�s��
                        Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                        '-----------------------------------------------�R�s��
                                                                                'BOX����
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '�\�����e
                        
                        
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Hinban)
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
                        If Inspection_QTY = 1 Then

                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        Else
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM                             '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM        '2007.04.21
                        End If
                        
                        Y_SYU_CNT = 0
                        SYUKA_QTY = 0
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                End If
                            End If
                        Next j
                                                                                '�\�����e
                        
                        If Y_SYU_CNT < 2 Then

                            If Inspection_QTY = 1 Then

                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide))
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka_Su1)                         '2007.04.21
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka_Su1)    '2007.04.21
                            End If
                                                                                
                        Else
                        
                            If Inspection_QTY = 1 Then
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide) & "*")
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide) & "*")
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka_Su2)                       '2007.04.21
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka_Su2)  '2007.04.21
                            End If
                        
                        End If
                                                                                
                                                                                '���l�����\��
                        If Inspection_QTY = 1 Then
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                        Else
                            Send_Text.Box_Type(4).INIT = ""                                                     '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""                                '2007.04.21
                        End If
                                                                                '�����J�[�\���ʒu
                        If Inspection_QTY = 1 Then

                            Send_Text.Box_Type(4).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                        Else
                            Send_Text.Box_Type(4).Start_Pos = "10"                                          '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "10"                     '2007.04.21
                        End If
                                                                                
                                                                                '���͌���
                        If Inspection_QTY = 1 Then
                            Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                        Else
                            Send_Text.Box_Type(4).Max_Size = "07"                                           '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "07"                      '2007.04.21
                        End If
                                                                                
                                                                                
                        '2009.04.15
                        If SYUKA_QTY > 1 Then
                            Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                                                                
                        End If
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        
                        Sendbuf = Text_Create_Proc()
                
                End Select
            
            Next i
'''''''''''''''''''''''''''''''
    
    
        Case Step_Sagyo4_RES        '�S��ڂ̎�M�i���i���@��M�j
            
            For i = 0 To M_Gyo - 1
                
'                Select Case RTrim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), _
'                                Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), 10))
                    
                    Case LCD_Syuka_Su1, LCD_Syuka_Su2, "�o�א��F"  '�o�א�(���i��)
                        
                        SURYO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        
                        If Not IsNumeric(SURYO) Then
                        
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        
                        End If
                
                        Y_SYU_CNT = 0
                        SYUKA_QTY = 0
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                End If
                            End If
                        Next j
                
                        If CLng(SURYO) <> SYUKA_QTY Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�א��G���[", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�א��G���[", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_DEN = False
                            Exit Function
                        End If
                
                End Select
            
            Next i
            
            Y_SYU_CNT = 0
            SYUKA_QTY = 0
            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
            
                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                    
                        Y_SYU_CNT = Y_SYU_CNT + 1
                        SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                    End If
                End If
            Next j
            
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                            
            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                        
                        '------------------------------------   �o�ח\��̏���
                        Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                            'ID��
                        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
        
                        Do
                        
                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrKeyNotFound
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2107.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_DEN = False
                                    GoTo Abort_Tran
                                 Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_DEN = False
                                    GoTo Abort_Tran
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                    
                        Loop
        
                    '''���s�̎g�p���p�����邽��
                    '''Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    '''Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                                    
                        Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
                        Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, Format(Now, "HHMMSS"))
                                                
                                                    '�o�ח\�菑����
                        Do
                            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2107.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_DEN = False
                                    GoTo Abort_Tran
                            
                                Case Else
                                    
                                    Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                    Inspe_Proc_DEN = SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                        Loop
                        '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                        
                        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)    'ID��
        
                        Do
                        
                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrKeyNotFound
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_DEN = False
                                    GoTo Abort_Tran
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_DEN = False
                                    GoTo Abort_Tran
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                    
                        Loop
                                            
                                            
                        Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, Format(Now, "YYYYMMDDHHMMSS"))
                        Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
                                            
                        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
                        '�^����Еϊ�
'                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, 3) = UNSOU_KAISHA_CODE Then
'                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU_KAISHA_NAME)
'                        End If
'                        '�V�^����Еϊ� 2007.01.09
'
'                        If KURUME_F Then        '�v����
'                            For k = 1 To UBound(KURUME)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(KURUME(k))) = KURUME(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
'
'                        If FUKUYAMA_F Then      '���R
'                            For k = 1 To UBound(FUKUYAMA)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(FUKUYAMA(k))) = FUKUYAMA(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
'
'                        If SAGAWA_F Then        '����
'                            For k = 1 To UBound(SAGAWA)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(SAGAWA(k))) = SAGAWA(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
                                                    
                                                    
                                                    
                        Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                        Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                                    
                                                    
                                                    '�o�ח\��(νĲҰ��)������
                        Do
                            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_DEN = False
                                    GoTo Abort_Tran
                            
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                    Inspe_Proc_DEN = SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                        Loop
                                            
                                            
                        '------------------------------------   �݌Ɉړ������̏���
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                            MENU_NO = ""
                        End If
                                            
                        '�����o�ׂ͂̈̓ǂݍ���
                        Call UniCode_Conv(K0_MTS.MUKE_CODE, ID_KANRI_TBL(ING_No).MTS_CODE)
                        Call UniCode_Conv(K0_MTS.SS_CODE, "")
                        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(MTSREC.MUKE_DNAME, "")
                                Call UniCode_Conv(MTSREC.MUKE_NAME, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "������}�X�^", 0)
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                                            
                        sts = IDOREKI_OUTPUT_PROC("", _
                                                    "", _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    ID_KANRI_TBL(ING_No).KEN_HINBAN, _
                                                    "", _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    0, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    FILE_RETRY, _
                                                    CYU_KBN_SPO, _
                                                    Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) & " �����:" & Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)), _
                                                    , , , , MENU_NO, _
                                                    ID_KANRI_TBL(ING_No).MTS_CODE, _
                                                    "", _
                                                    ID_KANRI_TBL(ING_No).ID_NO & "-" & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO, , , , 1)
                        Select Case sts
                            Case False      '����I��
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc_DEN = SYS_ERR
                                GoTo Abort_Tran
                        End Select
                                            
                        '���i�ρI�I
                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI = True
                        
                        '�^�����
                        ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA = StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)
                                        
                    End If
                End If
            
            Next j
            '���۸ޏo��    2009.04.17
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
                                                     ID_KANRI_TBL(ING_No).KEN_HINBAN, , , , , _
                                                     ID_KANRI_TBL(ING_No).ID_NO) Then
                    Inspe_Proc_DEN = SYS_ERR
                    Exit Function
                End If
            End If
            '���۸ޏo��    2009.04.17
                                '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                    
                    
                    
                    
                    
                    
            '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
            Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                    
            Select Case ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND
            
            
                Case "1"
                    '�W������Ȃ�
                
                
                    KENPIN_END = True
                    
                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                KENPIN_END = False
                                Exit For
                            End If
                        End If
                    Next j
                
                
                
                
                
                
                    Select Case KENPIN_END
                    
                        Case False
                            '�c����@���i�Ԃ�
''' �׎D���u
                            If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Then
                                ID_KANRI_TBL(ING_No).LABEL_ON = False
                            Else
                            
                                PRINT_OFF = False
                                
                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) Then
                                    If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                        PRINT_OFF = True
                                    End If
                                End If
                                
                                If Not PRINT_OFF Then
                                
                                    '2010.06.16
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                                            
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                            
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            End If
                                        
                                        Case BtErrKeyNotFound
                                        
                                        
                                            Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                
                                
                                    '2010.06.16
                                
                                    PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                            
                                    Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1
                                                            
                                    PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                
                                    ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                    
                                    'Y_SYU_CNT = 0
                                    SYUKA_QTY = 0
                                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                    
                                        If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                            
                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                            
                                                'Y_SYU_CNT = Y_SYU_CNT + 1
                                                SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                            
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                                    PRINT_OFF = True
                                                Else
                                            
                                                   ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = True
                                                
                                                End If
                                            
                                            End If
                                        End If
                                    Next j
                                End If
                                
                                If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD)) Then
                                    PRINT_MAISU = PRINT_MAISU - 1
                                    If PRINT_MAISU < 1 Then
                                        
                                    '�}�ԍX�V   2010.02.15
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                    
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                            
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                        
                                                        GoTo Abort_Tran
                                                    
                                                    End If
                                                                                                    
                                                
                                                End If
                                            Else
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                    
                                
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                            
                                                            GoTo Abort_Tran
                                                        
                                                        End If
                                                    
                                                    
                                                    
                                                    End If
                                                                        
                                                End If
                                            End If
                                        Next j
                                    '�}�ԍX�V   2010.02.15
                                        
                                        PRINT_OFF = True
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                    
                                    
                                    Else
'                                            PRINT_MAISU = PRINT_MAISU - 1
                                        Start_Page_No = Start_Page_No + 1
                                    End If
                                End If
                                
                                If Not PRINT_OFF Then
                                                            
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                    '2010.06.16
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    '2010.06.16
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                            
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            End If
                                        
                                        Case BtErrKeyNotFound
                                        
                                        
                                            Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                                            
                                                            
'2010.02.21                                    PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                            
'2010.02.21                                    Start_Page_No = Start_Page_No + 1
                                                            
'2010.02.21                                    PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                                            
                                    ID_KANRI_TBL(ING_No).LAST_END_PAGE = Start_Page_No + PRINT_MAISU - 1    '2012.04.01
                                                            
                                                            
                                    If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                    End If
                                
                                
                                
                                    '�}�ԍX�V   2010.02.15
                                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                
                                        If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                        
                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            
                                                If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 And _
                                                    (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD)) Then

                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU, "000")) Then
                                                    
                                                        GoTo Abort_Tran
                                                    End If
                                            
                                            
                                                Else
                                            
                                            
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                    
                                                        GoTo Abort_Tran
                                                    End If
                                            
                                                End If
                                            End If
                                        Else
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                
                            
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    
                                                    End If
                                                
                                                
                                                
                                                End If
                                                                    
                                                                    
                                                                    
                                            End If
                                        End If
                                    Next j
                                    '�}�ԍX�V   2010.02.15
                                
                                
                                
                                
'2010.02.21                                    If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
'                                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
'                                    Else
'                                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
'                                    End If
                                    
'                                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                        If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'                                                If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                    GoTo Abort_Tran
'                                                End If
'
'                                            End If
'                                        End If
'                                    Next j
                                    
                                    '�f�[�^���M
                                                                
                                    ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                                                
                                                                
                                    ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                                
                                    ID_KANRI_TBL(ING_No).LABEL_ON = True
                                
                                    ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No + PRINT_MAISU - 1
                                    '-----------------------------------------------�w�b�_�[
                                
                                    Call Wel_Head_Print_Text_Proc(FileName)
                                
                                    Sendbuf = Text_Create_Proc()
                                    
                                
                                    Inspe_Proc_DEN = False
                                    Exit Function
                                End If
                            
                            End If
''' �׎D���u
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            
                            '-----------------------------------------------�w�b�_�[ 02.24
                            Call Wel_Head_Text_Proc
                            
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
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
                            '2010.12.07
'                            Send_Text.Box_Type(3).Max_Size = "13"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                            Send_Text.Box_Type(3).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                            '2010.12.07
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                    
                    
                    
                    
                    
                        Case True
                            '�c�Ȃ��@������
                    
                    
''' �׎D���u
                                PRINT_OFF = False
            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" And KONPOU_OFF_SUMI = 0 Then
                                    PRINT_OFF = True
                                Else
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "2" Then
                                    
                                        If KONPOU_ON = KONPOU_ON_SUMI Then
                
                                            PRINT_OFF = True
                
                                        End If
                
                                    End If
                                End If
                                
                                If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Or PRINT_OFF Then
                                    ID_KANRI_TBL(ING_No).LABEL_ON = False
                                Else
                                    
            '                        Y_SYU_CNT = 0
            '                        SYUKA_QTY = 0
                                    
                                    PRINT_OFF = False
                                    
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) Then
                                        If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                            
                                                        '�}�ԍX�V   2010.02.15
                                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                                                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                    Start_Page_No = 1
                                                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                    
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                        
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                                    
                                                        PRINT_OFF = True
                                        
                                                    End If
                                                
                                                Else
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                        
                                    
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            
                                                            End If
                                                    
                                                        
                                                        End If
                                                                            
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                        '2010.06.16
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            
                                                If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                                
                                                    Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                                
                                                End If
                                            
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                                Sendbuf = Text_Create_Proc()
                                                GoTo Abort_Tran
                                        End Select
                                        
                                        PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                                
                                        Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1
                                                                
                                        PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                        
                                        If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                               Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                               Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD)) Then
                                            PRINT_MAISU = PRINT_MAISU - 1
                                            
                                            
                                            
                                            If PRINT_MAISU < 1 Then
                                                
                                                        '�}�ԍX�V   2010.02.15
                                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                            
                                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    
                                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                            
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            End If
                                        
                                                        End If
                                                    
                                                    
                                                    
                                                    Else
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                            
                                        
                                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                    
                                                                    GoTo Abort_Tran
                                                                End If
                                                            
                                                            
                                                        
                                                            End If
                                                                                
                                                        End If
                                                    
                                                    
                                                    
                                                    End If
                                                Next j
                                                        '�}�ԍX�V   2010.02.15
                                                
                                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                                
                                                PRINT_OFF = True
                                            Else
            '                                    PRINT_MAISU = PRINT_MAISU - 1
                                                Start_Page_No = Start_Page_No + 1
                                            End If
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                        
                                        Y_SYU_CNT = 0
                                        SYUKA_QTY = 0
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                        
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                
                                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                                
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                                        PRINT_OFF = True
                                                    Else
                                                
                                                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = True
                                                    
                                                    End If
                                                
                                                End If
                                            End If
                                        Next j
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                        
                                        
                                        If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                        End If
                                        
                                        
                                        
                                        
                                        
                                                        '�}�ԍX�V   2010.02.15
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                    
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                            
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                                        
                                                    If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 And Start_Page_No = 2 Then
                                                    
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    Else
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    End If
                                                
                                                End If
                                            
                                            
                                            
                                            Else
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                    
                                
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    
                                                    End If
                                                                        
                                                End If
                
                                            
                                            
                                            
                                            
                                            End If
                                        Next j
                                                        '�}�ԍX�V   2010.02.15
                                        
                                        
                                        
                                        
                                        If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
                                        Else
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
                                        End If
                                        
'                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'
'
'                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                        GoTo Abort_Tran
'                                                    End If
'                                                End If
'                                            End If
'                                        Next j
                                        
                                        '�f�[�^���M
                                                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 2
                                                                    
                                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                                    
                                        '-----------------------------------------------�w�b�_�[
                                
                                        Call Wel_Head_Print_Text_Proc(FileName)
                                    
                                        Sendbuf = Text_Create_Proc()
                                        
                                    
                                        Inspe_Proc_DEN = False
                                        Exit Function
                                    
                                    End If
                                End If
            ''' �׎D���u
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Text_Proc
                                '-----------------------------------------------�P�s��
                                Call Wel_DETAIL_0_Text_Proc
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                '-----------------------------------------------�R�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                        
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
                                '-----------aaa------------------------------------�S�s��
                                
'����INPUT �P
                                
                                wkKONPO_F = ""
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                    
                                        wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                                        Exit For
                                    End If
                                Next j
                                
                                If wkKONPO_F = "1" Then
                                                        
                                    If Inspection_Input Then
                                        KUTI_SU_INPUT_F = False
                                    Else
                                        KUTI_SU_INPUT_F = True
                                    End If
                                
                                
                                    TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                    TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                            
                                Else
                                    TOTAL_KUTI_SU = 1
                                    KUTI_SU_INPUT_F = True
                                    TOTAL_SAI_SU = 0#
                                End If
                                            
                                            
                                            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                End If
                                            
                                            
                                If KUTI_SU_INPUT_F Then
                                
                                    Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                
                                
                                Else
                                    Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                
                                    
                                    If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                    
                                    
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                                
                                
                                
                                End If
                                
                                Sendbuf = Text_Create_Proc()
                        
                        End Select
                        
                
                
                Case "2"
                    '�W������̂�
                
                    KENPIN_END = True
                    
                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                KENPIN_END = False
                                Exit For
                            End If
                        End If
                    Next j
                
                
                
                
                
                
                    Select Case KENPIN_END
                    
                        Case False
                            '�c����@���i�Ԃ�
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            
                            
                            '-----------------------------------------------�w�b�_�[ 02.24
                            Call Wel_Head_Text_Proc
                            
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                                                                                '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                                                                    
                                                                                    
                                                                                    
                                                                                    '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    '���͌���
                            '2010.12.07
'                            Send_Text.Box_Type(3).Max_Size = "13"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                    
                            Send_Text.Box_Type(3).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                            '2010.12.07
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                    
                    
                    
                    
                    
                        Case True
                            '�c�Ȃ��@������
                    
                    
''' �׎D���u
                                PRINT_OFF = False
            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" And KONPOU_OFF_SUMI = 0 Then
                                    PRINT_OFF = True
                                Else
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "2" Then
                                    
                                        If KONPOU_ON = KONPOU_ON_SUMI Then
                
                                            PRINT_OFF = True
                
                                        End If
                
                                    End If
                                End If
                                
                                If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Or PRINT_OFF Then
                                    ID_KANRI_TBL(ING_No).LABEL_ON = False
                                Else
                                    
            '                        Y_SYU_CNT = 0
            '                        SYUKA_QTY = 0
                                    
                                    PRINT_OFF = False
                                    
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) Then
                                        If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                            
                                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                                                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                    
                                                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                    
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                                        Start_Page_No = 1
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        PRINT_OFF = True
                                                    End If
                                                
                                                
                                                Else
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                        
                                    
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            
                                                            End If
                                                        
                                                        
                                                        
                                                        End If
                                                                            
                                                    End If
                                                
                                                
                                                
                                                
                                                
                                                End If
                                            Next j
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                        '2010.06.16
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        '2010.06.16
                                        Select Case sts
                                            Case BtNoErr
                                            
                                                If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                                
                                                    Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                                
                                                End If
                                            
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                                Sendbuf = Text_Create_Proc()
                                                GoTo Abort_Tran
                                        End Select
                                        
                                        PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                                
                                        Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1
                                                                
                                        PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                        
                                        If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                                Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                                Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD)) Then
                                            PRINT_MAISU = PRINT_MAISU - 1
                                            If PRINT_MAISU < 1 Then
                                                
                                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                            
                                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    
                                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                        
                                                        '�}�ԍX�V   2010.02.15
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            
                                                            End If
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                        
                                        
                                        
                                                        End If
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    Else
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                            
                                        
                                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                    
                                                                    GoTo Abort_Tran
                                                                End If
                                                                
                                                        
                                                            
                                                            End If
                                                                                
                                                        End If
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    End If
                                                Next j
                                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                                PRINT_OFF = True
                                            Else
            '                                    PRINT_MAISU = PRINT_MAISU - 1
                                                Start_Page_No = Start_Page_No + 1
                                            End If
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                        
                                        Y_SYU_CNT = 0
                                        SYUKA_QTY = 0
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                        
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                
                                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                                
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                                        PRINT_OFF = True
                                                    Else
                                                
                                                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = True
                                                    
                                                    End If
                                                
                                                End If
                                            End If
                                        Next j
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                        
                                        If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                        End If
                                        
                                        
                                        
                                        
                                        
                                        '�}�ԍX�V   2010.02.15
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                    
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                            
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    End If
                                                End If
                                            
                                            Else
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                    
                                
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    
                                                    
                                                    End If
                                                                        
                                                End If
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            End If
                                        Next j
                                        '�}�ԍX�V   2010.02.15
                                        
                                        
                                        
                                        
                                        
                                        If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
                                        Else
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
                                        End If
                                        
'                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'
'
'                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                        GoTo Abort_Tran
'                                                    End If
'
'                                                End If
'                                            End If
'                                        Next j
                                        
                                        '�f�[�^���M
                                                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 2
                                                                    
                                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                                    
                                        '-----------------------------------------------�w�b�_�[
                                
                                        Call Wel_Head_Print_Text_Proc(FileName)
                                    
                                        Sendbuf = Text_Create_Proc()
                                        
                                    
                                        Inspe_Proc_DEN = False
                                        Exit Function
                                    
                                    End If
                                End If
            ''' �׎D���u
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Text_Proc
                                '-----------------------------------------------�P�s��
                                Call Wel_DETAIL_0_Text_Proc
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                '-----------------------------------------------�R�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                        
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
                                '-----------aaa------------------------------------�S�s��
                                
'����INPUT �Q
                                wkKONPO_F = ""
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                    
                                        wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                                        Exit For
                                    End If
                                Next j
                                
                                If wkKONPO_F = "1" Then
                                                        
                                    If Inspection_Input Then
                                        KUTI_SU_INPUT_F = False
                                    Else
                                        KUTI_SU_INPUT_F = True
                                    End If
                                
                                
                                    TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                    TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                            
                                Else
                                    TOTAL_KUTI_SU = 1
                                    KUTI_SU_INPUT_F = True
                                    TOTAL_SAI_SU = 0#
                                End If
                                            
                                            
                                            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                End If
                                            
                                            
                                If KUTI_SU_INPUT_F Then
                                
                                    Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                
                                
                                Else
                                    Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                
                                    If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                
                                
                                
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                                
                                
                                
                                End If
                                
                                Sendbuf = Text_Create_Proc()
                        
                        End Select
                    
                Case "3"
                    '����
            
                    
                    '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                    Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
            
            
                    Select Case (KONPOU_ON - KONPOU_ON_SUMI)
            
                        Case 0
                            '������
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                            
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                    
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
                            '-----------aaa------------------------------------�S�s��
'����input �R
                            wkKONPO_F = ""
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            
                                If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                
                                    wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                                    Exit For
                                End If
                            Next j
                            
                            If wkKONPO_F = "1" Then
                                                    
                                If Inspection_Input Then
                                    KUTI_SU_INPUT_F = False
                                Else
                                    KUTI_SU_INPUT_F = True
                                End If
                            
                            
                                TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                        
                            Else
                                TOTAL_KUTI_SU = 1
                                KUTI_SU_INPUT_F = True
                                TOTAL_SAI_SU = 0#
                            End If
                                        
                                        
                                        
                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                            End If
                                        
                                        
                            If KUTI_SU_INPUT_F Then
                            
                                Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                            
                            
                            Else
                                Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                            
                                
                                
                                
                                
                                If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                End If
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                                ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                            
                            
                            
                            End If
                        
                        
                            Sendbuf = Text_Create_Proc()
                        
                        
                        
                        
                        
                        Case Else
                            '�i�Ԃ�
            
                            '�c����@���i�Ԃ�
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                                                                                '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                                                                    
                                                                                    
                                                                                    
                                                                                    '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                                        
                                                                                        '���͌���
                            '2010.12.07
'                            Send_Text.Box_Type(3).Max_Size = "13"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                            Send_Text.Box_Type(3).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                            '2010.12.07
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                    End Select
            
            End Select
                    
                    
                    
                    
        Case Step_Sagyo5_RES        '�T��ڂ̎�M�i�����j
                
            For i = 0 To M_Gyo - 1
                
                Select Case Left(Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)), 6)
                    '����
                    Case LCD_KUTI_SU_S
                
                
                        If ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU < 0 Then
                        
                
                
                
                            If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")         '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")     '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            
                            End If
                    
                            KUTI_SU = CInt(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If KUTI_SU < 1 Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")         '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")     '2107.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            End If
                        Else
                            KUTI_SU = ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU
                        End If
                    
                    
                    
                    '�ː�
                    Case LCD_SAI_SU_S
                
                        If ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU < 0 Then
                        
                        
                            If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")         '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")     '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            
                            End If
                    
                            SAI_SU = CDbl(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If SAI_SU <= 0 Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")         '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")     '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_DEN = False
                                Exit Function
                            End If
                        Else
                            SAI_SU = ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU
                        
                        
                            If SAI_SU < 1 Then
                                SAI_SU = 1
                            Else
                                If SAI_SU > 1 Then
                                    SAI_SU = CLng(ToHalfAdjust(CCur(SAI_SU), 0))
                                End If
                            End If
                        
                        End If
                    
                        '�����ő��������l�� 2010.01.21
                            
                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                            If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                
                                
                                    If Label_Print_Total_Su_Proc(KUTI_SU, PRINT_TOTAL_SU) Then
                                
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                
                                
                                
                                Else
                                
                                    If Label_Print_Total_Su_Proc(0, PRINT_TOTAL_SU) Then
                                
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                
                                End If
                            End If
                        Next j
                                        
                            
                            
                            
                        
'                        If Label_Print_Total_Su_Proc(KUTI_SU, PRINT_TOTAL_SU) Then
'
'                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                            Sendbuf = Text_Create_Proc()
'                            Exit Function
'                        End If
                    
                        ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = PRINT_TOTAL_SU
                
                        
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                            If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU = KUTI_SU
                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SAI_SU = SAI_SU
                            Else
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU <= 1 Then
                                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU = KUTI_SU
                                    End If
                                
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SAI_SU = SAI_SU
                                
                                End If
                            End If
                        Next j
                        
                        
                        
                        '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        
                        '------------------------------------   �o�ח\��̏���
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                'ID��
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
            
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
                                        Inspe_Proc_DEN = False
                                        GoTo Abort_Tran
                                     Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_DEN = False
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
                                        Inspe_Proc_DEN = False
                                        GoTo Abort_Tran
                                
                                    Case Else
                                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                        Inspe_Proc_DEN = SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                            Loop
                                '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                    
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F And ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")         '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                            
                                
                                
                                
                                
                                
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                
                                    'Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F)
                                    If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) Then
                                        If CInt(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) > 0 Then
                                        Else
                                            Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "0000"))
                                        End If
                                    Else
'''''''                                        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "00.00"))
                                        
                                        
                                        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "00.00"))
                                                                            
                                    End If
                                    Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN, Format(SAI_SU, "00.00"))
                                                    
                    Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN_SAV, Format(SAI_SU, "00.00"))
                                                    
                                                    
                                Else
                                                
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                        
                                        'Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F)
                                        If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) Then
                                            If CInt(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) > 0 Then
                                            Else
                                                Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "0000"))
                                            End If
                                        Else
                                            Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "0000"))
                                                                                
                                        End If
                                    End If
                                    Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN, Format(SAI_SU, "00.00"))
                                                
                                
                    Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN_SAV, Format(SAI_SU, "00.00"))
                                
                                
                                
                                End If
                                                    
                                
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                
                                
                                
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")    '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_DEN = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                

'                        Call Syuka_KUTI_SU_Count_Proc(TOTAL_KUTI_SU)

        
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                   
                        '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F And ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                
                                Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                    'ID��
                                Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
                
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
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                         Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                            
                                Loop
                                
                                
                                
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "") '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                                
                                Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, Format(KUTI_SU, "0000"))
                                Call UniCode_Conv(Y_SYU_HREC.SAI_SU, Format(SAI_SU, "00.00"))
                                                    
                                                    
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                                    
                                                    
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_DEN = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_DEN = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                            '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                

''' �i�ԒP�ʂł̊ۂߏ���
                        
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
''' �i�ԒP�ʂł̊ۂߏ���

                            
'                        PRINT_OFF = False
'
'                        If KONPOU_OFF = KONPOU_OFF_SUMI Then
'
'                            PRINT_OFF = True
'
'                        End If

                        If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 0 Then
                            ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = KUTI_SU
                        End If


                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                            Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                            Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) Then
                            If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                PRINT_OFF = True
                            End If
                        End If



                        PRINT_MAISU = KUTI_SU
                                                
                                                
                        Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1

                        PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU


                        If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD)) Then
                            PRINT_MAISU = PRINT_MAISU - 1
                            If PRINT_MAISU < 1 Then
                                
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                    
                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                        
                        
                                                        '�}�ԍX�V   2010.02.15
                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                
                                                GoTo Abort_Tran
                                            End If
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                        
                                        End If
                                    Else
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                            
                        
                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                    
                                                    
                                                    GoTo Abort_Tran
                                                
                                                End If
                                            
                                            
                                            End If
                                                                
                                        End If
                                    End If
                                Next j
                                
                                
                                
                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                
                                
                                
                                
                                PRINT_OFF = True
                            Else
'                                PRINT_MAISU = PRINT_MAISU - 1
                                Start_Page_No = Start_Page_No + 1
                            
                            End If
                        End If

''' �׎D���u
                        If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Or PRINT_OFF Then
                            ID_KANRI_TBL(ING_No).LABEL_ON = False
                        Else
                                
                            Y_SYU_CNT = 0
                            SYUKA_QTY = 0
                            
                            PRINT_OFF = False
                            
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            
                                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    
                                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                    
                                        Y_SYU_CNT = Y_SYU_CNT + 1
                                        SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                    
                                    
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                        
                                            PRINT_OFF = True
                                        Else
                                    
                                        End If
                                    End If
                                End If
                            Next j
                                
                            If Not PRINT_OFF Then
                                
                                If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                End If
                            
                            
                            
                            
                                '�}�ԍX�V   2010.02.15
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                        
                        
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" And Start_Page_No = 2 Then
                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU, "000")) Then
                                                    
                                                    GoTo Abort_Tran
                                                
                                                
                                                End If
                        
                                            Else
                        
                        
                                                        '�}�ԍX�V   2010.02.15
                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                    
                                                    GoTo Abort_Tran
                                                
                                                
                                                End If
                                                        
                                            End If
                                                        
                                                        '�}�ԍX�V   2010.02.15
                        
                        
                        
                        
                        
                                        End If
                                    Else
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                            
                        
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" And Start_Page_No = 2 Then
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    End If
                                                Else
                                                
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                                        
                                                
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                                '�}�ԍX�V   2010.02.15
                            
                            
                            
'                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No + PRINT_MAISU
                                
                                If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                    ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
                                Else
                                    ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
                                End If
                                
'                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'
'
'                                            If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                GoTo Abort_Tran
'                                            End If
'
'                                        End If
'                                    End If
'                                Next j
                            
                                ID_KANRI_TBL(ING_No).LABEL_STEP = 9
                                
                                '�f�[�^���M
                                                            
                                ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                            
                                ID_KANRI_TBL(ING_No).LABEL_ON = True
                            
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Print_Text_Proc(FileName)
                                '-----------------------------------------------�{�f�B
                                Call Wel_Hin_No_Req_Text_Proc(Sumi_CNT, Y_SYU_CNT)
                            
                                Sendbuf = Text_Create_Proc()
                            
                                Inspe_Proc_DEN = False
                                Exit Function
                            
                            End If

''' �׎D���u
                        End If

                        KENPIN_END = True
                        
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                            
                                                            
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                    KENPIN_END = False
                                    Exit For
                                End If
                            End If
                        Next j



                        Select Case KENPIN_END
                    
                            Case True
    
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
                                            
                                            
                                            
                                            
                                            
                    
                    
                            Case Else
                    
    ''''''''''''''''''''''''''''''''''''''''''''''
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                    
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Text_Proc
                                '-----------------------------------------------�P�s��
                                Call Wel_DETAIL_0_Text_Proc
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                
                                '-----------------------------------------------�R�s��
                                Call Wel_HIN_NO_Req_Text_3_Proc
                                '-----------------------------------------------�S�s��
                                Call Wel_HIN_NO_Req_Text_4_Proc
                                
                                '-----------------------------------------------�R�s��
'                                                                                        'BOX����
'                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                                                                                        '�\�����e
'
'                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'
'
'                                                                                        '���l�����\��
'                                Send_Text.Box_Type(2).INIT = ""
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
'                                                                                        '�����J�[�\���ʒu
'                                Send_Text.Box_Type(2).Start_Pos = ""
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
'                                                                                        '���͌���
'                                Send_Text.Box_Type(2).Max_Size = "00"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
'
'                                Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
'                                '-----------------------------------------------�S�s��
'                                Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                                                                                        '�\�����e
'                                                                                    '�\�����e
'                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                                                                                        '���l�����\��
'                                Send_Text.Box_Type(3).INIT = ""
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
'                                                                                        '�����J�[�\���ʒu
'                                Send_Text.Box_Type(3).Start_Pos = "01"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
'                                                                                        '���͌���
'                                Send_Text.Box_Type(3).Max_Size = "13"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
'
'                                Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                '-----------------------------------------------�T�s��
                                Call Wel_Clear_Text_Proc
        
                                Sendbuf = Text_Create_Proc()
                    
                        End Select
                    
                    End Select
                
            Next i
        

        Case Step_PRINT_RES        '����I��
    
    
            '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
            Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
    
            '-----------------------------------------------�w�b�_�[
            Call Wel_Head_Text_Proc
    
            Select Case ID_KANRI_TBL(ING_No).LABEL_STEP
                Case 1      '�i�Ԃ�
                                                                            
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                                                                            
                    '-----------------------------------------------�w�b�_�[ 02.24
                    Call Wel_Head_Text_Proc
                                                                            
                    '-----------------------------------------------�P�s��
                    Call Wel_DETAIL_0_Text_Proc
                    '-----------------------------------------------�Q�s��
                    Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                    
                    
                    
                    
                    '-----------------------------------------------�R�s��
                    Call Wel_HIN_NO_Req_Text_3_Proc
                    '-----------------------------------------------�S�s��
                    Call Wel_HIN_NO_Req_Text_4_Proc
                    
                    
                    '-----------------------------------------------�R�s��
                                                                            'BOX����
'                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                                                                            '�\�����e
'
'                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'
'
'                                                                            '���l�����\��
'                    Send_Text.Box_Type(2).INIT = ""
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
'                                                                            '�����J�[�\���ʒu
'                    Send_Text.Box_Type(2).Start_Pos = ""
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
'                                                                            '���͌���
'                    Send_Text.Box_Type(2).Max_Size = "00"
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
'
'                    Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
'                    '-----------------------------------------------�S�s��
'                    Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                                                                            '�\�����e
'                                                                        '�\�����e
'                    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                                                                            '���l�����\��
'                    Send_Text.Box_Type(3).INIT = ""
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
'                                                                            '�����J�[�\���ʒu
'                    Send_Text.Box_Type(3).Start_Pos = "01"
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
'                                                                            '���͌���
'                    Send_Text.Box_Type(3).Max_Size = "13"
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
'
'                    Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                    '-----------------------------------------------�T�s��
                    Call Wel_Clear_Text_Proc

                    Sendbuf = Text_Create_Proc()
                
                Case 2      '������
                
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                    '-----------------------------------------------�w�b�_�[
                    Call Wel_Head_Text_Proc
                    
                    '-----------------------------------------------�P�s��
                    Call Wel_DETAIL_0_Text_Proc
                    '-----------------------------------------------�Q�s��
                    Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                    
                    '-----------------------------------------------�R�s��
                                                                            'BOX����
                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                            '�\�����e
                                                                            
                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                            
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
'�����@�S
                    wkKONPO_F = ""
                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                    
                        If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                        
                            wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                            Exit For
                        End If
                    Next j
                    
                    If wkKONPO_F = "1" Then
                                            
                        If Inspection_Input Then
                            KUTI_SU_INPUT_F = False
                        Else
                            KUTI_SU_INPUT_F = True
                        End If
                    
                    
                        TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                        TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                
                    Else
                        TOTAL_KUTI_SU = 1
                        KUTI_SU_INPUT_F = True
                        TOTAL_SAI_SU = 0#
                    End If
                                
                                
                                
                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                    End If
                                
                                
                    If KUTI_SU_INPUT_F Then
                    
                        Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                        ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                        ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                    
                    
                    Else
                        Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                    
                    
                    
                        If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
                    
                    
                    
                    
                    
                    
                        ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                        ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                    
                    
                    
                    End If
                    
                    Sendbuf = Text_Create_Proc()
                
                Case 9
        
                    Select Case (Y_SYU_CNT - Sumi_CNT)
                
                        Case 0
                
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
            
            
                        Case Else
                        
        ''''''''''''''''''''''''''''''''''''''''''''''
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                
                            '-----------------------------------------------�w�b�_�[ 02.24
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                            Call Wel_HIN_NO_Req_Text_3_Proc
                            '-----------------------------------------------�S�s��
                            Call Wel_HIN_NO_Req_Text_4_Proc
                            '-----------------------------------------------�T�s��
                            Call Wel_Clear_Text_Proc
        
                            Sendbuf = Text_Create_Proc()
        
            
                End Select
            End Select
    
    End Select
                    
                    

    Inspe_Proc_DEN = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function

Public Function Inspe_Proc_E_BAG(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")
'-------------------------------------------------------
'
'   �w���i�����i��z�`�[�ǂݍ��� ���o�b�����j���ޯ��Ή��x
'
'       2010.01.21
'-------------------------------------------------------
Dim sts             As Integer

'2010.12.07
'Dim Hinban          As String * 13
Dim Hinban          As String * 20
'2010.12.07

Dim SYUKA_QTY       As Long
Dim MTS_CODE        As String * 8

'2010.12.07
'Dim HIN_NO          As String * 13
Dim HIN_NO          As String * 20
'2010.12.07

Dim OKURI_NO        As String
Dim KUTI_SU         As Integer
Dim UNSOU_KAISHA    As String
Dim SURYO           As String


Dim Y_SYU_TBL()     As KEN_DEN_TBL_Tag


Dim KAN_FLG         As String * 1

Dim i               As Integer
Dim j               As Integer
Dim k               As Integer

Dim DEN_ID_LOOP     As Integer
Dim DEN_ID_JGYOBU   As String * 1

Dim Y_SYU_CNT       As Integer
Dim Sumi_CNT        As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 7
Dim KAN_KBN         As String * 1


Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2


Dim KENPIN_END      As Boolean

Dim OKURI_SAKI      As String

Dim CANCEL_F        As Boolean

Dim FAST_F          As Boolean
Dim Found_F         As Boolean


Dim OKURI_NO_F      As Boolean

Dim FUKUYAMA_CHK_F  As Boolean

    Inspe_Proc_E_BAG = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�`�[�h�c�j
        
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No  '�`�[�h�c
                                
                        '�e�`���j�d�d�o
                        ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                         
                        Erase Y_SYU_TBL
                                        
                        sts = Y_Syuka_H_Chek_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                                MTS_CODE, _
                                                Y_SYU_CNT, _
                                                Sumi_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                Y_SYU_TBL(), _
                                                OKURI_NO, _
                                                UNSOU_KAISHA, _
                                                OKURI_SAKI, _
                                                Found_F)
                        
                        
                        
                        
                        '���[���Ŏg�p�� 2011.04.07
                        If sts = SYS_CANCEL Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\��g�p��", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        End If
                        '���[���Ŏg�p�� 2011.04.07
                        
                        
                        
                        
                        '��ݾٓ`�[������
                        
                        If Found_F Then
                        
                            CANCEL_F = True
                                                     
                            For j = 0 To UBound(Y_SYU_TBL)
                            
                                If Not Y_SYU_TBL(j).CANCEL_F Then
                                    CANCEL_F = False
                                    Exit For
                                End If
                            
                            Next j
                                                     
                                                     
                            If CANCEL_F Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�L�����Z���`�[�ł��B", "", "")         '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�L�����Z���`�[�ł��B", "", "")     '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_E_BAG = False
                                Exit Function
                            End If
                        End If
                        
                        
                        
                        If Y_SYU_CNT = 0 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")                     '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\�薳��", "", "")                 '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        End If
                                                 
                                                 
                                                 
                                                 
                                                 
                                                 
                        If Sumi_CNT = Y_SYU_CNT Then
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "���i�����ρI", "", "")                     '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "���i�����ρI", "", "")                 '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        
                        End If
                                                 
                                                             
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                                                 
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        
                        Erase ID_KANRI_TBL(ING_No).KEN_DEN_TBL
                        For j = 0 To UBound(Y_SYU_TBL)
                        
                            ReDim Preserve ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j)
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO = Y_SYU_TBL(j).SEQ_NO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO = Y_SYU_TBL(j).HIN_NO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO = Y_SYU_TBL(j).SURYO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI = Y_SYU_TBL(j).SUMI
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F = Y_SYU_TBL(j).CANCEL_F
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KAN_KBN = Y_SYU_TBL(j).KAN_KBN      '2007.05.14
                        
                        
                        Next j
                        
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO = OKURI_NO
                        
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI = OKURI_SAKI
                        
                        '�^�����
                        ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA = UNSOU_KAISHA
                        
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_NO) & _
                                                                "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_NO) & _
                                                                "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
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
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, OKURI_SAKI)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, OKURI_SAKI)
                                                                                
                                                                                
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
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '�\�����e
                        
                                            
                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = "" Then
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
                        Else
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
                        End If
                                                                                
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
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i����󇂁j
                
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    '�����
                    Case Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO, _
                                LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
                    
                        
                        If Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) = LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) Then
                            
                            If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) > Len(LCD_OKURI_NO_S) Then
                                If Left(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), Len(LCD_OKURI_NO_S)) = LCD_OKURI_NO_S Then
                                    OKURI_NO = Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
                                Else
                                    OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                                End If
                            Else
                                OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                            End If
                        Else
                            OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        End If
                        
                        
                        If Trim(OKURI_NO) = Trim(KEN_CHARTER_CD) Or Trim(OKURI_NO) = Trim(KEN_AKABOU_CD) Or Trim(OKURI_NO) = Trim(KEN_LOGISTIC_CD) Then
                        
                        '�`���[�^�[��   2010.01.21
                        
                        Else
                        
'2009.10.14                         If Len(OKURI_NO) < 11 Or Len(OKURI_NO) > 13 Then
'                        If Len(OKURI_NO) < 10 Or Len(OKURI_NO) > 13 Then
'                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
'
'                            Sendbuf = Text_Create_Proc()
'                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                            Inspe_Proc_E_BAG = False
'                            Exit Function
'                        End If
                    
                            If Not IsNumeric(OKURI_NO) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "����󇂃G���[", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_E_BAG = False
                                Exit Function
                            End If
                        
                            If OKURI_NO_CHECK_PROC(OKURI_NO, OKURI_NO_F, FUKUYAMA_CHK_F) Then
                            End If
                            
                            
                            
                            
                            
                            If Not OKURI_NO_F Then
                            
                        
                        
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "����󇂃G���[", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_E_BAG = False
                                Exit Function
                                                
                            End If
                        
                        
                            '2009.04.28
                            If FUKUYAMA_CHK_F Then
                            
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "���R �����޼ޯĴװ", "", "")         '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "���R �����޼ޯĴװ", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_E_BAG = False
                                Exit Function
                            
                            End If
                            '2009.04.28
                        
                        
                        
    '                        Select Case Len(Trim(OKURI_NO))
    '
    '                            Case FUKUYAMA_Length
    '                            Case SEIBU_Length
    '                            Case KURUME_Length
    '
    '                                For k = 0 To UBound(KURUME_CODE)
    '
    '                                    If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
    '                                        Exit For
    '                                    End If
    '                                Next k
    '
    '                                If k > UBound(KURUME_CODE) Then
    '
    '
    '                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
    '
    '                                    Sendbuf = Text_Create_Proc()
    '                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
    '                                    Inspe_Proc_E_BAG = False
    '                                    Exit Function
    '
    '                                End If
    '
    '                            Case SAGAWA_Length, YAMATO_Length
    '
    '                                For k = 0 To UBound(KURUME_CODE)
    '
    '                                    If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
    '                                        Exit For
    '                                    End If
    '                                Next k
    '
    '                                If k > UBound(SAGAWA_CODE) Then
    '
    '                                    For k = 0 To UBound(YAMATO_CODE)
    '
    '                                        If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
    '                                            Exit For
    '                                        End If
    '
    '                                    Next k
    '
    '                                    If k > UBound(YAMATO_CODE) Then
    '
    '                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
    '
    '                                        Sendbuf = Text_Create_Proc()
    '                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
    '                                        Inspe_Proc_E_BAG = False
    '                                        Exit Function
    '
    '
    '
    '                                    End If
    '
    '                                End If
    '
    '
    '                        End Select
                        End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO = OKURI_NO

                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                
                
                        '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                            
                                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        
                        '------------------------------------   �o�ח\��̏���
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                'ID��
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
            
                            Do
                            
                                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrKeyNotFound
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_E_BAG = False
                                        GoTo Abort_Tran
                                     Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_E_BAG = False
                                        GoTo Abort_Tran
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                        
                            Loop
    
                            '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                    
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "") '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_E_BAG = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2107.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_E_BAG = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                            
                                Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, OKURI_NO)
                                Call OKURI_NO_SET_PROC(OKURI_NO)
                                            
'                                '�^����Еϊ�
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, 3) = UNSOU_KAISHA_CODE Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU_KAISHA_NAME)
'                                End If
'                                '�V�^����Еϊ� 2007.01.09
'
'                                If KURUME_F Then        '�v����
'                                    For k = 1 To UBound(KURUME)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(KURUME(k))) = KURUME(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
'
'                                If FUKUYAMA_F Then      '���R
'                                    For k = 1 To UBound(FUKUYAMA)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(FUKUYAMA(k))) = FUKUYAMA(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
'
'                                If SAGAWA_F Then        '����
'                                    For k = 1 To UBound(SAGAWA)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(SAGAWA(k))) = SAGAWA(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
                                                    
                                                    
                                                    
'                                Select Case Len(Trim(OKURI_NO))
'
'                                    Case FUKUYAMA_Length
'                                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA_Name)
'                                    Case SEIBU_Length
'                                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEIBU_Name)
'
'                                    Case KURUME_Length
'
'                                        For k = 0 To UBound(KURUME_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
'                                                Exit For
'                                            End If
'                                        Next k
'
'                                        If k > UBound(KURUME_CODE) Then
'                                        Else
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME_Name)
'                                        End If
'
'                                    Case SAGAWA_Length, YAMATO_Length
'
'                                        For k = 0 To UBound(KURUME_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
'                                                Exit For
'                                            End If
'                                        Next k
'
'                                        If k > UBound(SAGAWA_CODE) Then
'
'                                            For k = 0 To UBound(YAMATO_CODE)
'
'                                                If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
'                                                    Exit For
'                                                End If
'
'                                            Next k
'
'                                            If k > UBound(YAMATO_CODE) Then
'                                            Else
'
'                                                Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, YAMATO_Name)
'                                            End If
'
'                                        Else
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA_Name)
'                                        End If
'
'
'                                End Select
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                '���� 2010.02.19
'                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_CHARTER_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "����")
'                                End If
'
'                                '�ԖX 2010.02.19
'                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_AKABOU_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "�ԖX")
'                                End If
                                                    
                                                    
                                                    
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                                    
                                                    
                                                    
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_E_BAG = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_E_BAG = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
                                            '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                
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
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                                                                                
                                                                                
                        Sumi_CNT = 0
                        Y_SYU_CNT = 0
                                                                                

''' �i�ԒP�ʂł̊ۂߏ���
                        FAST_F = True
                                                                                                    
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        
                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            Else
                                If FAST_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                        Sumi_CNT = Sumi_CNT + 1
                                    End If
                                    FAST_F = False
                                
                                Else
                                    For k = 0 To j - 1
                                        If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(k).HIN_NO) Then
                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(k).CANCEL_F Then
                                                Exit For
                                            End If
                                        End If
                                    Next k
                            
                                    If k > j - 1 Then
                            
                                        Y_SYU_CNT = Y_SYU_CNT + 1
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                            Sumi_CNT = Sumi_CNT + 1
                                        End If
                                    End If
                                End If
                            End If
                        Next j
''' �i�ԒP�ʂł̊ۂߏ���
                                                                                
                                                                                
                                                                                
                                                                                
                                                                                
                                                                                
                                                                                
                                                                                
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                                "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                                "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
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
                                                                                
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                                                                                
                                                                                
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
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                '�\�����e
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
'                        Send_Text.Box_Type(3).Max_Size = "13"
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                        Send_Text.Box_Type(3).Max_Size = "20"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
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
        
        
        
        
        
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
            
                
                '2010.12.07
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                Select Case Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), 2)
                '2010.12.07
                    
                    Case LCD_Hinban     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")      '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")  '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_E_BAG = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                
                        End Select
                        
                        '�Y���i�ԗL��������
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                Exit For
                            End If
                        Next j
                        
                        If j > UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")      '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")  '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        End If
                        
                        
                        '��ݾق�����
                        CANCEL_F = True
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                    CANCEL_F = False
                                    Exit For
                                End If
                            
                            End If
                        Next j
                        
                        If CANCEL_F Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�L�����Z���i�Ԃł��B", "")        '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�L�����Z���i�Ԃł��B", "")    '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        
                        
                        End If
                        
                        
                        
                                
                        '���i�ς݂�����
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                Exit For
                            End If
                        Next j
                
                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���i�ς݁I", "")      '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���i�ς݁I", "")  '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        End If
                
                        '���o�ɂ�����   2007.05.14
                        If Inspection_Flg = 0 Then
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KAN_KBN <> KAN_KBN_FIN Then
                                        Exit For
                                    End If
                                End If
                            Next j
                                            
                            If j <= UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���o�ɕ��L��I�I", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���o�ɕ��L��I�I", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_E_BAG = False
                                Exit Function
                            End If
                        End If
                        '���o�ɂ�����   2007.05.14
                
                
                
                        ID_KANRI_TBL(ING_No).KEN_HINBAN = Hinban
                
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
                
                        Send_Text.buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        '>>>>>>>>>>>    2107.09.22
                        'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                        '>>>>>>>>>>>    2107.09.22
                                                                                
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
                                                                                
                                                                                
                        Sumi_CNT = 0
                        Y_SYU_CNT = 0
                                                                                

''' �i�ԒP�ʂł̊ۂߏ���
                        FAST_F = True
                                                                                                    
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        
                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            Else
                                If FAST_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                        Sumi_CNT = Sumi_CNT + 1
                                    End If
                                    FAST_F = False
                                
                                Else
                                    For k = 0 To j - 1
                                        If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(k).HIN_NO) Then
                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(k).CANCEL_F Then
                                                Exit For
                                            End If
                                        End If
                                    Next k
                            
                                    If k > j - 1 Then
                            
                                        Y_SYU_CNT = Y_SYU_CNT + 1
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                            Sumi_CNT = Sumi_CNT + 1
                                        End If
                                    End If
                                End If
                            End If
                        Next j
''' �i�ԒP�ʂł̊ۂߏ���


                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                                "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                                "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
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
                        
                        
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Hinban)
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
                        If Inspection_QTY = 1 Then

                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        Else
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM                             '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM        '2007.04.21
                        End If
                        
                        Y_SYU_CNT = 0
                        SYUKA_QTY = 0
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                End If
                            End If
                        Next j
                                                                                
                                                                                
                                                                                '�\�����e
                        
                        If Y_SYU_CNT < 2 Then
                        

                            If Inspection_QTY = 1 Then

                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide))
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka_Su1)                         '2007.04.21
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka_Su1)    '2007.04.21
                            End If
                                                                                
                                                                                
                        Else
                        
                            If Inspection_QTY = 1 Then
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide) & "*")
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide) & "*")
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka_Su2)                       '2007.04.21
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka_Su2)  '2007.04.21
                            End If
                        
                        End If
                                                                                
                                                                                '���l�����\��
                        If Inspection_QTY = 1 Then
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                        Else
                            Send_Text.Box_Type(4).INIT = ""                                                     '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""                                '2007.04.21
                        End If
                                                                                
                                                                                
                                                                                '�����J�[�\���ʒu
                        If Inspection_QTY = 1 Then

                            Send_Text.Box_Type(4).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                        Else
                            Send_Text.Box_Type(4).Start_Pos = "10"                                          '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "10"                     '2007.04.21
                        End If
                                                                                
                                                                                '���͌���
                        If Inspection_QTY = 1 Then
                            Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                        Else
                            Send_Text.Box_Type(4).Max_Size = "07"                                           '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "07"                      '2007.04.21
                        End If
                                                                                
                                                                                
                        '2009.04.15
                        If SYUKA_QTY > 1 Then
                            Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                                                                
                        End If
                                                                                
                                                                                
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        
                        
                        
                        Sendbuf = Text_Create_Proc()
                
                
                
                End Select
            
            Next i
        
                
        Case Step_Sagyo4_RES        '�S��ڂ̎�M�i���i���@��M�j
            
            
            
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case RTrim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), _
                                Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Syuka_Su1, LCD_Syuka_Su2   '�o�א�(���i��)
                        SURYO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        
                        If Not IsNumeric(SURYO) Then
                        
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        
                        End If
                
                        Y_SYU_CNT = 0
                        SYUKA_QTY = 0
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                End If
                            End If
                        Next j
                
                
                
                        If CLng(SURYO) <> SYUKA_QTY Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�א��G���[", "", "")        '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�א��G���[", "", "")   '2107.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        End If
                
                End Select
            
            Next i
            
            
            
            
            
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                            
                                            
            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                    
                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                        
                        
                        
                        
                        '------------------------------------   �o�ח\��̏���
                        Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                            'ID��
                        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
        
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
                                    Inspe_Proc_E_BAG = False
                                    GoTo Abort_Tran
                                 Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_E_BAG = False
                                    GoTo Abort_Tran
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                    
                        Loop
        
                    '''���s�̎g�p���p�����邽��
                    '''Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    '''Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    
                    
                    
                                
                    
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                                    
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
                                    Inspe_Proc_E_BAG = False
                                    GoTo Abort_Tran
                            
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                    Inspe_Proc_E_BAG = SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                        Loop
                        '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                        
                        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)    'ID��
        
                        Do
                        
                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrKeyNotFound
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_E_BAG = False
                                    GoTo Abort_Tran
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_E_BAG = False
                                    GoTo Abort_Tran
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                    
                        Loop
                                            
                                            
                        Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, Format(Now, "YYYYMMDDHHMMSS"))
                        Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
                                            
                        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
'                        '�^����Еϊ�
'                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, 3) = UNSOU_KAISHA_CODE Then
'                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU_KAISHA_NAME)
'                        End If
'                        '�V�^����Еϊ� 2007.01.09
'
'                        If KURUME_F Then        '�v����
'                            For k = 1 To UBound(KURUME)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(KURUME(k))) = KURUME(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
'
'                        If FUKUYAMA_F Then      '���R
'                            For k = 1 To UBound(FUKUYAMA)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(FUKUYAMA(k))) = FUKUYAMA(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
'
'                        If SAGAWA_F Then        '����
'                            For k = 1 To UBound(SAGAWA)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(SAGAWA(k))) = SAGAWA(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
                                            
                                            
                                            
                                            
                                            
                        Select Case Len(Trim(OKURI_NO))
                        
                            Case FUKUYAMA_Length
                                Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA_Name)
                            Case SEIBU_Length
                                Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEIBU_Name)
                        
                            Case KURUME_Length
                        
                                For k = 0 To UBound(KURUME_CODE)
                                
                                    If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
                                        Exit For
                                    End If
                                Next k
                        
                                If k > UBound(KURUME_CODE) Then
                                Else
                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME_Name)
                                End If
                        
                            Case SAGAWA_Length, YAMATO_Length
                        
                                For k = 0 To UBound(KURUME_CODE)
                                
                                    If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
                                        Exit For
                                    End If
                                Next k
                        
                                If k > UBound(SAGAWA_CODE) Then
                                
                                    For k = 0 To UBound(YAMATO_CODE)
                                    
                                        If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
                                            Exit For
                                        End If
                                    
                                    Next k
                                 
                                    If k > UBound(YAMATO_CODE) Then
                                    Else
                                
                                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, YAMATO_Name)
                                    End If
                                
                                Else
                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA_Name)
                                End If
                        
                        
                        End Select
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                        '���� 2010.02.19
                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_CHARTER_CD) Then
                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "����")
                        End If
                        
                        '�ԖX 2010.02.19
                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_AKABOU_CD) Then
                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "�ԖX")
                        End If
                                                                    
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                        Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                        Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                            
                                            
                                            
                                            
                                            
                                                    '�o�ח\��(νĲҰ��)������
                        Do
                            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_E_BAG = False
                                    GoTo Abort_Tran
                            
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                    Inspe_Proc_E_BAG = SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                        Loop
                                            
                                            
                        '------------------------------------   �݌Ɉړ������̏���
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                            MENU_NO = ""
                        End If
                                            
                        '�����o�ׂ͂̈̓ǂݍ���
                        Call UniCode_Conv(K0_MTS.MUKE_CODE, ID_KANRI_TBL(ING_No).MTS_CODE)
                        Call UniCode_Conv(K0_MTS.SS_CODE, "")
                        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(MTSREC.MUKE_DNAME, "")
                                Call UniCode_Conv(MTSREC.MUKE_NAME, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "������}�X�^", 0)
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                                            
                                            
                        sts = IDOREKI_OUTPUT_PROC("", _
                                                    "", _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    ID_KANRI_TBL(ING_No).KEN_HINBAN, _
                                                    "", _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    0, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    FILE_RETRY, _
                                                    CYU_KBN_SPO, _
                                                    Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) & " �����:" & Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)), _
                                                    , , , , MENU_NO, _
                                                    ID_KANRI_TBL(ING_No).MTS_CODE, _
                                                    "", _
                                                    ID_KANRI_TBL(ING_No).ID_NO & "-" & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO, , , , 1)
                        Select Case sts
                            Case False      '����I��
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc_E_BAG = SYS_ERR
                                GoTo Abort_Tran
                        End Select
                                            
                        '���i�ρI�I
                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI = True
                                            
                        '�^�����
                        ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA = StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)
                                        
                    End If
                End If
            
            Next j
                                
                                
                                
                                
                                
            '���۸ޏo��    2009.04.17
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
                                                     ID_KANRI_TBL(ING_No).KEN_HINBAN, , , , , _
                                                     ID_KANRI_TBL(ING_No).ID_NO) Then
                    Inspe_Proc_E_BAG = SYS_ERR
                    Exit Function
                End If
            End If
            '���۸ޏo��    2009.04.17
                                
                                
                                
                                
                                
                                '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
        
                    
            KENPIN_END = True
            
            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                
                                                
                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                        KENPIN_END = False
                        Exit For
                    End If
                End If
            Next j
        
            Sumi_CNT = 0
            Y_SYU_CNT = 0
                                                                                
''' �i�ԒP�ʂł̊ۂߏ���
            FAST_F = True
                                                                                        
            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
            
            
                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                Else
                    If FAST_F Then
                    
                        Y_SYU_CNT = Y_SYU_CNT + 1
                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                            Sumi_CNT = Sumi_CNT + 1
                        End If
                        FAST_F = False
                    
                    Else
                        For k = 0 To j - 1
                            If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(k).HIN_NO) Then
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(k).CANCEL_F Then
                                    Exit For
                                End If
                            End If
                        Next k
                
                        If k > j - 1 Then
                
                            Y_SYU_CNT = Y_SYU_CNT + 1
                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                Sumi_CNT = Sumi_CNT + 1
                            End If
                        End If
                    End If
                End If
            Next j
''' �i�ԒP�ʂł̊ۂߏ���
                                                                                
                                                                                
                                                                                
        
        
        
            Select Case KENPIN_END
                Case False
                    '���i�c����A�i�Ԃ�
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
                    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                            "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                            "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
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
                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                            
                                                                            
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
                    '2010.12.07
'                    Send_Text.Box_Type(3).Max_Size = "13"
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                    Send_Text.Box_Type(3).Max_Size = "20"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                    '2010.12.07
                                                                            
                    Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
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
                    Send_Text.Box_Type(4).Start_Pos = ""                    '���l�͂T���Œ�
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                                                                            '���͌���
                     Send_Text.Box_Type(4).Max_Size = "00"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                                                                            
                    Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                    Sendbuf = Text_Create_Proc()
                Case True
                    '���i�����A������
            
            
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                    
                    
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
                    Send_Text.Box_Type(0).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                            '�\�����e
                    '>>>>>>>>>>>>>> 2017.09.22
                    'Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                    'Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                            
                    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME)
                    '>>>>>>>>>>>>>> 2017.09.22
                                                                            
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
                    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                            "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, LCD_ID_No & ":" & Trim(ID_KANRI_TBL(ING_No).ID_NO) & _
                                                            "(" & Format(Sumi_CNT, "0") & "/" & Format(Y_SYU_CNT, "0") & ")")
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
                                                                            
                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                            
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
                    
                    
                    '-----------------------------------------------�T�s��
                    KUTI_SU = 1
                                
                    Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                    '�\�����e
                    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_KUTI_SU)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_KUTI_SU)
                                                                    '���l�����\��
                    Send_Text.Box_Type(4).INIT = Space(9) & "1"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(9) & "1"
                                                                    '�����J�[�\���ʒu
                    Send_Text.Box_Type(4).Start_Pos = "07"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "07"
                                                            '���͌���
                    Send_Text.Box_Type(4).Max_Size = "04"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "04"
                    
                    
                    
                    Sendbuf = Text_Create_Proc()
            
            
            
            
            
            
            
            End Select
            
                
                
                
                
        Case Step_Sagyo5_RES        '�T��ڂ̎�M�i�����j
                
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    
                    '����
                    Case LCD_KUTI_SU
                
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")     '2017.09.22
                
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
                            Exit Function
                        
                        End If
                
                        KUTI_SU = CInt(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If KUTI_SU < 1 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")     '2107.09.22
                
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_E_BAG = False
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
                                            
                                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        
                        '------------------------------------   �o�ח\��̏���
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                'ID��
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
            
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
                                        Inspe_Proc_E_BAG = False
                                        GoTo Abort_Tran
                                     Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_E_BAG = False
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
                                        Inspe_Proc_E_BAG = False
                                        GoTo Abort_Tran
                                
                                    Case Else
                                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                        Inspe_Proc_E_BAG = SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                            Loop
                            '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                    
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "") '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_E_BAG = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_E_BAG = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                            
                                            
                                Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, Format(KUTI_SU, "0000"))
                                            
                                                            
                                                            
                                                            
                                                            
                                Call UniCode_Conv(Y_SYU_HREC.JURYO, "0002.0")
                                Call UniCode_Conv(Y_SYU_HREC.SAI_SU, "01.00")
                                                            
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                                            
                                                            
                                                            
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_E_BAG = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_E_BAG = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
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
                
                
                
                
                
            Next i
                
                
                
                
    
    
    End Select

    Inspe_Proc_E_BAG = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function

Public Function Inspe_Proc_LOGISTIC(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����i��z�`�[�ǂݍ��� ���o�b�����j�x
'       '���W�X�e�b�N�X�Ή�
'       2010.02.25
'-------------------------------------------------------
Dim sts             As Integer

'2010.12.07
'Dim Hinban          As String * 13
Dim Hinban          As String * 20
'2010.12.07


Dim SYUKA_QTY       As Long
Dim MTS_CODE        As String * 8

'2010.12.07
'Dim HIN_NO          As String * 13
Dim HIN_NO          As String * 20
'2010.12.07

Dim OKURI_NO        As String
Dim KUTI_SU         As Long

Dim SAI_SU          As Double

Dim UNSOU_KAISHA    As String
 
Dim SYUKA_YMD       As String
Dim JYUSHO          As String
Dim BIKOU           As String

Dim SURYO           As String

Dim Y_SYU_TBL()     As KEN_DEN_TBL_Tag

Dim KAN_FLG         As String * 1

Dim i               As Integer
Dim j               As Integer
Dim k               As Integer

Dim DEN_ID_LOOP     As Integer
Dim DEN_ID_JGYOBU   As String * 1

Dim Y_SYU_CNT       As Integer
Dim Sumi_CNT        As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 7
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

Dim KENPIN_END      As Boolean

Dim OKURI_SAKI      As String

Dim CANCEL_F        As Boolean

Dim FAST_F          As Boolean
Dim Found_F         As Boolean

'2010.01.21
Dim KONPOU_ON           As Integer


Dim KONPOU_ON_SUMI      As Integer
Dim KONPOU_OFF          As Integer
Dim KONPOU_OFF_SUMI     As Integer
Dim PRINT_OFF           As Boolean
Dim Start_Page_No       As Long
Dim PRINT_TOTAL_SU      As Long
Dim PRINT_MAISU         As Long
Dim FileName            As String
Dim ID_SEQ              As Integer
Dim DISP_SAI_SU         As Double

Dim wkKUTI_SU           As String
Dim wkKONPO_F           As String * 1

Dim TOTAL_KUTI_SU       As Integer
Dim TOTAL_SAI_SU        As Double
Dim MUKE_NAME           As String
Dim OKURI_NO_MAX        As Integer
Dim KUTI_SU_INPUT_F     As Boolean

Dim KEN_TEL_NO          As String * 20

Dim KEN_TYAKUTEN        As String * 3       '2017.04.06

Dim LOGIXTICS_F         As Boolean

Dim OKURI_NO_F          As Boolean
'2010.01.21
Dim FUKUYAMA_CHK_F      As Boolean

    Inspe_Proc_LOGISTIC = True

    LOGIXTICS_F = True
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�`�[�h�c�j
        
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No  '�`�[�h�c
                                
                        '�e�`���j�d�d�o
                        ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                         
                        Erase Y_SYU_TBL
                                        
                        sts = Y_Syuka_H_Chek_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                                MTS_CODE, _
                                                Y_SYU_CNT, _
                                                Sumi_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                Y_SYU_TBL(), _
                                                OKURI_NO, _
                                                UNSOU_KAISHA, _
                                                OKURI_SAKI, _
                                                Found_F, _
                                                SYUKA_YMD, _
                                                JYUSHO, _
                                                BIKOU, _
                                                Start_Page_No, _
                                                KUTI_SU, _
                                                MUKE_NAME, _
                                                OKURI_NO_MAX, , , _
                                                KEN_TEL_NO, , , _
                                                KEN_TYAKUTEN)
                        
                        '���[���Ŏg�p�� 2011.04.07
                        If sts = SYS_CANCEL Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\��g�p��", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        End If
                        '���[���Ŏg�p�� 2011.04.07
                        
                        
                        '��ݾٓ`�[������
                        
                        If Found_F Then
                        
                            CANCEL_F = True
                                                     
                            For j = 0 To UBound(Y_SYU_TBL)
                            
                                If Not Y_SYU_TBL(j).CANCEL_F Then
                                    CANCEL_F = False
                                    Exit For
                                End If
                            
                            Next j
                                                     
                            If CANCEL_F Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�L�����Z���`�[�ł��B", "", "")     '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�L�����Z���`�[�ł��B", "", "") '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            End If
                        End If
                        
                        
                        If Y_SYU_CNT = 0 Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "�o�ח\�薳��", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        End If
                                                 
                        If Sumi_CNT = Y_SYU_CNT And Start_Page_No <> 0 Then
                            
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "���i�����ρI", "", "")         '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_NO, "���i�����ρI", "", "")     '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        
                        End If
                                                             
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                                                 
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        
                        Erase ID_KANRI_TBL(ING_No).KEN_DEN_TBL
                        For j = 0 To UBound(Y_SYU_TBL)
                        
                            ReDim Preserve ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j)
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO = Y_SYU_TBL(j).SEQ_NO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO = Y_SYU_TBL(j).HIN_NO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO = Y_SYU_TBL(j).SURYO
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI = Y_SYU_TBL(j).SUMI
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F = Y_SYU_TBL(j).CANCEL_F
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KAN_KBN = Y_SYU_TBL(j).KAN_KBN      '2007.05.14
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F = Y_SYU_TBL(j).KONPOU_F        '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_CND = Y_SYU_TBL(j).KONPOU_CND    '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).TOTAL_SU = Y_SYU_TBL(j).TOTAL_SU        '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SAI_SU = Y_SYU_TBL(j).SAI_SU            '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU = Y_SYU_TBL(j).KUTI_SU          '2010.01.21
                        
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = Y_SYU_TBL(j).PRINT_SUMI    '2010.01.21
                        
                                                                                                        '2010.01.21
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).OKURI_NO_SEQ = Y_SYU_TBL(j).OKURI_NO_SEQ
                        
                        Next j
                        
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO = OKURI_NO
                        
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI = OKURI_SAKI
                        
                        '�^�����
                        ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA = UNSOU_KAISHA
                        
                        '�o�ד�
                        ID_KANRI_TBL(ING_No).KEN_SYUKA_YMD = SYUKA_YMD
                        '�Z��
                        ID_KANRI_TBL(ING_No).KEN_JYUSHO = JYUSHO
                        
                        '���l
                        ID_KANRI_TBL(ING_No).KEN_BIKOU = BIKOU
                        '������
                        ID_KANRI_TBL(ING_No).KEN_MUKE_NAME = MUKE_NAME
                        
                        '�}��
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO_MAX = OKURI_NO_MAX
                        '�d�b�ԍ�
                        ID_KANRI_TBL(ING_No).KEN_TEL_NO = KEN_TEL_NO
                        '���X�R�[�h
                        ID_KANRI_TBL(ING_No).KEN_TYAKUTEN = KEN_TYAKUTEN
                        
                        
                        
                        '���يJ�n�y�[�W��
                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                        
                        ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = KUTI_SU
                        If Label_Print_Total_Su_Proc(KUTI_SU, PRINT_TOTAL_SU) Then
                        
                    
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        
                        
                        
                        End If
                        ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = PRINT_TOTAL_SU
                        
                        
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                        
                        
                        If KONPOU_ON <> 0 Then
                            If KONPOU_ON = KONPOU_ON_SUMI Then
                                                    
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                End If
                                                    
                            End If
                        End If
                        
                        
                        
                        
                        
                        
                        
                        
                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = "" Then
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                            
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, OKURI_SAKI)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, OKURI_SAKI)
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                            If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = "" Then
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
                            End If
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
    
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                
                
                
                        Else
                
                
                
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                                                                                    'BOX����
                                                                                    
                                                                                    
                            '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                            Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                            
KONPOU_ON = KONPOU_ON - KONPOU_ON_SUMI   '2011.03.17
        
    ''' �i�ԒP�ʂł̊ۂߏ���
Select Case KONPOU_ON    '2011.03.17
                            
                            
''''''''''2011.03.17Select Case (KONPOU_ON - KONPOU_ON_SUMI)
                                Case 0
                                '�W������Ȃ�
                                
                                    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                                    If KONPOU_ON_SUMI <> 0 And KONPOU_OFF_SUMI = 0 Then
                                        If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                        
                                        
                                        
                                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                            
                                            '-----------------------------------------------�w�b�_�[
                                            Call Wel_Head_Text_Proc
                                            '-----------------------------------------------�P�s��
                                            Call Wel_DETAIL_0_Text_Proc
                                            '-----------------------------------------------�Q�s��
                                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                            '-----------------------------------------------�R�s��
                                                                                                    'BOX����
                                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                                    '�\�����e
                                                                                                    
                                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                                    
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
                                        
                                        
                                        
                                        
                                        
                                        
                                            TOTAL_KUTI_SU = 1
                                            KUTI_SU_INPUT_F = True
                                            TOTAL_SAI_SU = 0#
                                                
                                                
                                                
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                            End If
                                                
                                                
                                    
                                            Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                            ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                            ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                    
                                    
                                    
                                            Sendbuf = Text_Create_Proc()
                                        
                                            Inspe_Proc_LOGISTIC = False
                                            Exit Function
                                        
                                        
                                        End If
                                    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                                    
                                    
                                    
                                    '-----------------------------------------------�{�f�B
                                    Call Wel_Hin_No_Req_Text_Proc(Sumi_CNT, Y_SYU_CNT)
            
                                    Sendbuf = Text_Create_Proc()
                                Case Else
                                '�W�������
                            
                                    '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                                    Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                                    '-----------------------------------------------�Q�s��
                                    Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                    '-----------------------------------------------�R�s��
                                                                                            'BOX����
                                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                            '�\�����e
                                                                                            
                                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                                    Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                            '�\�����e
                                                                                        '�\�����e
                                    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                            "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                            "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                                                                            '���l�����\��
                                    Send_Text.Box_Type(3).INIT = ""
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                            '�����J�[�\���ʒu
                                    Send_Text.Box_Type(3).Start_Pos = "01"
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                            '���͌���
                                    '2010.12.07
'                                    Send_Text.Box_Type(3).Max_Size = "13"
'                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                    Send_Text.Box_Type(3).Max_Size = "13"
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                    '2010.12.07
                                                                                            
                                                                                            
                                                                                            
                                    Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                    '-----------------------------------------------�T�s��
                                    Call Wel_Clear_Text_Proc
            
                                    Sendbuf = Text_Create_Proc()
                            
                            End Select
                        End If
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
'                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
'
'                        '-----------------------------------------------�w�b�_�[
'                        Call Wel_Head_Text_Proc
'                        '-----------------------------------------------�P�s��
'                        Call Wel_DETAIL_0_Text_Proc
'                        '-----------------------------------------------�Q�s��
'                        Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, SUMI_CNT, Y_SYU_CNT)
'                        '-----------------------------------------------�R�s��
'                                                                                'BOX����
'                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                                                                                '�\�����e
'                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, OKURI_SAKI)
'                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, OKURI_SAKI)
'                                                                                '���l�����\��
'                        Send_Text.Box_Type(2).INIT = ""
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
'                                                                                '�����J�[�\���ʒu
'                        Send_Text.Box_Type(2).Start_Pos = ""
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
'                                                                                '���͌���
'                        Send_Text.Box_Type(2).Max_Size = "00"
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
'
'                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
'                        '-----------------------------------------------�S�s��
'                        Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                                                                                '�\�����e
'                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = "" Then
'                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO)
'                        Else
'                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO))
'                        End If
'                                                                                '���l�����\��
'                        Send_Text.Box_Type(3).INIT = ""
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
'                                                                                '�����J�[�\���ʒu
'                        Send_Text.Box_Type(3).Start_Pos = "01"
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
'                                                                                '���͌���
'                        Send_Text.Box_Type(3).Max_Size = "20"
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
'
'                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
'                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
'                        '-----------------------------------------------�T�s��
'                                                                                'BOX����'
'
'                        Call Wel_Clear_Text_Proc
'
'                        Sendbuf = Text_Create_Proc()
                End Select
            Next i
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i����󇂁j
                
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                    '�����
                    Case Trim(ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA) & LCD_OKURI_NO, _
                                LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
                        
                        If Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) = LCD_OKURI_NO_S & Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) Then
                            
                            If Trim(ID_KANRI_TBL(ING_No).Recv_text(i)) > Len(LCD_OKURI_NO_S) Then
                                If Left(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), Len(LCD_OKURI_NO_S)) = LCD_OKURI_NO_S Then
                                    OKURI_NO = Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
                                Else
                                    OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                                End If
                            Else
                                OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                            End If
                        Else
                            OKURI_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        End If
                        
                        If Trim(OKURI_NO) = Trim(KEN_CHARTER_CD) Or Trim(OKURI_NO) = Trim(KEN_AKABOU_CD) Or Trim(OKURI_NO) = Trim(KEN_LOGISTIC_CD) Then
                        
                        '�`���[�^�[��   2010.01.21
                        
                        Else
'2009.10.14                         If Len(OKURI_NO) < 11 Or Len(OKURI_NO) > 13 Then
'                            If Len(OKURI_NO) < 10 Or Len(OKURI_NO) > 13 Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
'
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                Inspe_Proc_LOGISTIC = False
'                                Exit Function
'                            End If
                        
                            If Not IsNumeric(OKURI_NO) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "����󇂃G���[", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            End If
                        
                        
                            If OKURI_NO_CHECK_PROC(OKURI_NO, OKURI_NO_F, FUKUYAMA_CHK_F) Then
                            End If
                            
                            
                            
                            
                            
                            If Not OKURI_NO_F Then
                            
                        
                        
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "����󇂃G���[", "", "")    '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                                                
                            End If
                        
                        
                            '2009.04.28
                            If FUKUYAMA_CHK_F Then
                            
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "���R �����޼ޯĴװ", "", "")            '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, OKURI_NO, "���R �����޼ޯĴװ", "", "")        '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            
                            End If
                            '2009.04.28
                        
                        
                        
'                            Select Case Len(Trim(OKURI_NO))
'
'                                Case FUKUYAMA_Length
'                                Case SEIBU_Length
'                                Case KURUME_Length
'
'                                    For k = 0 To UBound(KURUME_CODE)
'
'                                        If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
'                                            Exit For
'                                        End If
'                                    Next k
'
'                                    If k > UBound(KURUME_CODE) Then
'
'
'                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
'
'                                        Sendbuf = Text_Create_Proc()
'                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                        Inspe_Proc_LOGISTIC = False
'                                        Exit Function
'
'                                    End If
'
'                                Case SAGAWA_Length, YAMATO_Length
'
'                                    For k = 0 To UBound(KURUME_CODE)
'
'                                        If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
'                                            Exit For
'                                        End If
'                                    Next k
'
'                                    If k > UBound(SAGAWA_CODE) Then
'
'                                        For k = 0 To UBound(YAMATO_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
'                                                Exit For
'                                            End If
'
'                                        Next k
'
'                                        If k > UBound(YAMATO_CODE) Then
'
'                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, OKURI_NO, "����󇂃G���[", "", "")
'
'                                            Sendbuf = Text_Create_Proc()
'                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                            Inspe_Proc_LOGISTIC = False
'                                            Exit Function
'
'
'
'                                        End If
'
'                                    End If
'
'
'                            End Select
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        End If
                    
                        '�����
                        ID_KANRI_TBL(ING_No).KEN_OKURI_NO = OKURI_NO

                
                        '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        '------------------------------------   �o�ח\��̏���
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                'ID��
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
            
                            Do
                            
                                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrKeyNotFound
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_LOGISTIC = False
                                        GoTo Abort_Tran
                                     Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_LOGISTIC = False
                                        GoTo Abort_Tran
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                        
                            Loop
    
                            '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                    
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "") '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                            
                                Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, OKURI_NO)
                                '2018.12.20Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "���W�X�e�B�N�X")
                                            
                                Call OKURI_NO_SET_PROC(OKURI_NO)
                                            
                                            
'                                '�^����Еϊ�
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, 3) = UNSOU_KAISHA_CODE Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU_KAISHA_NAME)
'                                End If
'                                '�V�^����Еϊ� 2007.01.09
'
'                                If KURUME_F Then        '�v����
'                                    For k = 1 To UBound(KURUME)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(KURUME(k))) = KURUME(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
'
'                                If FUKUYAMA_F Then      '���R
'                                    For k = 1 To UBound(FUKUYAMA)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(FUKUYAMA(k))) = FUKUYAMA(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
'
'                                If SAGAWA_F Then        '����
'                                    For k = 1 To UBound(SAGAWA)
'
'                                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(SAGAWA(k))) = SAGAWA(k) Then
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA(0))
'                                            Exit For
'                                        End If
'                                    Next k
'                                End If
'
'                                '���� 2010.02.19
'                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_CHARTER_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "����")
'                                End If
'
'                                '�ԖX 2010.02.19
'                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_AKABOU_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "�ԖX")
'                                End If
'
'                                '���W�X�e�b�N 2010.02.19
'                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_LOGISTIC_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "���W�X�e�b�N")
'                                End If
                                
                                
                                
                                
                                
'                                Select Case Len(Trim(OKURI_NO))
'
'                                    Case FUKUYAMA_Length
'                                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA_Name)
'                                    Case SEIBU_Length
'                                        Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SEIBU_Name)
'
'                                    Case KURUME_Length
'
'                                        For k = 0 To UBound(KURUME_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(KURUME_CODE(k))) = KURUME_CODE(k) Then
'                                                Exit For
'                                            End If
'                                        Next k
'
'                                        If k > UBound(KURUME_CODE) Then
'                                        Else
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME_Name)
'                                        End If
'
'                                    Case SAGAWA_Length, YAMATO_Length
'
'                                        For k = 0 To UBound(KURUME_CODE)
'
'                                            If Mid(OKURI_NO, 1, Len(SAGAWA_CODE(k))) = SAGAWA_CODE(k) Then
'                                                Exit For
'                                            End If
'                                        Next k
'
'                                        If k > UBound(SAGAWA_CODE) Then
'
'                                            For k = 0 To UBound(YAMATO_CODE)
'
'                                                If Mid(OKURI_NO, 1, Len(YAMATO_CODE(k))) = YAMATO_CODE(k) Then
'                                                    Exit For
'                                                End If
'
'                                            Next k
'
'                                            If k > UBound(YAMATO_CODE) Then
'                                            Else
'
'                                                Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, YAMATO_Name)
'                                            End If
'
'                                        Else
'                                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA_Name)
'                                        End If
'
'
'                                End Select
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
'                                '���� 2010.02.19
'                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_CHARTER_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "����")
'                                End If
'
'                                '�ԖX 2010.02.19
'                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) = Trim(KEN_AKABOU_CD) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, "�ԖX")
'                                End If
                                
                                
                                
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                
                                
                                
                                                    
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_LOGISTIC = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
                                            '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        '-----------------------------------------------�w�b�_�[
                        Call Wel_Head_Text_Proc
                        
                        '-----------------------------------------------�P�s��
                        Call Wel_DETAIL_0_Text_Proc
                                                                                'BOX����
                                                                                
                                                                                
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                        
KONPOU_ON = KONPOU_ON - KONPOU_ON_SUMI '2011.03.17
    
''' �i�ԒP�ʂł̊ۂߏ���
Select Case KONPOU_ON  '2011.03.17
'''''''''''''''''Select Case (KONPOU_ON - KONPOU_ON_SUMI)    '2011.03.17
                            Case 0
                            '�W������Ȃ�
                            
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                                
                                If KONPOU_ON_SUMI <> 0 And KONPOU_OFF_SUMI = 0 Then
                                    If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                    
                                    
                                    
                                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                        
                                        '-----------------------------------------------�w�b�_�[
                                        Call Wel_Head_Text_Proc
                                        '-----------------------------------------------�P�s��
                                        Call Wel_DETAIL_0_Text_Proc
                                        '-----------------------------------------------�Q�s��
                                        Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                        '-----------------------------------------------�R�s��
                                                                                                'BOX����
                                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                                '�\�����e
                                                                                                
                                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                                
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
                                    
                                    
                                    
                                    
                                    
                                    
                                        TOTAL_KUTI_SU = 1
                                        KUTI_SU_INPUT_F = True
                                        TOTAL_SAI_SU = 0#
                                            
                                            
                                            
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                        End If
                                            
                                            
                                
                                        Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                        ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                        ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                
                                
                                
                                        Sendbuf = Text_Create_Proc()
                                    
                                        Inspe_Proc_LOGISTIC = False
                                        Exit Function
                                    
                                    
                                    End If
                                End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   2010.03.05
                            
                            
                            
                            
                                '-----------------------------------------------�{�f�B
                                Call Wel_Hin_No_Req_Text_Proc(Sumi_CNT, Y_SYU_CNT)
        
                                Sendbuf = Text_Create_Proc()
                            Case Else
                            '�W�������
                        
                                '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                                Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                '-----------------------------------------------�R�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                                Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                        '�\�����e
                                                                                    '�\�����e
                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                        "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                        "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                                                                        '���l�����\��
                                Send_Text.Box_Type(3).INIT = ""
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                        '�����J�[�\���ʒu
                                Send_Text.Box_Type(3).Start_Pos = "01"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                        '���͌���
                                '2010.12.07
'                                Send_Text.Box_Type(3).Max_Size = "13"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                Send_Text.Box_Type(3).Max_Size = "20"
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                                '2010.12.07
                                                                                        
                                Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                '-----------------------------------------------�T�s��
                                Call Wel_Clear_Text_Proc
        
                                Sendbuf = Text_Create_Proc()
                        
                        End Select
                    End Select
                Next i
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
'                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                Select Case Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), 2)
                    
                    Case LCD_Hinban     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                        
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                
                        End Select
                '�W������L�莞�̕i�ԃ`�F�b�N
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                        '�Y���i�ԗL��������
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                Exit For
                            End If
                        Next j
                        
                        If j > UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL) Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")      '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")  '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        End If
                        
                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND <> "1" Then
                        
                            If KONPOU_ON <> KONPOU_ON_SUMI Then
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F = "1" Then
                            
                            
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")      '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�i�ԃG���[", "")  '
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_LOGISTIC = False
                                    Exit Function
                            
                                End If
                            End If
                        
                        End If
                        '��ݾق�����
                        CANCEL_F = True
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                    CANCEL_F = False
                                    Exit For
                                End If
                            
                            End If
                        Next j
                        
                        If CANCEL_F Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�L�����Z���i�Ԃł��B", "")        '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "�L�����Z���i�Ԃł��B", "")    '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        
                        
                        End If
                        
                        
                        
                                
                        '���i�ς݂�����
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                Exit For
                            End If
                        Next j
                
                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���i�ς݁I", "")  '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���i�ς݁I", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        End If
                
                        '���o�ɂ�����   2007.05.14
                        If Inspection_Flg = 0 Then
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KAN_KBN <> KAN_KBN_FIN Then
                                        Exit For
                                    End If
                                End If
                            Next j
                                            
                            If j <= UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���o�ɕ��L��I�I", "")        '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, Hinban, "���o�ɕ��L��I�I", "")    '2017.09.22
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            End If
                        End If
                        '���o�ɂ�����   2007.05.14
                        ID_KANRI_TBL(ING_No).KEN_HINBAN = Hinban
                
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo4_REQ
                        
                        
                        '-----------------------------------------------�w�b�_�[
                        Call Wel_Head_Text_Proc
                        '-----------------------------------------------�P�s��
                        Call Wel_DETAIL_0_Text_Proc

''' �i�ԒP�ʂł̊ۂߏ���
                        
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
''' �i�ԒP�ʂł̊ۂߏ���
                        '-----------------------------------------------�Q�s��
                        Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                        
                        
                        
                        
                        
                        
                        
                        '-----------------------------------------------�R�s��
                                                                                'BOX����
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '�\�����e


                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)


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
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Hinban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, Hinban)
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
                        If Inspection_QTY = 1 Then

                            Send_Text.Box_Type(4).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        Else
                            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM                             '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM        '2007.04.21
                        End If
                        
                        Y_SYU_CNT = 0
                        SYUKA_QTY = 0
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                            If Trim(Hinban) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                End If
                            End If
                        Next j
                                                                                '�\�����e
                        
                        If Y_SYU_CNT < 2 Then

                            If Inspection_QTY = 1 Then

                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide))
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide))
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka_Su1)                         '2007.04.21
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka_Su1)    '2007.04.21
                            End If
                                                                                
                        Else
                        
                            If Inspection_QTY = 1 Then
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide) & "*")
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(SYUKA_QTY, "#0"), vbWide) & "*")
                            Else
                                Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka_Su2)                       '2007.04.21
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka_Su2)  '2007.04.21
                            End If
                        
                        End If
                                                                                
                                                                                '���l�����\��
                        If Inspection_QTY = 1 Then
                            Send_Text.Box_Type(4).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                        Else
                            Send_Text.Box_Type(4).INIT = ""                                                     '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""                                '2007.04.21
                        End If
                                                                                '�����J�[�\���ʒu
                        If Inspection_QTY = 1 Then

                            Send_Text.Box_Type(4).Start_Pos = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
                        Else
                            Send_Text.Box_Type(4).Start_Pos = "10"                                          '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "10"                     '2007.04.21
                        End If
                                                                                
                                                                                '���͌���
                        If Inspection_QTY = 1 Then
                            Send_Text.Box_Type(4).Max_Size = "00"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
                        Else
                            Send_Text.Box_Type(4).Max_Size = "07"                                           '2007.04.21
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "07"                      '2007.04.21
                        End If
                                                                                
                                                                                
                        '2009.04.15
                        If SYUKA_QTY > 1 Then
                            Send_Text.buzzer = Buzzer_DOUBLE                    '�u�U�[���@�W��
                            ID_KANRI_TBL(ING_No).Send_Text.buzzer = Buzzer_DOUBLE
                                                                                
                        End If
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
                        
                        Sendbuf = Text_Create_Proc()
                
                End Select
            
            Next i
'''''''''''''''''''''''''''''''
    
    
        Case Step_Sagyo4_RES        '�S��ڂ̎�M�i���i���@��M�j
            
            For i = 0 To M_Gyo - 1
                
'                Select Case RTrim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), _
'                                Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), 10))
                    
                    Case LCD_Syuka_Su1, LCD_Syuka_Su2, "�o�א��F"  '�o�א�(���i��)
                        
                        SURYO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                        
                        If Not IsNumeric(SURYO) Then
                        
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        
                        End If
                
                        Y_SYU_CNT = 0
                        SYUKA_QTY = 0
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                End If
                            End If
                        Next j
                
                        If CLng(SURYO) <> SYUKA_QTY Then
                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�א��G���[", "", "")       '2017.09.22
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�o�א��G���[", "", "")   '2017.09.22
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_LOGISTIC = False
                            Exit Function
                        End If
                
                End Select
            
            Next i
            
            Y_SYU_CNT = 0
            SYUKA_QTY = 0
            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
            
                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                    
                        Y_SYU_CNT = Y_SYU_CNT + 1
                        SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                    End If
                End If
            Next j
            
            '----------------------------------- �f�[�^�X�V�����J�n -----------
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                Exit Function
            End If
                                            
            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                        
                        '------------------------------------   �o�ח\��̏���
                        Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                            'ID��
                        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
        
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
                                    Inspe_Proc_LOGISTIC = False
                                    GoTo Abort_Tran
                                 Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_LOGISTIC = False
                                    GoTo Abort_Tran
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                    
                        Loop
        
                    '''���s�̎g�p���p�����邽��
                    '''Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    '''Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    
                        Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, Format(Now, "YYYYMMDD"))
                                                    
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
                                    Inspe_Proc_LOGISTIC = False
                                    GoTo Abort_Tran
                            
                                Case Else
                                    
                                    Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                    Inspe_Proc_LOGISTIC = SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                        Loop
                        '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                        
                        Call UniCode_Conv(K4_Y_SYU_H.ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)    'ID��
        
                        Do
                        
                            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrKeyNotFound
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")        '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")    '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_LOGISTIC = False
                                    GoTo Abort_Tran
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")    '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_LOGISTIC = False
                                    GoTo Abort_Tran
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                    
                        Loop
                                            
                                            
                        Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, Format(Now, "YYYYMMDDHHMMSS"))
                        Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
                                            
                        Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, ID_KANRI_TBL(ING_No).KEN_OKURI_NO)
'                        '�^����Еϊ�
'                        If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, 3) = UNSOU_KAISHA_CODE Then
'                            Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU_KAISHA_NAME)
'                        End If
'                        '�V�^����Еϊ� 2007.01.09
'
'                        If KURUME_F Then        '�v����
'                            For k = 1 To UBound(KURUME)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(KURUME(k))) = KURUME(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, KURUME(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
'
'                        If FUKUYAMA_F Then      '���R
'                            For k = 1 To UBound(FUKUYAMA)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(FUKUYAMA(k))) = FUKUYAMA(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, FUKUYAMA(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
'
'                        If SAGAWA_F Then        '����
'                            For k = 1 To UBound(SAGAWA)
'
'                                If Left(ID_KANRI_TBL(ING_No).KEN_OKURI_NO, Len(SAGAWA(k))) = SAGAWA(k) Then
'                                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, SAGAWA(0))
'                                    Exit For
'                                End If
'                            Next k
'                        End If
                                                    
                                                    
                                                    
                        Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                        Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                                    
                                                    
                                                    
                                                    '�o�ח\��(νĲҰ��)������
                        Do
                            sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Inspe_Proc_LOGISTIC = False
                                    GoTo Abort_Tran
                            
                                Case Else
                                    Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                    Inspe_Proc_LOGISTIC = SYS_ERR
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    GoTo Abort_Tran
                            End Select
                        Loop
                                            
                                            
                        '------------------------------------   �݌Ɉړ������̏���
                        If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                            MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                                    
                        Else
                            MENU_NO = ""
                        End If
                                            
                        '�����o�ׂ͂̈̓ǂݍ���
                        Call UniCode_Conv(K0_MTS.MUKE_CODE, ID_KANRI_TBL(ING_No).MTS_CODE)
                        Call UniCode_Conv(K0_MTS.SS_CODE, "")
                        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(MTSREC.MUKE_DNAME, "")
                                Call UniCode_Conv(MTSREC.MUKE_NAME, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "������}�X�^", 0)
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                        End Select
                                            
                        sts = IDOREKI_OUTPUT_PROC("", _
                                                    "", _
                                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                                    ID_KANRI_TBL(ING_No).KEN_HINBAN, _
                                                    "", _
                                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                                    0, _
                                                    0, _
                                                    (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                                    FILE_RETRY, _
                                                    CYU_KBN_SPO, _
                                                    Trim(StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)) & " �����:" & Trim(StrConv(Y_SYU_HREC.OKURI_NO, vbUnicode)), _
                                                    , , , , MENU_NO, _
                                                    ID_KANRI_TBL(ING_No).MTS_CODE, _
                                                    "", _
                                                    ID_KANRI_TBL(ING_No).ID_NO & "-" & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO, , , , 1)
                        Select Case sts
                            Case False      '����I��
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc_LOGISTIC = SYS_ERR
                                GoTo Abort_Tran
                        End Select
                                            
                        '���i�ρI�I
                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI = True
                        
                        '�^�����
                        ID_KANRI_TBL(ING_No).KEN_UNSOU_KAISHA = StrConv(Y_SYU_HREC.UNSOU_KAISHA, vbUnicode)
                                        
                    End If
                End If
            
            Next j
            '���۸ޏo��    2009.04.17
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
                                                     ID_KANRI_TBL(ING_No).KEN_HINBAN, , , , , _
                                                     ID_KANRI_TBL(ING_No).ID_NO) Then
                    Inspe_Proc_LOGISTIC = SYS_ERR
                    Exit Function
                End If
            End If
            '���۸ޏo��    2009.04.17
                                '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                    
                    
                    
                    
                    
                    
            '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
            Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
                    
            Select Case ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND
            
            
                Case "1"
                    '�W������Ȃ�
                
                
                    KENPIN_END = True
                    
                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                KENPIN_END = False
                                Exit For
                            End If
                        End If
                    Next j
                
                
                
                
                
                
                    Select Case KENPIN_END
                    
                        Case False
                            '�c����@���i�Ԃ�
''' �׎D���u
                            If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Then
                                ID_KANRI_TBL(ING_No).LABEL_ON = False
                            Else
                            
                                PRINT_OFF = False
                                
                                If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                    Not LOGIXTICS_F Then
                                    If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                        PRINT_OFF = True
                                    End If
                                End If
                                
                                If Not PRINT_OFF Then
                                
                                
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                    '2010.06.16
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            End If
                                        
                                        Case BtErrKeyNotFound
                                        
                                            Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                
                                
                                
                                
                                
                                
                                
                                    PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                            
                                    Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1
                                                            
                                    PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                
                                    ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                    
                                    'Y_SYU_CNT = 0
                                    SYUKA_QTY = 0
                                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                    
                                        If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                            
                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                            
                                                'Y_SYU_CNT = Y_SYU_CNT + 1
                                                SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                            
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                                    PRINT_OFF = True
                                                Else
                                            
                                                   ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = True
                                                
                                                End If
                                            
                                            End If
                                        End If
                                    Next j
                                End If
                                
                                If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                                       Not LOGIXTICS_F) Then
                                    PRINT_MAISU = PRINT_MAISU - 1
                                    If PRINT_MAISU < 1 Then
                                        
                                    '�}�ԍX�V   2010.02.15
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                    
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                            
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                        
                                                        GoTo Abort_Tran
                                                    
                                                    End If
                                                                                                    
                                                
                                                End If
                                            Else
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                    
                                
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                            
                                                            GoTo Abort_Tran
                                                        
                                                        End If
                                                    
                                                    
                                                    
                                                    End If
                                                                        
                                                End If
                                            End If
                                        Next j
                                    '�}�ԍX�V   2010.02.15
                                        
                                        PRINT_OFF = True
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                    
                                    
                                    Else
'                                            PRINT_MAISU = PRINT_MAISU - 1
                                        Start_Page_No = Start_Page_No + 1
                                    End If
                                End If
                                
                                If Not PRINT_OFF Then
                                                            
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                                            
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                            
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            End If
                                        
                                        Case BtErrKeyNotFound
                                        
                                        
                                            Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                                            
                                                            
'2010.02.21                                    PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                            
'2010.02.21                                    Start_Page_No = Start_Page_No + 1
                                                            
'2010.02.21                                    PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                                            
                                                            
                                    If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                    End If
                                
                                
                                
                                    '�}�ԍX�V   2010.02.15
                                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                
                                        If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                        
                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            
                                                If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 And _
                                                    (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                       Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                                       Not LOGIXTICS_F) Then

                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU, "000")) Then
                                                    
                                                        GoTo Abort_Tran
                                                    End If
                                            
                                            
                                                Else
                                            
                                            
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                    
                                                        GoTo Abort_Tran
                                                    End If
                                            
                                                End If
                                            End If
                                        Else
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                
                            
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    
                                                    End If
                                                
                                                
                                                
                                                End If
                                                                    
                                                                    
                                                                    
                                            End If
                                        End If
                                    Next j
                                    '�}�ԍX�V   2010.02.15
                                
                                
                                
                                
'2010.02.21                                    If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
'                                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
'                                    Else
'                                        ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
'                                    End If
                                    
'                                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                        If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'                                                If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                    GoTo Abort_Tran
'                                                End If
'
'                                            End If
'                                        End If
'                                    Next j
                                    
                                    '�f�[�^���M
                                                                
                                    ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                                                
                                                                
                                    ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                                
                                    ID_KANRI_TBL(ING_No).LABEL_ON = True
                                
                                    ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No + PRINT_MAISU - 1
                                    '-----------------------------------------------�w�b�_�[
                                
                                    Call Wel_Head_Print_Text_Proc(FileName)
                                
                                    Sendbuf = Text_Create_Proc()
                                    
                                
                                    Inspe_Proc_LOGISTIC = False
                                    Exit Function
                                End If
                            
                            End If
''' �׎D���u
                            
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            
                            '-----------------------------------------------�w�b�_�[ 02.24
                            Call Wel_Head_Text_Proc
                            
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            
                            
                            '-----------------------------------------------�R�s��
                            Call Wel_HIN_NO_Req_Text_3_Proc
                            '-----------------------------------------------�S�s��
                            Call Wel_HIN_NO_Req_Text_4_Proc
                            
                            
                            
'                            '-----------------------------------------------�R�s��
'                                                                                    'BOX����
'                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                                                                                    '�\�����e
'
'                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'                                                                                    '���l�����\��
'                            Send_Text.Box_Type(2).INIT = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
'                                                                                    '�����J�[�\���ʒu
'                            Send_Text.Box_Type(2).Start_Pos = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
'                                                                                    '���͌���
'                            Send_Text.Box_Type(2).Max_Size = "00"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
'
'                            Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
'                            '-----------------------------------------------�S�s��
'                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                                                                                    '�\�����e
'                                                                                '�\�����e
'                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                                                                                    '���l�����\��
'                            Send_Text.Box_Type(3).INIT = ""
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
'                                                                                    '�����J�[�\���ʒu
'                            Send_Text.Box_Type(3).Start_Pos = "01"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
'                                                                                    '���͌���
'                            Send_Text.Box_Type(3).Max_Size = "13"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
'
'                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                    
                    
                    
                    
                    
                        Case True
                            '�c�Ȃ��@������
                    
                    
''' �׎D���u
                                PRINT_OFF = False
            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" And KONPOU_OFF_SUMI = 0 Then
                                    PRINT_OFF = True
                                Else
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "2" Then
                                    
                                        If KONPOU_ON = KONPOU_ON_SUMI Then
                
                                            PRINT_OFF = True
                
                                        End If
                
                                    End If
                                End If
                                
                                If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Or PRINT_OFF Then
                                    ID_KANRI_TBL(ING_No).LABEL_ON = False
                                Else
                                    
            '                        Y_SYU_CNT = 0
            '                        SYUKA_QTY = 0
                                    
                                    PRINT_OFF = False
                                    
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                        Not LOGIXTICS_F Then
                                        If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                            
                                                        '�}�ԍX�V   2010.02.15
                                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                                                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                    Start_Page_No = 1
                                                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                    
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                        
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                                    
                                                        PRINT_OFF = True
                                        
                                                    End If
                                                
                                                Else
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                        
                                    
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            
                                                            End If
                                                    
                                                        
                                                        End If
                                                                            
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            
                                                If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                                
                                                    Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                                
                                                End If
                                            
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                                Sendbuf = Text_Create_Proc()
                                                GoTo Abort_Tran
                                        End Select
                                        
                                        PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                                
                                        Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1
                                                                
                                        PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                        
                                        If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                               Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                               Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                                               Not LOGIXTICS_F) Then
                                            PRINT_MAISU = PRINT_MAISU - 1
                                            
                                            
                                            
                                            If PRINT_MAISU < 1 Then
                                                
                                                        '�}�ԍX�V   2010.02.15
                                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                            
                                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    
                                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                            
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            End If
                                        
                                                        End If
                                                    
                                                    
                                                    
                                                    Else
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                            
                                        
                                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                    
                                                                    GoTo Abort_Tran
                                                                End If
                                                            
                                                            
                                                        
                                                            End If
                                                                                
                                                        End If
                                                    
                                                    
                                                    
                                                    End If
                                                Next j
                                                        '�}�ԍX�V   2010.02.15
                                                
                                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                                
                                                PRINT_OFF = True
                                            Else
            '                                    PRINT_MAISU = PRINT_MAISU - 1
                                                Start_Page_No = Start_Page_No + 1
                                            End If
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                        
                                        Y_SYU_CNT = 0
                                        SYUKA_QTY = 0
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                        
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                
                                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                                
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                                        PRINT_OFF = True
                                                    Else
                                                
                                                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = True
                                                    
                                                    End If
                                                
                                                End If
                                            End If
                                        Next j
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                        
                                        If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                        End If
                                        
                                        
                                        
                                        
                                        
                                                        '�}�ԍX�V   2010.02.15
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                    
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                            
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                
                                                        
                                                    If ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 And Start_Page_No = 2 Then
                                                    
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    Else
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    End If
                                                
                                                End If
                                            
                                            
                                            
                                            Else
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                    
                                
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    
                                                    End If
                                                                        
                                                End If
                
                                            
                                            
                                            
                                            
                                            End If
                                        Next j
                                                        '�}�ԍX�V   2010.02.15
                                        
                                        
                                        
                                        
                                        If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
                                        Else
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
                                        End If
                                        
'                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'
'
'                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                        GoTo Abort_Tran
'                                                    End If
'                                                End If
'                                            End If
'                                        Next j
                                        
                                        '�f�[�^���M
                                                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 2
                                                                    
                                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                                    
                                        '-----------------------------------------------�w�b�_�[
                                
                                        Call Wel_Head_Print_Text_Proc(FileName)
                                    
                                        Sendbuf = Text_Create_Proc()
                                        
                                    
                                        Inspe_Proc_LOGISTIC = False
                                        Exit Function
                                    
                                    End If
                                End If
            ''' �׎D���u
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Text_Proc
                                '-----------------------------------------------�P�s��
                                Call Wel_DETAIL_0_Text_Proc
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                '-----------------------------------------------�R�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                        
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
                                '-----------aaa------------------------------------�S�s��
                                
'����INPUT �P
                                
                                wkKONPO_F = ""
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                    
                                        wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                                        Exit For
                                    End If
                                Next j
                                
                                If wkKONPO_F = "1" Then
                                                        
                                    If Inspection_Input Then
                                        KUTI_SU_INPUT_F = False
                                    Else
                                        KUTI_SU_INPUT_F = True
                                    End If
                                
                                
                                    TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                    TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                            
                                Else
                                    TOTAL_KUTI_SU = 1
                                    KUTI_SU_INPUT_F = True
                                    TOTAL_SAI_SU = 0#
                                End If
                                            
                                            
                                            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                End If
                                            
                                            
                                If KUTI_SU_INPUT_F Then
                                
                                    Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                
                                
                                Else
                                    Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                
                                    
                                    
                                    If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                    
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                                
                                
                                
                                End If
                                
                                Sendbuf = Text_Create_Proc()
                        
                        End Select
                        
                
                
                Case "2"
                    '�W������̂�
                
                    KENPIN_END = True
                    
                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                KENPIN_END = False
                                Exit For
                            End If
                        End If
                    Next j
                
                
                
                
                
                
                    Select Case KENPIN_END
                    
                        Case False
                            '�c����@���i�Ԃ�
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            
                            
                            '-----------------------------------------------�w�b�_�[ 02.24
                            Call Wel_Head_Text_Proc
                            
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                                                                                '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                                                                    
                                                                                    
                                                                                    
                                                                                    '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    '���͌���
                            '2010.12.07
'                            Send_Text.Box_Type(3).Max_Size = "13"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                            Send_Text.Box_Type(3).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                            '2010.12.07
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                    
                    
                    
                    
                    
                        Case True
                            '�c�Ȃ��@������
                    
                    
''' �׎D���u
                                PRINT_OFF = False
            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" And KONPOU_OFF_SUMI = 0 Then
                                    PRINT_OFF = True
                                Else
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "2" Then
                                    
                                        If KONPOU_ON = KONPOU_ON_SUMI Then
                
                                            PRINT_OFF = True
                
                                        End If
                
                                    End If
                                End If
                                
                                If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Or PRINT_OFF Then
                                    ID_KANRI_TBL(ING_No).LABEL_ON = False
                                Else
                                    
            '                        Y_SYU_CNT = 0
            '                        SYUKA_QTY = 0
                                    
                                    PRINT_OFF = False
                                    
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                        Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                        Not LOGIXTICS_F Then
                                        If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                            
                                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                        
                                                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                    
                                                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                    
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                                        Start_Page_No = 1
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        PRINT_OFF = True
                                                    End If
                                                
                                                
                                                Else
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                        
                                    
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            
                                                            End If
                                                        
                                                        
                                                        
                                                        End If
                                                                            
                                                    End If
                                                
                                                
                                                
                                                
                                                
                                                End If
                                            Next j
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).KEN_HINBAN)
                                                                
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            
                                                If Not IsNumeric(StrConv(ITEMREC.KUTI_SU, vbUnicode)) Then
                                                
                                                    Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                                
                                                End If
                                            
                                            Case BtErrKeyNotFound
                                            
                                                Call UniCode_Conv(ITEMREC.KUTI_SU, "0001")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                                Sendbuf = Text_Create_Proc()
                                                GoTo Abort_Tran
                                        End Select
                                        
                                        PRINT_MAISU = SYUKA_QTY * CInt(StrConv(ITEMREC.KUTI_SU, vbUnicode))
                                                                
                                        Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1
                                                                
                                        PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                        
                                        If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                                Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                                Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                                                Not LOGIXTICS_F) Then
                                            PRINT_MAISU = PRINT_MAISU - 1
                                            If PRINT_MAISU < 1 Then
                                                
                                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                            
                                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    
                                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                        
                                                        '�}�ԍX�V   2010.02.15
                                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                
                                                                GoTo Abort_Tran
                                                            
                                                            End If
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                                        
                                        
                                        
                                                        End If
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    Else
                                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                            
                                        
                                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                                    
                                                                    GoTo Abort_Tran
                                                                End If
                                                                
                                                        
                                                            
                                                            End If
                                                                                
                                                        End If
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    End If
                                                Next j
                                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                                PRINT_OFF = True
                                            Else
            '                                    PRINT_MAISU = PRINT_MAISU - 1
                                                Start_Page_No = Start_Page_No + 1
                                            End If
                                        End If
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 1
                                        
                                        Y_SYU_CNT = 0
                                        SYUKA_QTY = 0
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                        
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                                
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                
                                                    Y_SYU_CNT = Y_SYU_CNT + 1
                                                    SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                                
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                                        PRINT_OFF = True
                                                    Else
                                                
                                                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI = True
                                                    
                                                    End If
                                                
                                                End If
                                            End If
                                        Next j
                                    End If
                                    
                                    If Not PRINT_OFF Then
                                        
                                        If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                        End If
                                        
                                        
                                        
                                        
                                        
                                        '�}�ԍX�V   2010.02.15
                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                    
                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                            
                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    End If
                                                End If
                                            
                                            Else
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                                    
                                
                                                        If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                            
                                                            GoTo Abort_Tran
                                                        End If
                                                    
                                                    
                                                    End If
                                                                        
                                                End If
                                            
                                            
                                            
                                            
                                            
                                            
                                            
                                            End If
                                        Next j
                                        '�}�ԍX�V   2010.02.15
                                        
                                        
                                        
                                        
                                        
                                        If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
                                        Else
                                            ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
                                        End If
                                        
'                                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                            If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'
'
'                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                        GoTo Abort_Tran
'                                                    End If
'
'                                                End If
'                                            End If
'                                        Next j
                                        
                                        '�f�[�^���M
                                                                    
                                        ID_KANRI_TBL(ING_No).LABEL_STEP = 2
                                                                    
                                        ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                                    
                                        ID_KANRI_TBL(ING_No).LABEL_ON = True
                                    
                                        '-----------------------------------------------�w�b�_�[
                                
                                        Call Wel_Head_Print_Text_Proc(FileName)
                                    
                                        Sendbuf = Text_Create_Proc()
                                        
                                    
                                        Inspe_Proc_LOGISTIC = False
                                        Exit Function
                                    
                                    End If
                                End If
            ''' �׎D���u
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                                
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Text_Proc
                                '-----------------------------------------------�P�s��
                                Call Wel_DETAIL_0_Text_Proc
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                '-----------------------------------------------�R�s��
                                                                                        'BOX����
                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                        '�\�����e
                                                                                        
                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                        
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
                                '-----------aaa------------------------------------�S�s��
                                
'����INPUT �Q
                                wkKONPO_F = ""
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                    
                                        wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                                        Exit For
                                    End If
                                Next j
                                
                                If wkKONPO_F = "1" Then
                                                        
                                    If Inspection_Input Then
                                        KUTI_SU_INPUT_F = False
                                    Else
                                        KUTI_SU_INPUT_F = True
                                    End If
                                
                                
                                    TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                    TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                            
                                Else
                                    TOTAL_KUTI_SU = 1
                                    KUTI_SU_INPUT_F = True
                                    TOTAL_SAI_SU = 0#
                                End If
                                            
                                            
                                            
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                                End If
                                            
                                            
                                If KUTI_SU_INPUT_F Then
                                
                                    Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                                
                                
                                Else
                                    Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                
                                    If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                
                                
                                    ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                                    ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                                
                                
                                
                                End If
                                
                                Sendbuf = Text_Create_Proc()
                        
                        End Select
                    
                Case "3"
                    '����
            
                    
                    '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                    Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
            
            
                    Select Case (KONPOU_ON - KONPOU_ON_SUMI)
            
                        Case 0
                            '������
                            ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                            
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                                    
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
                            '-----------aaa------------------------------------�S�s��
'����input �R
                            wkKONPO_F = ""
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            
                                If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                
                                    wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                                    Exit For
                                End If
                            Next j
                            
                            If wkKONPO_F = "1" Then
                                                    
                                If Inspection_Input Then
                                    KUTI_SU_INPUT_F = False
                                Else
                                    KUTI_SU_INPUT_F = True
                                End If
                            
                            
                                TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                                TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                        
                            Else
                                TOTAL_KUTI_SU = 1
                                KUTI_SU_INPUT_F = True
                                TOTAL_SAI_SU = 0#
                            End If
                                        
                                        
                                        
                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                            End If
                                        
                                        
                            If KUTI_SU_INPUT_F Then
                            
                                Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                                ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                                ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                            
                            
                            Else
                                Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                            
                                
                                If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Exit Function
                                End If
                                
                                
                                ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                                ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                            
                            
                            
                            End If
                        
                        
                            Sendbuf = Text_Create_Proc()
                        
                        
                        
                        
                        
                        Case Else
                            '�i�Ԃ�
            
                            '�c����@���i�Ԃ�
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                            '-----------------------------------------------�w�b�_�[
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                                                                                    'BOX����
                            Send_Text.Box_Type(2).Box_Type = TYPE_REF
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                    '�\�����e
                                                                                    
                            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
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
                            Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                                    '�\�����e
                                                                                '�\�����e
                            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_HIN_SYUKON & _
                                                                    "(" & Format(KONPOU_ON_SUMI, "0") & "/" & Format(KONPOU_ON, "0") & ")")
                                                                                    
                                                                                    
                                                                                    
                                                                                    '���l�����\��
                            Send_Text.Box_Type(3).INIT = ""
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                                    '�����J�[�\���ʒu
                            Send_Text.Box_Type(3).Start_Pos = "01"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                                    '���͌���
                            '2010.12.07
'                            Send_Text.Box_Type(3).Max_Size = "13"
'                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                            Send_Text.Box_Type(3).Max_Size = "20"
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                            '2010.12.07
                                                                                    
                            Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                            '-----------------------------------------------�T�s��
                                                                                    'BOX����
                            Call Wel_Clear_Text_Proc
    
                            Sendbuf = Text_Create_Proc()
                    End Select
            
            End Select
                    
                    
                    
                    
        Case Step_Sagyo5_RES        '�T��ڂ̎�M�i�����j
                
            For i = 0 To M_Gyo - 1
                
                Select Case Left(Trim(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)), 6)
                    '����
                    Case LCD_KUTI_SU_S
                
                
                        If ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU < 0 Then
                        
                
                
                
                            If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")     '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "") '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            
                            End If
                    
                            KUTI_SU = CInt(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If KUTI_SU < 1 Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "")     '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�����G���[", "", "") '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            End If
                        Else
                            KUTI_SU = ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU
                        End If
                    
                    
                    
                    '�ː�
                    Case LCD_SAI_SU_S
                
                        If ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU < 0 Then
                        
                        
                            If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")     '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "") '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            
                            End If
                    
                            SAI_SU = CDbl(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If SAI_SU <= 0 Then
                                'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "")     '2017.09.22
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "�ː��G���[", "", "") '2017.09.22
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            End If
                        Else
                            SAI_SU = ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU
                        
                            If SAI_SU < 1 Then
                                SAI_SU = 1
                            Else
                                If SAI_SU > 1 Then
                                    SAI_SU = CLng(ToHalfAdjust(CCur(SAI_SU), 0))
                                End If
                            End If
                        
                        
                        End If
                    
                        '�����ő��������l�� 2010.01.21
                            
                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                            If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                
                                
                                    If Label_Print_Total_Su_Proc(KUTI_SU, PRINT_TOTAL_SU) Then
                                
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                
                                
                                
                                Else
                                
                                    If Label_Print_Total_Su_Proc(0, PRINT_TOTAL_SU) Then
                                
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        Exit Function
                                    End If
                                
                                End If
                            End If
                        Next j
                                        
                            
                            
                            
                        
'                        If Label_Print_Total_Su_Proc(KUTI_SU, PRINT_TOTAL_SU) Then
'
'                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                            Sendbuf = Text_Create_Proc()
'                            Exit Function
'                        End If
                    
                        ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = PRINT_TOTAL_SU
                
                        
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                            If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU = KUTI_SU
                                ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SAI_SU = SAI_SU
                            Else
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU <= 1 Then
                                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KUTI_SU = KUTI_SU
                                    End If
                                
                                    ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SAI_SU = SAI_SU
                                
                                End If
                            End If
                        Next j
                        
                        
                        
                        '----------------------------------- �f�[�^�X�V�����J�n -----------
                                                        '�g�����U�N�V�����J�n
                        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                            Exit Function
                        End If
                                            
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                        
                        
                        '------------------------------------   �o�ח\��̏���
                            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                'ID��
                            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
            
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
                                        Inspe_Proc_LOGISTIC = False
                                        GoTo Abort_Tran
                                     Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_LOGISTIC = False
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
                
                                                '�o�ח\�菑����
                            Do
                                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")       '2017.09.22
                                        Sendbuf = Text_Create_Proc()
                                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                        Inspe_Proc_LOGISTIC = False
                                        GoTo Abort_Tran
                                
                                    Case Else
                                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                                        Inspe_Proc_LOGISTIC = SYS_ERR
                                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                        Sendbuf = Text_Create_Proc()
                                        GoTo Abort_Tran
                                End Select
                            Loop
                                '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                    
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F And ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "") '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")    '2017.09.11
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                            
                                
                                
                                
                                
                                
                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                
                                    'Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F)
                                    If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) Then
                                        If CInt(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) > 0 Then
                                        Else
                                            Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "0000"))
                                        End If
                                    Else
'''''''                                        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "00.00"))
                                        
                                        
                                        Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "00.00"))
                                                                            
                                    End If
                                    Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN, Format(SAI_SU, "00.00"))
                                                    
                                Else
                                                
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                                        
                                        'Call UniCode_Conv(Y_SYU_HREC.KONPOU_F, ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F)
                                        If IsNumeric(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) Then
                                            If CInt(StrConv(Y_SYU_HREC.KUTI_SU_TAN, vbUnicode)) > 0 Then
                                            Else
                                                Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "0000"))
                                            End If
                                        Else
                                            Call UniCode_Conv(Y_SYU_HREC.KUTI_SU_TAN, Format(KUTI_SU, "0000"))
                                                                                
                                        End If
                                    End If
                                    Call UniCode_Conv(Y_SYU_HREC.SAI_SU_TAN, Format(SAI_SU, "00.00"))
                                                
                                End If
                                                    
                                
                                
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                
                                
                                
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_LOGISTIC = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                

'                        Call Syuka_KUTI_SU_Count_Proc(TOTAL_KUTI_SU)

        
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                   
                        '------------------------------------   �o�ח\��(νĲҰ��)�̏���
                
                            'ID_NO
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F And ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                
                                Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
                                                                                                    'ID��
                                Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Trim(ID_KANRI_TBL(ING_No).ID_NO) & ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SEQ_NO)
                
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
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                         Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")      '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")  '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                            
                                Loop
                                
                                
                                
                                Call UniCode_Conv(K4_Y_SYU_H.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))                                                                                           '�ǔ�
        
                                Do
                        
                                    sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrKeyNotFound
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "")     '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�s��", "", "") '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��(νĲҰ��)", 0)
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                    
                                Loop
                                                
                                Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, Format(KUTI_SU, "0000"))
                                Call UniCode_Conv(Y_SYU_HREC.SAI_SU, Format(SAI_SU, "00.00"))
                                                    
                                                    
                                                    
                                Call UniCode_Conv(Y_SYU_HREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))  '2015.01.10
                                Call UniCode_Conv(Y_SYU_HREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))   '2015.01.10
                                                    
                                                    
                                                    '�o�ח\��(νĲҰ��)������
                                Do
                                    sts = BTRV(BtOpUpdate, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K4_Y_SYU_H, Len(K4_Y_SYU_H), 4)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                            'Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")       '2017.09.22
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_LONG_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��(H)�g�p��", "", "")   '2017.09.22
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                            Inspe_Proc_LOGISTIC = False
                                            GoTo Abort_Tran
                                    
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�o�ח\��(νĲҰ��)", 0)
                                            Inspe_Proc_LOGISTIC = SYS_ERR
                                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                            Sendbuf = Text_Create_Proc()
                                            GoTo Abort_Tran
                                    End Select
                                Loop
                            End If
                                        
            
                        Next j
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                            '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                

''' �i�ԒP�ʂł̊ۂߏ���
                        
                        '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
                        Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
''' �i�ԒP�ʂł̊ۂߏ���

                            
'                        PRINT_OFF = False
'
'                        If KONPOU_OFF = KONPOU_OFF_SUMI Then
'
'                            PRINT_OFF = True
'
'                        End If

                        If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 0 Then
                            ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = KUTI_SU
                        End If


                        If Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                            Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                            Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                            Not LOGIXTICS_F Then
                            If ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU = 1 Then
                                PRINT_OFF = True
                            End If
                        End If



                        PRINT_MAISU = KUTI_SU
                                                
                                                
                        Start_Page_No = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + 1

                        PRINT_TOTAL_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU


                        If Start_Page_No = 1 And (Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_CHARTER_CD) And _
                                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_AKABOU_CD) And _
                                                    Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_NO) <> Trim(KEN_LOGISTIC_CD) And _
                                                    Not LOGIXTICS_F) Then
                            PRINT_MAISU = PRINT_MAISU - 1
                            If PRINT_MAISU < 1 Then
                                
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                    
                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                        
                        
                                                        '�}�ԍX�V   2010.02.15
                                            If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                
                                                GoTo Abort_Tran
                                            End If
                                                        
                                                        
                                                        '�}�ԍX�V   2010.02.15
                        
                                        End If
                                    Else
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                            
                        
                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf) Then
                                                    
                                                    
                                                    GoTo Abort_Tran
                                                
                                                End If
                                            
                                            
                                            End If
                                                                
                                        End If
                                    End If
                                Next j
                                
                                
                                
                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No
                                
                                
                                
                                
                                PRINT_OFF = True
                            Else
'                                PRINT_MAISU = PRINT_MAISU - 1
                                Start_Page_No = Start_Page_No + 1
                            
                            End If
                        End If

''' �׎D���u
                        If Trim(F0_SendFile) = "" Or Trim(ID_KANRI_TBL(ING_No).CTR_TYPE) = "" Or PRINT_OFF Then
                            ID_KANRI_TBL(ING_No).LABEL_ON = False
                        Else
                                
                            Y_SYU_CNT = 0
                            SYUKA_QTY = 0
                            
                            PRINT_OFF = False
                            
                            For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                            
                                If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                                    
                                    If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                    
                                        Y_SYU_CNT = Y_SYU_CNT + 1
                                        SYUKA_QTY = SYUKA_QTY + ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SURYO
                                    
                                    
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).PRINT_SUMI Then
                                        
                                            PRINT_OFF = True
                                        Else
                                    
                                        End If
                                    End If
                                End If
                            Next j
                                
                            If Not PRINT_OFF Then
                                
                                If Label_File_Make_Proc(FileName, PRINT_MAISU, Start_Page_No, PRINT_TOTAL_SU) Then
                                End If
                            
                            
                            
                            
                                '�}�ԍX�V   2010.02.15
                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                            
                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
                    
                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                        
                        
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" And Start_Page_No = 2 Then
                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU, "000")) Then
                                                    
                                                    GoTo Abort_Tran
                                                
                                                
                                                End If
                        
                                            Else
                        
                        
                                                        '�}�ԍX�V   2010.02.15
                                                If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                    
                                                    GoTo Abort_Tran
                                                
                                                
                                                End If
                                                        
                                            End If
                                                        
                                                        '�}�ԍX�V   2010.02.15
                        
                        
                        
                        
                        
                                        End If
                                    Else
                                        If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                            If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" Then
                                            
                        
                                                If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F <> "1" And Start_Page_No = 2 Then
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No - 1, "000"), Sendbuf, Format(Start_Page_No - 1 + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    End If
                                                Else
                                                
                                                    If OKURI_NO_SEQ_Update_Proc(j, Format(Start_Page_No, "000"), Sendbuf, Format(Start_Page_No + PRINT_MAISU - 1, "000")) Then
                                                        
                                                        GoTo Abort_Tran
                                                    End If
                                                
                                                                        
                                                
                                                End If
                                            End If
                                        End If
                                    End If
                                Next j
                                '�}�ԍX�V   2010.02.15
                            
                            
                            
'                                ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = Start_Page_No + PRINT_MAISU
                                
                                If Start_Page_No = 2 And ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = 0 Then
                                    ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU + 1
                                Else
                                    ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO = ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO + PRINT_MAISU
                                End If
                                
'                                For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
'
'                                    If Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) = Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) Then
'
'                                        If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
'
'
'
'                                            If OKURI_NO_SEQ_Update_Proc(j, Format(ID_KANRI_TBL(ING_No).LABEL_START_PAGE_NO, "000"), Sendbuf) Then
'                                                GoTo Abort_Tran
'                                            End If
'
'                                        End If
'                                    End If
'                                Next j
                            
                                ID_KANRI_TBL(ING_No).LABEL_STEP = 9
                                
                                '�f�[�^���M
                                                            
                                ID_KANRI_TBL(ING_No).Step = Step_PRINT_REQ
                            
                                ID_KANRI_TBL(ING_No).LABEL_ON = True
                            
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Print_Text_Proc(FileName)
                                '-----------------------------------------------�{�f�B
                                Call Wel_Hin_No_Req_Text_Proc(Sumi_CNT, Y_SYU_CNT)
                            
                                Sendbuf = Text_Create_Proc()
                            
                                Inspe_Proc_LOGISTIC = False
                                Exit Function
                            
                            End If

''' �׎D���u
                        End If

                        KENPIN_END = True
                        
                        For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                                                            
                                                            
                            If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).CANCEL_F Then
                                If Not ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).SUMI Then
                                    KENPIN_END = False
                                    Exit For
                                End If
                            End If
                        Next j



                        Select Case KENPIN_END
                    
                            Case True
    
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
                                            
                                            
                                            
                                            
                                            
                    
                    
                            Case Else
                    
    ''''''''''''''''''''''''''''''''''''''''''''''
                                ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                    
                                '-----------------------------------------------�w�b�_�[
                                Call Wel_Head_Text_Proc
                                '-----------------------------------------------�P�s��
                                Call Wel_DETAIL_0_Text_Proc
                                '-----------------------------------------------�Q�s��
                                Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                                
                                 '-----------------------------------------------�R�s��
                                Call Wel_HIN_NO_Req_Text_3_Proc
                                '-----------------------------------------------�S�s��
                                Call Wel_HIN_NO_Req_Text_4_Proc
                               
'                                '-----------------------------------------------�R�s��
'                                                                                        'BOX����
'                                Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
'                                                                                        '�\�����e
'
'                                Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
'
'
'                                                                                        '���l�����\��
'                                Send_Text.Box_Type(2).INIT = ""
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
'
'�@                                                                                      '�����J�[�\���ʒu
'                                Send_Text.Box_Type(2).Start_Pos = ""
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
'                                                                                        '���͌���
'                                Send_Text.Box_Type(2).Max_Size = "00"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
'
'                                Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
'                                '-----------------------------------------------�S�s��
'                                Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
'                                                                                        '�\�����e
'                                                                                    '�\�����e
'                                Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                                Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Hinban)
'                                                                                        '���l�����\��
'                                Send_Text.Box_Type(3).INIT = ""
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
'                                                                                        '�����J�[�\���ʒu
'                                Send_Text.Box_Type(3).Start_Pos = "01"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
'                                                                                        '���͌���
'                                Send_Text.Box_Type(3).Max_Size = "13"
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
'
'                                Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
'                                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                                '-----------------------------------------------�T�s��
                                Call Wel_Clear_Text_Proc
        
                                Sendbuf = Text_Create_Proc()
                    
                        End Select
                    
                    End Select
                
            Next i
        

        Case Step_PRINT_RES        '����I��
    
    
            '�o�׎��с^�o�ח\�萔�^�W��������i�\��j�^�P�̍�����i�\��j�^�W��������i���сj�^�P�̍�����i���сj�̃J�E���g
            Call Syuka_Kenpin_Count_Proc(Sumi_CNT, Y_SYU_CNT, KONPOU_ON, KONPOU_OFF, KONPOU_ON_SUMI, KONPOU_OFF_SUMI)
    
            '-----------------------------------------------�w�b�_�[
            Call Wel_Head_Text_Proc
    
            Select Case ID_KANRI_TBL(ING_No).LABEL_STEP
                Case 1      '�i�Ԃ�
                                                                            
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                                                                            
                    '-----------------------------------------------�w�b�_�[ 02.24
                    Call Wel_Head_Text_Proc
                                                                            
                    '-----------------------------------------------�P�s��
                    Call Wel_DETAIL_0_Text_Proc
                    '-----------------------------------------------�Q�s��
                    Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                    '-----------------------------------------------�R�s��
                                                                            'BOX����
                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                            '�\�����e
                                                                            
                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, Trim(ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI))
                                                                            
                                                                            
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
                    Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                            '�\�����e
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
                    '2010.12.07
'                    Send_Text.Box_Type(3).Max_Size = "13"
'                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                    Send_Text.Box_Type(3).Max_Size = "20"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "20"
                    '2010.12.07
                                                                            
                    Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                    '-----------------------------------------------�T�s��
                    Call Wel_Clear_Text_Proc

                    Sendbuf = Text_Create_Proc()
                
                Case 2      '������
                
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo5_REQ
                    '-----------------------------------------------�w�b�_�[
                    Call Wel_Head_Text_Proc
                    
                    '-----------------------------------------------�P�s��
                    Call Wel_DETAIL_0_Text_Proc
                    '-----------------------------------------------�Q�s��
                    Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                    
                    '-----------------------------------------------�R�s��
                                                                            'BOX����
                    Send_Text.Box_Type(2).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                            '�\�����e
                                                                            
                    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).KEN_OKURI_SAKI)
                                                                            
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
'�����@�S
                    wkKONPO_F = ""
                    For j = 0 To UBound(ID_KANRI_TBL(ING_No).KEN_DEN_TBL)
                    
                        If Trim(ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).HIN_NO) = Trim(ID_KANRI_TBL(ING_No).KEN_HINBAN) Then
                        
                            wkKONPO_F = ID_KANRI_TBL(ING_No).KEN_DEN_TBL(j).KONPOU_F
                            Exit For
                        End If
                    Next j
                    
                    If wkKONPO_F = "1" Then
                                            
                        If Inspection_Input Then
                            KUTI_SU_INPUT_F = False
                        Else
                            KUTI_SU_INPUT_F = True
                        End If
                    
                    
                        TOTAL_KUTI_SU = ID_KANRI_TBL(ING_No).LABEL_PRINT_TOTAL_SU
                        TOTAL_SAI_SU = Syuka_END_Count_Proc()
                                
                    Else
                        TOTAL_KUTI_SU = 1
                        KUTI_SU_INPUT_F = True
                        TOTAL_SAI_SU = 0#
                    End If
                                
                                
                                
                    If ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "3" Then
                        ID_KANRI_TBL(ING_No).KEN_DEN_TBL(0).KONPOU_CND = "1"
                    End If
                                
                                
                    If KUTI_SU_INPUT_F Then
                    
                        Call Wel_Kuti_Su_Input_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                        ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = -1
                        ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = -1
                    
                    
                    Else
                        Call Wel_Kuti_Su_Notinput_text_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU)
                    
                        
                        If KutiSai_Update_Proc(TOTAL_KUTI_SU, TOTAL_SAI_SU) Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Exit Function
                        End If
                        
                        
                        ID_KANRI_TBL(ING_No).KEN_INP_KUTI_SU = TOTAL_KUTI_SU
                        ID_KANRI_TBL(ING_No).KEN_INP_SAI_SU = TOTAL_SAI_SU
                    
                    
                    
                    End If
                    
                    Sendbuf = Text_Create_Proc()
                
                Case 9
        
                    Select Case (Y_SYU_CNT - Sumi_CNT)
                
                        Case 0
                
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
            
            
                        Case Else
                        
        ''''''''''''''''''''''''''''''''''''''''''''''
                        
                        
                            ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                
                            '-----------------------------------------------�w�b�_�[ 02.24
                            Call Wel_Head_Text_Proc
                            '-----------------------------------------------�P�s��
                            Call Wel_DETAIL_0_Text_Proc
                            '-----------------------------------------------�Q�s��
                            Call Wel_DETAIL_1_Text_Proc(ID_KANRI_TBL(ING_No).ID_NO, Sumi_CNT, Y_SYU_CNT)
                            '-----------------------------------------------�R�s��
                            Call Wel_HIN_NO_Req_Text_3_Proc
                            '-----------------------------------------------�S�s��
                            Call Wel_HIN_NO_Req_Text_4_Proc
                            '-----------------------------------------------�T�s��
                            Call Wel_Clear_Text_Proc
        
                            Sendbuf = Text_Create_Proc()
        
            
                End Select
            End Select
    
    End Select
                    
                    

    Inspe_Proc_LOGISTIC = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function

