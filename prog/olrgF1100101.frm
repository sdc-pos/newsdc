VERSION 5.00
Object = "{D4A17F03-6EDB-11D2-A6E0-0040262B3978}#2.0#0"; "CtrsWsk.ocx"
Begin VB.Form F1100101 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�X�L���i����u��~���v"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6255
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�ꎞ��~"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   10695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   10695
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ɩ��I��"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ɩ��J�n"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
   Begin CTRSWSKLib.CtrsWsk CtrsWsk1 
      Left            =   240
      Top             =   360
      _Version        =   131072
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "F1100101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LocalPort               As Long                 '�f�[�^��M�|�[�g�ԍ�
Private RemotePort              As Long                 '�f�[�^���M�|�[�g�ԍ�

Private Normal_End              As Boolean              '�I�����

Private Next_Step               As Integer              '�����s�w��

Private Menu_Type               As Integer              '1:�� 2:����


Private Type NAIGAI_tag
    CODE                        As String * 1
    NAME                        As String
End Type

Private NAIGAI()                As NAIGAI_tag

Private Const M_Gyo% = 5                                '�ő��ʍs��
Private Const M_Keta% = 20                              '�ő��ʌ���

Private Type Recv_Text_Tag                              '��M�e�L�X�g
    ID                          As Integer              'IDNO
    LCD(0 To M_Gyo - 1)         As String * 20          '��M���e�P�`�T�s��
    Time                        As String * 8           '���M����
    RETRY                       As Integer              '���M���g���C2004.04.10
End Type

Private Recv_text               As Recv_Text_Tag

Private Const Start_Para$ = "START"     '�q�@�d���n�m
Private Const Can_Para$ = "CANCEL"      '�L�����Z���v��
Private Const Fin_Para$ = "FINISHI"     '��ƏI���v��
Private Const Ent_Para$ = "ENT"         'ENT�̂�

Private Const Qty_OK_Para$ = "OK"       '���ʂn�j
Private Const Loc_OK_Para$ = "T"        '�I�Ԃn�j
Private Const DEN_OK_Para$ = "ALL"      '�`�[�n�j
Private Const LAST_Para$ = "LAST"       '�ŏI���


Private Type Box_Type_tag               '�a�n�w�w��
    Box_Type                    As String * 1           'BOX����
    LCD(0 To M_Keta - 1)        As Byte                 '�\�����e
    INIT                        As String * 10          '�������e
    Start_Pos                   As String * 2           '�J�n�J�[�\���ʒu
    Max_Size                    As String * 2           '���͌����i�ő�j
    MENU                        As String * 9           '���j���[���e
End Type

Private Type Send_Text_tag              '���M�p�o�b�t�@�[
    sts                         As String * 1           '�X�e�[�^�X 1:OK 2:NG
    Display_Flg                 As String * 1           '�\����ʃt���O 1:�ʏ���͉�� 2:���j���[��� 3:�Q�Ɖ��
    End_Menu                    As String * 1           '�ŏI���j���[�t���O 1:����ʂ���@2:�ŏI���
    Menu_Suu                    As String * 2           '���j���[��
    fileName                    As String * 12          '�t�@�C����(*.*)
    Buzzer                      As String * 1           '�u�U�[�w��
    Box_Type(0 To M_Gyo - 1)    As Box_Type_tag
    CRLF                        As String * 2
End Type

Private Send_Text           As Send_Text_tag
Private Const Sts_OK$ = "1"             '�X�e�[�^�X�@OK
Private Const Sts_NG$ = "2"             '�X�e�[�^�X�@NG

Private Const Display_DEF$ = "1"        '�\����ʃt���O�@�ʏ���͉��
Private Const Display_MENU$ = "2"       '�\����ʃt���O�@���j���[���
Private Const Display_REF$ = "3"        '�\����ʃt���O�@�Q�Ɖ��

Private Const Menu_Head$ = "1"          '�ŏI���j���[�t���O �擪(�O����^��Ȃ�)
Private Const Menu_Mid$ = "2"           '�ŏI���j���[�t���O ����(�O����^�゠��)
Private Const Menu_End$ = "3"           '�ŏI���j���[�t���O �ŏI(�O����^��Ȃ�)
Private Const Menu_Only$ = "4"          '�ŏI���j���[�t���O �P��(�O�Ȃ��^��Ȃ�)

Private Const Buzzer_DEF$ = "1"         '�W���̉�
Private Const Buzzer_CONTI$ = "4"       '�A����

Private Const TYPE_REF$ = "X"           '�\���a�n�w
Private Const TYPE_BCANK$ = 1           '�o�[�R�[�h���p����
Private Const TYPE_BCNUM$ = 2           '�o�[�R�[�h������
Private Const TYPE_BCONLY$ = 3          '�o�[�R�[�h�̂�
Private Const TYPE_MENU$ = "M"          '���j���[�\��


Private Type Sagyo_Code_tag
    
    CODE_TYPE       As String * 1       '��o�[�R�[�h�^�C�v
    YOIN_CODE       As String * 1       '�v��
    PARAM           As String * 2       '�p�����[�^

End Type

Private Type ID_KANRI_TBL_tag               '�q�@���̊Ǘ�
    RETRY                       As Integer              '�q�@���M���g���C
    ID                          As Integer              'IDNO
    Step                        As Integer              '�i���t���O
    JGYOBU                      As String * 1           '���ƕ�
    NAIGAI                      As String * 1           '�����O
    Hinban                      As String * 13          '�i�ԁi�Q���X�|���X�ȏ�̍�Ǝ��̂ݎg�p�j
    Tanaban                     As String * 8           '�I�ԁi�Q���X�|���X�ȏ�̍�Ǝ��̂ݎg�p�j
    GOODS_ON_F                  As String * 1           '���i���p�q��
    
    '---------------------------------------------------'���M����
    Send_SUMI_QTY               As Long                 '���i���ςݐ��ʁi�ړ����j
    Send_MI_QTY                 As Long                 '�����i���ʁi�ړ����j
    Send_Syuka_QTY              As Long                 '�o�א��ʁi�o�׎��j
    '---------------------------------------------------'�o�׏����p��
    MTS_CODE                    As String * 8           '���Ӑ�R�[�h
    SS_CODE                     As String * 8           '������R�[�h
    CYU_KBN                     As String * 1           '�����敪
    Y_SYU_CNT                   As Integer              '�Ώۓ`�[����
    ID_NO                       As String * 12          '�`�[ID
    DEN_NO                      As String * 6           '�`�[�ԍ�
    YUKO_SUMI_QTY               As Long                 '�g�p�\�ȏ��i���ςݍ݌�
    YUKO_MI_QTY                 As Long                 '�g�p�\�Ȗ����i�݌�
    SYUKA_QTY                   As Long                 '�o�א��ʁi�S���j
    
    '---------------------------------------------------'�o�׏����p��
'    MENU_GRP                    As String * 2           '�g�p���j���[
    
    
    MENU_LV1                    As String * 2           '���j���[���x���P   2006.01.30 3-->2
    MENU_LV2                    As String * 2           '���j���[���x���Q
'    MENU_LV3                    As String * 3           '���j���[���x���R  2006.01.30
    
    SAGYO_LOG                   As String * 1           '���۸ޏo�� 0:�Ȃ��@1:���� 2006.01.30
    
    
    PageNo_LV1                  As Integer              '�y�[�W���i���j���[�j
    PageNo_LV2                  As Integer              '�y�[�W���i���j���[�j
'    PageNo_LV3                  As Integer              '�y�[�W���i���j���[�j  2006.01.30
    Sagyo_Code                  As Sagyo_Code_tag       '��ƃR�[�h
    YOIN_DNAME                  As String * 5           '�\���p����
    TANTO_CODE                  As String * 5           '�S���҃R�[�h
    Recv_text(0 To M_Gyo - 1)   As String * 20          '�ŏI��M���e�P�`�T�s��
    Send_Text                   As Send_Text_tag        '�ŏI���M���e(����l)
    Last_Send_Text              As Send_Text_tag        '�ŏI���M���e(�S��)
    Time                        As String * 8           '���M����
    Last_Send                   As Integer              '0:�ʏ�e�L�X�g 1:�G���[���


    S_JGYOBU                    As String * 1           '���ނ𓥂܂��Ă̎��ƕ�
    S_NAIGAI                    As String * 1           '���ނ𓥂܂��Ă̍����O



End Type

Private ID_KANRI_TBL()      As ID_KANRI_TBL_tag

Private ING_No              As Integer  '�������̓Y��

Private Const Step_Start% = 0           '�q�@�d���n�m
Private Const Step_TANTO_REQ% = 1       '�S���җv��
Private Const Step_TANTO_RES% = 2       '�S���҉�

Private Const Step_JGYOBU_REQ% = 3      '���ƕ��v��
Private Const Step_JGYOBU_RES% = 4      '���ƕ���

Private Const Step_NAIGAI_REQ% = 5      '�����O�v��
Private Const Step_NAIGAI_RES% = 6      '�����O��


Private Const Step_MENU1_REQ% = 10      '���j���[�P�v��
Private Const Step_MENU1_RES% = 11      '���j���[�P��
Private Const Step_MENU2_REQ% = 12      '���j���[�Q�v��
Private Const Step_MENU2_RES% = 13      '���j���[�Q��
'2006.01.30 Private Const Step_MENU3_REQ% = 14      '���j���[�R�v��
'2006.01.30 Private Const Step_MENU3_RES% = 15      '���j���[�R��

Private Const Step_Sagyo1_REQ% = 20     '��ƂP�v��
Private Const Step_Sagyo1_RES% = 21     '��ƂP��
Private Const Step_Sagyo2_REQ% = 22     '��ƂQ�v��
Private Const Step_Sagyo2_RES% = 23     '��ƂQ��
Private Const Step_Sagyo3_REQ% = 24     '��ƂR�v��
Private Const Step_Sagyo3_RES% = 25     '��ƂR��
Private Const Step_Sagyo4_REQ% = 26     '��ƂS�v��
Private Const Step_Sagyo4_RES% = 27     '��ƂS��



Private Const BEF_Page$ = "$B"          '�O��
Private Const NEXT_Page$ = "$N"         '����


Private Type Menu_Tbl_tag               '���j���[���M�p�e�[�u��
    MENU_NO     As String * 2
    PARAM       As String * 16
    Disp        As String
    Log_Out     As String * 1

    
End Type



Private Type Wel_Para_Tag               'WELCAT���M�p�e�[�u��
    Box_Type    As String * 1
    LCD         As String * 10
    Keta        As Integer
End Type

Private Type WEL_Para_Tbl_tag
    Action      As String * 2
    Wel_Para(0 To M_Gyo - 1) As _
                Wel_Para_Tag
End Type
                                        '����Ɛ錾�̐��ɂ�葝���K�{�I�I
Private WEL_Para_Tbl(0 To 14, 0 To 9) As WEL_Para_Tbl_tag

Private Const LCD_Tanaban$ = "�I��"
Private Const LCD_Hinban$ = "�i��"
Private Const LCD_Suryo$ = "����"
Private Const LCD_Syuka$ = "�o�׎c"
Private Const LCD_SUMI_Suryo$ = "���i"
Private Const LCD_MI_Suryo$ = "�����i"
Private Const LCD_ID_No$ = "�`�[ID"
Private Const LCD_SYUKO_HYO_No$ = "�o�ɕ\��"
Private Const LCD_MTS$ = "������"


Private Const LCD_To_Tanaban$ = "�ړ���I��"



Private FILE_RETRY  As Integer          '�t�@�C���g�p�����̃��g���C��

Private Const Wel_TANAOROSI$ = "B0"         '�uWEL �I�����v�̗v��
Private Const Wel_TANAHYOJI$ = "B1"         '�uWEL �I�ԕ\���v�̗v��
Private Const Wel_HIN_SHOGO$ = "B2"         '�uWEL �i�ԕʏƍ��v�̗v��
Private Const Wel_AVE_SYUKA$ = "B3"         '�uWEL �����Ϗo�א��v�̗v��
Private Const Wel_HOST_ZAIKO$ = "B4"        '�uWEL �z�X�g�݌ɏƉ�v�̗v��
Private Const Wel_ST_TANABAN$ = "B5"        '�uWEL �W���I�Ԑݒ�v�̗v��
Private Const Wel_RIREKI$ = "B6"            '�uWEL �����o�ɗ����v�̗v��
Private Const Wel_SUII$ = "B7"              '�uWEL �o�א��ځv�̗v��
Private Const Wel_TANA_HIN_SHOGO$ = "B8"    '�uWEL �I�ԁE�i�ԕʏƍ��v�̗v��

Private Const Wel_TANAHYOJI_KASO$ = "B9"    '�uWEL �I�ԕ\��(���z�D��)�v�̗v��


Private Const Wel_GOODS_ONOFF_ONO$ = "D0"   '�uWEL ���i/�����i�؂�ւ��@����v�̗v��
Private Const Wel_GOODS_ONOFF_SIGA$ = "D1"  '�uWEL ���i/�����i�؂�ւ��@����v�̗v��


Private Const Wel_RETURNED_GOODS$ = "E0"    '�uWEL �Ǖi�ԕi�v�̗v��
Private Const Wel_LOCATION_MOVE$ = "E1"     '�uWEL �I�ړ��v�̗v��



Private B1_SendFile As String               '�uWEL �I�ԕ\���v�̑��M�t�@�C����
Private B6_SendFile As String               '�uWEL �����o�ɗ����v�̑��M�t�@�C����
Private B7_SendFile As String               '�uWEL �o�א��ځv�̑��M�t�@�C����

Private B9_SendFile As String               '�uWEL �I�ԕ\��(���z�D��)�v�̑��M�t�@�C����



Private Type SendFileRec_Tag                '���M�t�@�C�����R�[�h��`
    Title           As String * 1           '�^�C�g��
    LCD(0 To 19)    As Byte                 '�\�����b�Z�[�W
    CRLF            As String * 2           'CR/LF
End Type


Private Const Wel_Kbn_Title$ = "0"      '�^�C�g���s
Private Const Wel_Kbn_Normal$ = "1"     '�ʏ�\���s


Private Type Tanahyoji_tag              '�I�\���p�W�v�e�[�u��
    Tanaban         As String * 8
    SUMI_QTY        As Long
    MI_QTY          As Long
End Type

Private Inspection_Flg      As Integer  '���i�`�F�b�N�t���O(0:�o�ɖ���NG 1:�o�ɖ����ł�OK)
Private B2_MEMO     As String           '�i�ԕʍ݌ɏƍ��i�v����B2�j�̃���
Private B8_MEMO     As String           '�i�ԕʒI�ʍ݌ɏƍ��i�v����B8�j�̃���

Private ALL_MENU_GRP    As String * 2





Private Sub Command1_Click(Index As Integer)
'-------------------------------------------------------
'
'   �w�Ɩ��J�n�w���x
'       �P�D �|�[�g�̊l��
'       �Q�D �|�[�g�̊J��
'-------------------------------------------------------
    
Dim ans As Integer
    
    On Error GoTo Error
    
    Select Case Index
        
        Case 0                              '�Ɩ��J�n
            
            CtrsWsk1.Bind LocalPort, RemotePort
            F1100101.Caption = "�X�L���i����u���s���v"
    
            Command1(0).Enabled = False
            Command1(1).Enabled = True
            Command1(2).Enabled = True
    
        Case 1                              '�Ɩ��I��
            
            
            ans = MsgBox("�{���̋Ɩ��I�����܂����H", vbYesNo + vbDefaultButton2, "�Ɩ��I��")
            If ans = vbNo Then
                Exit Sub
            End If
            
            CtrsWsk1.Unbind
            
            Normal_End = False              '����I��
            
            Next_Step = 1                   '�������N������
            Unload Me
    
        Case 2
            CtrsWsk1.Unbind
            
            Normal_End = False              '����I��
            Next_Step = 0                   '�������N�����Ȃ�
            Unload Me
    
    End Select
    
    Exit Sub

Error:
    MsgBox "Winsock Error= " & Err.Description    '�X�e�[�^�X�s�ɃG���[��\�����܂��B
    
    Call Log_Out(LOG_F, "Winsock Error= " & Err.Description)
    
    Normal_End = True                       '�ُ�I��
    Unload Me
    


End Sub

Private Sub CtrsWsk1_OnReceive(ByVal ID_NO As Integer, ByVal RecvText As String, ByVal Resp_Mode As Boolean)
'-------------------------------------------------------
'
'   �w���R�[�h��M�������x
'
'-------------------------------------------------------

Dim nErrCode    As Integer
Dim strErrMsg   As String
Dim intLine     As Integer
Dim i           As Integer
Dim j           As Integer
Dim Chk_Time    As String * 8
Dim Sendbuf     As String

Dim Errbuf      As String

Dim sts         As Integer

Dim Start_Flg   As Integer


    Text1(0).Text = Format(ID_NO, "000") & ", Recv=" & RecvText
    
'    Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Recv=" & RecvText)
        
    RecvText = Left(RecvText, Len(RecvText) - 2)
    
                                    '�h�c���Ŏ�M�ς݃e�L�X�g����
    ING_No = -1
    
    Start_Flg = False
    
    For i = 0 To UBound(ID_KANRI_TBL)
        If ID_NO = ID_KANRI_TBL(i).ID Then
            ING_No = i
            Chk_Time = ID_KANRI_TBL(i).Time
            Exit For
        End If
    Next i
    
    
    
    
    
    If i > UBound(ID_KANRI_TBL) Then
                                                '�h�c���V�K�o�^
        For i = 0 To UBound(ID_KANRI_TBL)
            If ID_KANRI_TBL(i).ID = 0 Then
                
                Start_Flg = True
                
                ID_KANRI_TBL(i).ID = ID_NO      'ID_No  �ۑ�
                
'                ID_KANRI_TBL(i).MENU_GRP = ""
                ID_KANRI_TBL(i).MENU_LV1 = ""
                ID_KANRI_TBL(i).MENU_LV2 = ""
''                ID_KANRI_TBL(i).MENU_LV3 = ""
                
                If UBound(JGYOBU_T) = 0 Then    '�P���ƕ��Œ�
                Else
                    ID_KANRI_TBL(i).JGYOBU = ""
                End If
                
                If UBound(NAIGAI) = 0 Then   '�����O�Œ�
                Else
                    ID_KANRI_TBL(i).NAIGAI = ""
                End If
                
                ING_No = i
                Chk_Time = ""
                Exit For
            End If
        
        Next i
    End If
    
    
'Call Log_Out(LOG_F, Format(ID_NO, "000") & ",Yoin= " & ID_KANRI_TBL(i).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(i).Sagyo_Code.YOIN_CODE)
    
    
    If ING_No = -1 Then
        MsgBox "�h�m�h�t�@�C���̎q�@���̐ݒ��ύX���ĉ������B", vbCritical
        Normal_End = True
        Unload Me
    End If
    
                                            '�O���M�l���Ď�M�����H
    If Left(Right(RecvText, 9), 8) = ID_KANRI_TBL(i).Time And _
         Right(RecvText, 1) = "1" Then
            Call Send_Err_Proc(Sendbuf)
    
            Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf & "[�đ��M]")
    
    Else
                                            '��M���e��ۑ�
        ID_KANRI_TBL(ING_No).Recv_text(0) = Left(RecvText, 20)       '��M���e�P�s��
        ID_KANRI_TBL(ING_No).Recv_text(1) = Mid(RecvText, 21, 20)    '��M���e�Q�s��
        ID_KANRI_TBL(ING_No).Recv_text(2) = Mid(RecvText, 41, 20)    '��M���e�R�s��
        ID_KANRI_TBL(ING_No).Recv_text(3) = Mid(RecvText, 61, 20)    '��M���e�S�s��
        ID_KANRI_TBL(ING_No).Recv_text(4) = Mid(RecvText, 81, 20)    '��M���e�S�s��
        ID_KANRI_TBL(ING_No).Time = Right(RecvText, 8)               '���M����
        
        
                                            
        If Start_Flg Then
            If Trim(ID_KANRI_TBL(ING_No).Recv_text(0)) <> Start_Para Then
                Call Err_Send_Proc("�ċN�����Ă��������B", "", "", "", "")
                Sendbuf = Text_Create_Proc()
            End If
        End If
                                            
                                            
                                            '[START][CANCEL][FINISHI]��M�͏���������
        Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
            Case Start_Para     '�J�n(�q�@�d��ON)
                
                
                            '�o�ח\��^�݌ɂ̗\�����
                sts = Data_Clear_Proc(0, Sendbuf)
                Select Case sts
                    Case SYS_ERR
                        Normal_End = True
                End Select
                
                
                ID_KANRI_TBL(ING_No).Step = Step_Start
        
                Call Start_Proc(Sendbuf)
            
            
            Case Ent_Para       'ENT
                If Not Start_Flg Then
                    
                    If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN Or _
                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_MTS Then
                        '���i���̊m�F
                        
                        
                        
    '                    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    '                    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    '                    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    '                    Select Case sts
    '                        Case BtNoErr
    '                        '   -------------------------------- �G���[���b�Z�[�W�쐬
    '                        Case Else
    '                       '�d�v�ȗv���Ȃ̂Ŗ��o�^�̓V�X�e����~�Ƃ���
    '                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
    '                        Sendbuf = Text_Create_Proc()
    '                        Call File_Error(sts, BtOpGetEqual, "�v���}�X�^", 0)
    '                        Normal_End = True
    '                    End Select
    '
                        
                        
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                        
                        If Sagyo_Main_Proc(Sendbuf) Then
                            Normal_End = True
    '                        Unload Me
                        End If
                    Else
                        
                        '�Q�Ɖ�ʂ̊m�F���̂�
                        
                        
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
                            Normal_End = True
                        End Select
                       
                        
                        If Sagyo_Send_Proc() Then
                            Sendbuf = Text_Create_Proc()
                            Normal_End = True
                        End If
                        Sendbuf = Text_Create_Proc()
                    End If
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            Case Can_Para       'CANCEL
                If Not Start_Flg Then
                
                    If ID_KANRI_TBL(ING_No).Last_Send = 1 Then
                                
                                
                        '���i���̓f�[�^�̊J�����s���@2004.06.14 ��
                        If ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN Or _
                            ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = ACT_KENPIN_MTS Then
                            sts = Data_Clear_Proc(0, Sendbuf)
                            Select Case sts
                                Case SYS_CANCEL
                                    If Sagyo_Send_Proc() Then
                                        Sendbuf = Text_Create_Proc()
                                        Normal_End = True
                                    End If
                                    Sendbuf = Text_Create_Proc()
                                
                                Case SYS_ERR
                                    Normal_End = True
                            End Select
                        
                        
                        End If
                                
                        '���i���̓f�[�^�̊J�����s���@2004.06.14 ��
                                
                                '�O�񂪃G���[���M
                        Call Re_Send_Proc(Sendbuf)
                
                    Else
                                '�o�ח\��^�݌ɂ̗\�����
                        sts = Data_Clear_Proc(0, Sendbuf)
                        Select Case sts
                            Case SYS_CANCEL
                                If Sagyo_Send_Proc() Then
                                    Sendbuf = Text_Create_Proc()
                                    Normal_End = True
                                End If
                                Sendbuf = Text_Create_Proc()
                            
                            Case SYS_ERR
                                Normal_End = True
                        End Select
                
                
                
                        Call Cancel_Proc(Sendbuf)
                
                    End If
            
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            
            Case Fin_Para       'FINISH
                    
                If Not Start_Flg Then
                
                    '�o�ח\��^�݌ɂ̗\�����
                    sts = Data_Clear_Proc(0, Sendbuf)
                    Select Case sts
                        Case SYS_CANCEL
                            If Sagyo_Send_Proc() Then
                                Normal_End = True
                            End If
                            Sendbuf = Text_Create_Proc()
                    
                        Case SYS_ERR
                            Normal_End = True
                    End Select
                
                
'                    If Step_MENU1_REQ < ID_KANRI_TBL(ING_No).Step Then
                    If Step_TANTO_REQ <> ID_KANRI_TBL(ING_No).Step Then      '2005.01.07 if �`�@else �`�@end if
                    
                        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                        ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'2006.01.30                        ID_KANRI_TBL(ING_No).MENU_LV3 = ""
                
                        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
                        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
'2006.01.03                        ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
                
                
                
                        ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                
                        If Menu_Send_Proc(Sendbuf) Then
                            Normal_End = True
    '                Unload Me
                        End If
                
                    Else                                                    '2005.01.07
'                        ID_KANRI_TBL(ING_No).Step = Step_Start
                                                                            '2005.01.07
                        Call Start_Proc(Sendbuf)                            '2005.01.07
                                                                            '2005.01.07
                    End If                                                  '2005.01.07
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
            Case Else
                If Not Start_Flg Then
                                            '�i���`�F�b�N
                    Select Case ID_KANRI_TBL(ING_No).Step
            
            
                        Case Step_TANTO_REQ         '�S���җv���ɑ΂��郌�X
                        
                            
                            
                            ID_KANRI_TBL(ING_No).Step = Step_TANTO_RES
        
                            If Normal_Proc(Sendbuf) Then
                                Normal_End = True
    '                        Unload Me
                            End If
        
                        Case Step_JGYOBU_REQ        '���ƕ��v���ɑ΂��郌�X
                
                            Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
    '                        Case BEF_Page       '�O��
    '                        Case NEXT_Page      '����
                                Case Else            '���ƕ���M
                
                                    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_RES
              
                                    ID_KANRI_TBL(ING_No).JGYOBU = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
        
                                    If Normal_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
                                    
                        Case Step_NAIGAI_REQ
                    
                            Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                Case Else           '���j���[�p�����[�^��M
                
                                    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_RES
                                        
                                    ID_KANRI_TBL(ING_No).NAIGAI = Trim(ID_KANRI_TBL(i).Recv_text(0))
                                
        
                                    If Normal_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
    
'2006.01.30                        Case Step_MENU1_REQ, Step_MENU2_REQ, Step_MENU3_REQ
                        Case Step_MENU1_REQ, Step_MENU2_REQ
                
                             Select Case Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                Case BEF_Page       '�O��
                            
                                    
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV1 = ID_KANRI_TBL(ING_No).PageNo_LV1 - 1
                                        Case Step_MENU2_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = ID_KANRI_TBL(ING_No).PageNo_LV2 - 1
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = ID_KANRI_TBL(ING_No).PageNo_LV3 - 1
                                    End Select
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                    
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                            
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            
                                Case NEXT_Page      '����
                                    
                                    
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV1 = ID_KANRI_TBL(ING_No).PageNo_LV1 + 1
                                        Case Step_MENU2_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = ID_KANRI_TBL(ING_No).PageNo_LV2 + 1
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = ID_KANRI_TBL(ING_No).PageNo_LV3 + 1
                                    End Select
                                    
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                            
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                
                                Case Else           '���j���[�p�����[�^��M
                
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_REQ
                                            ID_KANRI_TBL(ING_No).PageNo_LV2 = 0
                                        Case Step_MENU2_REQ
'                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
'2006.01.30                                        Case Step_MENU3_REQ
'2006.01.30                                            ID_KANRI_TBL(ING_No).PageNo_LV3 = 0
                                    End Select
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                                
                                    Select Case ID_KANRI_TBL(ING_No).Step
                                        Case Step_MENU1_RES
                                            ID_KANRI_TBL(ING_No).MENU_LV1 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                        Case Step_MENU2_RES
                                            ID_KANRI_TBL(ING_No).MENU_LV2 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
 
                                            ID_KANRI_TBL(ING_No).MTS_CODE = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 3, 8)
                                            ID_KANRI_TBL(ING_No).SS_CODE = Mid(ID_KANRI_TBL(ING_No).Recv_text(0), 11, 8)
                            
                            
                            
                                            

'2006.01.30                                        Case Step_MENU3_RES
'2006.01.30                                            ID_KANRI_TBL(ING_No).MENU_LV3 = Trim(ID_KANRI_TBL(ING_No).Recv_text(0))
                                    End Select
                
                                
                                    If Menu_Recv_Proc(Sendbuf) Then
                                        Normal_End = True
    '                                Unload Me
                                    End If
                
                            End Select
                        
                        Case Step_Sagyo1_REQ, Step_Sagyo2_REQ, Step_Sagyo3_REQ, Step_Sagyo4_REQ
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step + 1
                        
                            If Sagyo_Main_Proc(Sendbuf) Then
                                Normal_End = True
    '                        Unload Me
                            End If
                        Case Else
                    End Select
                Else
                    ID_KANRI_TBL(i).ID = 0
                End If
    
        
        End Select
    
    End If
    
    

    If Resp_Mode Then
        On Error GoTo ShowError

        CtrsWsk1.SendResp Sendbuf

'        Text1(1).Text = Format(ID_NO, "000") & ", Send=" & SendBuf
        
'        Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf)

        On Error GoTo 0
    
        ID_KANRI_TBL(ING_No).Last_Send_Text.sts = Send_Text.sts                     '�X�e�[�^�X
        ID_KANRI_TBL(ING_No).Last_Send_Text.Display_Flg = Send_Text.Display_Flg     '�\����ʃt���O
        ID_KANRI_TBL(ING_No).Last_Send_Text.End_Menu = Send_Text.End_Menu           '�ŏI���j���[�t���O
        ID_KANRI_TBL(ING_No).Last_Send_Text.Menu_Suu = Send_Text.Menu_Suu           '���j���[��
        ID_KANRI_TBL(ING_No).Last_Send_Text.fileName = Send_Text.fileName           '�t�@�C����
        ID_KANRI_TBL(ING_No).Last_Send_Text.Buzzer = Send_Text.Buzzer               '�u�U�[�w��
        
        For j = 0 To M_Gyo - 1
                                                                                    'BOX����
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Box_Type = Send_Text.Box_Type(j).Box_Type
                                                                                    '�\�����e
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).LCD, StrConv(Send_Text.Box_Type(j).LCD, vbUnicode))
                                                                                    '�������e
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).INIT = Send_Text.Box_Type(j).INIT
                                                                                    '�J�n�J�[�\���ʒu
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Start_Pos = Send_Text.Box_Type(j).Start_Pos
                                                                                    '���͌����i�ő�j
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).Max_Size = Send_Text.Box_Type(j).Max_Size
                                                                                    '���j���[���e
            ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(j).MENU = Send_Text.Box_Type(j).MENU
                    
        
        Next j
    
    
        If Normal_End Then
            
'            MsgBox "�V�X�e���ُ킪�������܂����I�I���������Ă��������B"
 
            
'            Unload Me
        End If
    End If

    Exit Sub

ShowError:
    nErrCode = Err.Number
    strErrMsg = Err.Description         '�G���[���b�Z�[�W
    
    intLine = CtrsWsk1.ErrLineNo        '�ڑ��ԍ����擾���܂��B
    If intLine > 0 Then
        strErrMsg = strErrMsg & Chr(&HD) & Chr(&HA) & "�ڑ��ԍ� = " & intLine
    End If

    Text1(2).Text = strErrMsg           '�X�e�[�^�X�s�ɃG���[��\�����܂��B
    
'    Call Log_Out(LOG_F, Format(ID_NO, "000") & ", Send=" & Sendbuf)


End Sub

Private Sub Form_Load()
    
Dim c           As String * 128
Dim Out_Data    As String

Dim Box_Type    As String * 1
Dim LCD         As String * 10
Dim Keta        As String * 2

Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
    
Dim sts         As Integer
    
    Normal_End = False
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
'---------------------------------------------- '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LOG_F = RTrim(c)

'---------------------------------------------- '�f�[�^��M�|�[�g�ԍ���荞��
    If GetIni(App.EXEName, "LocalPort", "SYS", c) Then
        Beep
        MsgBox "�f�[�^��M�|�[�g�ԍ��̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LocalPort = CLng(RTrim(c))

'---------------------------------------------- '�f�[�^���M�|�[�g�ԍ���荞��
    If GetIni(App.EXEName, "RemotePort", "SYS", c) Then
        Beep
        MsgBox "�f�[�^���M�|�[�g�ԍ��̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    RemotePort = CLng(RTrim(c))
'---------------------------------------------- '���ƕ���荞��
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
'---------------------------------------------- '�����O����荞��
    i = 0
    
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI_CODE" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI(i - 1)
        NAIGAI(i - 1).CODE = Trim(c)
        If GetIni(App.EXEName, "NAIGAI_NAME" & Format(i, "0"), "SYS", c) Then
            MsgBox "�����O�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
            End
        End If
        NAIGAI(i - 1).NAME = Trim(c)
    
    Loop
    
    If i = 1 Then
        Beep
        MsgBox "�����O���̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
'---------------------------------------------- �O�؂���׏��l��
    If GetIni("YOIN", "YOIN_MAEGARI", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_MAEGARI] READ ERROR")
        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_MAEGARI = Trim(c)
'---------------------------------------------- '�����O�U�֏��l��
    If GetIni("YOIN", "YOIN_FURIKAE", "SYS", c) Then
    Else
        YOIN_FURIKAE = RTrim(c)
        '�����O�U�֐ݒ莞�A�ȉ��̍��ڕK�{
        If GetIni("YOIN", "YOIN_FURIKAE_OUT", "SYS", c) Then
            Beep
            MsgBox "�����O�U�֏��[YOIN_FURIKAE_OUT]�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
            End
        End If
    
        YOIN_FURIKAE_OUT = RTrim(c)
    
        If GetIni("YOIN", "YOIN_FURIKAE_IN", "SYS", c) Then
            Beep
            MsgBox "�����O�U�֏��[YOIN_FURIKAE_IN]�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
            End
        End If
    
        YOIN_FURIKAE_IN = RTrim(c)
    
    End If
'---------------------------------------------- �I�ƍ����l��
    If GetIni("YOIN", "YOIN_WEL_TANASHOGO", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_WEL_TANASHOGO] READ ERROR")
        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_TANASHOGO = Trim(c)

'---------------------------------------------- �I�i�ƍ����l��
    If GetIni("YOIN", "YOIN_WEL_TANAHINSHOGO", "SYS", c) Then
        YOIN_TANAHINSHOGO = Wel_TANA_HIN_SHOGO
    Else
        YOIN_TANAHINSHOGO = Trim(c)
    End If

'---------------------------------------------- '�q�@�䐔��荞��
    If GetIni(App.EXEName, "KO_SU", "SYS", c) Then
        Beep
        MsgBox "�q�@�䐔�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    ReDim ID_KANRI_TBL(0 To CInt(RTrim(c)) - 1)

    For i = 0 To UBound(ID_KANRI_TBL)
        ID_KANRI_TBL(i).ID = 0          'IDNo�N���A�[
        ID_KANRI_TBL(i).Step = 0        '�i���N���A�[
    
    Next i
'---------------------------------------------- '���M�p�p�����[�^��荞��
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
            WEL_Para_Tbl(i, j).Action = ""
        Next j
    Next i
    
    i = 0
    Do
        i = i + 1
        
        If GetIni("ACTION", "ACTION_CD" & Format(i, "00"), "SYS", c) Then
            Beep
            MsgBox "WELCAT���M�p�p���[���[�^([ACTION] [ACTION_CD])�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
            End
        End If
        If Trim(c) = "NON" Then
            Exit Do
        End If
    
    
        j = 0
    
        Do
            j = j + 1
            If GetIni("ACTION", "ACTION_WEL_PARA" & Format(i, "00") & Format(j, "00"), "SYS", c) Then
                Beep
                MsgBox "WELCAT���M�p�p���[���[�^([ACTION] [ACTION_WEL_PARA])�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
                End
            End If
            If Trim(c) = "NON" Then
                Exit Do
            End If
        
            Call Data_Select(Trim(c), 1, 14, Out_Data)
            
            WEL_Para_Tbl(i - 1, j - 1).Action = Trim(Out_Data)
        
            Call Data_Select(Trim(c), 2, 14, Out_Data)
            
            WEL_Para_Tbl(i - 1, j - 1).Wel_Para(0).Box_Type = Trim(Out_Data)
            WEL_Para_Tbl(i - 1, j - 1).Wel_Para(0).LCD = ""
        
        
            k = 2
            Do
                
                k = k + 1
                
                If k > 14 Then
                    Exit Do
                End If
                
                Call Data_Select(Trim(c), k, 14, Out_Data)
                Box_Type = Trim(Out_Data)
                
                k = k + 1
                Call Data_Select(Trim(c), k, 14, Out_Data)
                LCD = Trim(Out_Data)
            
                k = k + 1
                Call Data_Select(Trim(c), k, 14, Out_Data)
                Keta = Trim(Out_Data)
            
                Select Case k
                    Case 5
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(1).Keta = CInt(Keta)
                    Case 8
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(2).Keta = CInt(Keta)
                    Case 11
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(3).Keta = CInt(Keta)
                    Case 14
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).Box_Type = Box_Type
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).LCD = LCD
                        WEL_Para_Tbl(i - 1, j - 1).Wel_Para(4).Keta = CInt(Keta)
                
                End Select
            
            Loop
            
        
        Loop
    Loop
'---------------------------------------------- '��WELCAT�@����M���O�t�@�C����荞��
    
    If GetIni(App.EXEName, "LOG_F", "SYS", c) Then
        CtrsWsk1.LogFile = ""
    Else
        CtrsWsk1.LogFile = Trim(c)
    End If
'---------------------------------------------- '��WELCAT�@�f�[�^���M�p�t�H���_��荞��
    If GetIni(App.EXEName, "SendFolder", "SYS", c) Then
        Beep
        MsgBox "WELCAT���M�p�t�H���_([F110010] [SendFolder])�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    Else
        CtrsWsk1.SendFolder = Trim(c)
    End If
'---------------------------------------------- '��WELCAT�@�I�ԕ\���p�f�[�^�t�@�C������荞��
    If GetIni(App.EXEName, "B1", "SYS", c) Then
        Beep
        MsgBox "WELCAT���M�p�t�@�C��([F110010] [B1])�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    Else
        B1_SendFile = Trim(c)
    End If
'---------------------------------------------- '��WELCAT�@�o�ɗ���p�f�[�^�t�@�C������荞��
    If GetIni(App.EXEName, "B6", "SYS", c) Then
        Beep
        MsgBox "WELCAT���M�p�t�@�C��([F110010] [B6])�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    Else
        B6_SendFile = Trim(c)
    End If
'---------------------------------------------- '��WELCAT�@�o�א��ڗp�f�[�^�t�@�C������荞��
    If GetIni(App.EXEName, "B7", "SYS", c) Then
        Beep
        MsgBox "WELCAT���M�p�t�@�C��([F110010] [B7])�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    Else
        B7_SendFile = Trim(c)
    End If
'---------------------------------------------- '��WELCAT�@�I�ԕ\��(���z�D��)�p�f�[�^�t�@�C������荞��
    If GetIni(App.EXEName, "B9", "SYS", c) Then
        Beep
        MsgBox "WELCAT���M�p�t�@�C��([F110010] [B9])�̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    Else
        B9_SendFile = Trim(c)
    End If

'---------------------------------------------- '���ʃ��j���[����荞��
    If GetIni(App.EXEName, "ALL_MENU_GRP", "SYS", c) Then
        Beep
        MsgBox "���ʃ��j���[���̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If


    ALL_MENU_GRP = Trim(c)

'---------------------------------------------- '���i�`�F�b�N
    If GetIni(App.EXEName, "Inspection", "SYS", c) Then
        Inspection_Flg = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            Inspection_Flg = 1
        Else
            Inspection_Flg = CInt(Trim(c))
        End If
    End If
'---------------------------------------------- '�݌ɏƍ���������
    If GetIni(App.EXEName, "B2_MEMO", "SYS", c) Then
        B2_MEMO = ""
    Else
        B2_MEMO = Trim(c)
    End If
'--
    If GetIni(App.EXEName, "B8_MEMO", "SYS", c) Then
        B8_MEMO = ""
    Else
        B8_MEMO = Trim(c)
    End If
'---------------------------------------------- '�t�@�C�����g���C�񐔎�荞��
    If GetIni("SYSTEM", "RETRY", "SYS", c) Then
        FILE_RETRY = 1
    Else
        If Not IsNumeric(Trim(c)) Then
            FILE_RETRY = 1
        Else
            FILE_RETRY = CInt(Trim(c))
        End If
    End If
'---------------------------------------------- '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�i�ڃ}�X�^(���[�N)�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
'---------------------------------------------- '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '���j���[�Ǘ��}�X�^�n�o�d�m
    If P_MENU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�S���ҕʃ��j���[�n�o�d�m
    If P_TMENU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�݌Ƀf�[�^�i�ړ������p�j�n�o�d�m
    If wZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�݌Ƀf�[�^�i���i�^�����i�؂�ւ��p�j�n�o�d�m
    If tmpZAIKO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�O�؃f�[�^�n�o�d�m
    If J_NYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�o�ח\��f�[�^�n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�݌Ɉړ����f�[�^�n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '�����Ϗo�א��n�o�d�m
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '���ޑO���ް��n�o�d�m
    If P_NYU_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '��Ǝ���۸ނn�o�d�m
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Normal_End = True
        Unload Me
    End If
'---------------------------------------------- '���j���[�@�\�`�F�b�N�i�� or ���ʁj
'    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
'    sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
'    Select Case sts
'        Case BtNoErr
'            Menu_Type = 1           '���ʃ��j���[�ŉ^�p
'        Case BtErrKeyNotFound
            Menu_Type = 2           '�S���ҕʃ��j���[�ŉ^�p
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "�S���ҕʃ��j���[")
'            Unload Me
'    End Select
    
    


    Show

    If Data_Clear_Proc(1, "") Then
        MsgBox "�f�[�^�����ݒ肪�o���܂���ł����B"
        Unload Me
    End If


    If tmpZaiko_Clear_Proc() Then
        MsgBox "�f�[�^�����ݒ肪�o���܂���ł����B"
        Unload Me
    End If
End Sub

Private Sub Start_Proc(Sendbuf As String)
'-------------------------------------------------------
'
'   �w�q�@�J�n�����x
'
'-------------------------------------------------------
Dim i   As Integer
                                                '���M�e�L�X�g�쐬���Ǘ��e�[�u���ɓ]��
    ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ                      '�S���җv��
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                             '�\����ʃt���O �ʏ����
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = ""                                         '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = ""
    
    Send_Text.Menu_Suu = "05"                                       '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                         '���M�f�[�^�t�@�C����
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    '---------------------------------------------------------------
    Send_Text.Box_Type(0).Box_Type = TYPE_REF                       '�{�b�N�X�����@�\���a�n�w
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, "�S���ғ���")      '�\�����b�Z�[�W
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, "�S���ғ���")
                                                                    '�����\�����e�i���l�j
    Send_Text.Box_Type(0).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
    
    
    Send_Text.Box_Type(0).Start_Pos = ""                            '�J�[�\���ʒu
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
    
    Send_Text.Box_Type(0).Max_Size = "00"                           '�ő包��
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
    
    Send_Text.Box_Type(0).MENU = ""                                 '���j���[�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(1).Box_Type = TYPE_BCANK                     '�{�b�N�X�����@�\���a�n�w
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_BCANK
    
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "")                '�\�����b�Z�[�W
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "")
                                                                    '�����\�����e�i���l�j
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
    
    Send_Text.Box_Type(1).Start_Pos = "01"                          '�J�[�\���ʒu
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
    
    Send_Text.Box_Type(1).Max_Size = "05"                           '�ő包��
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "05"
    
    Send_Text.Box_Type(1).MENU = ""                                 '���j���[�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(2).Box_Type = TYPE_REF                       '�{�b�N�X�����@�\���a�n�w
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "")                '�\�����b�Z�[�W
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "")
                                                                    '�����\�����e�i���l�j
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
    
    Send_Text.Box_Type(2).Start_Pos = ""                            '�J�[�\���ʒu
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
    
    Send_Text.Box_Type(2).Max_Size = "00"                           '�ő包��
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
    
    Send_Text.Box_Type(2).MENU = ""                                 '���j���[�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(3).Box_Type = TYPE_REF                       '�{�b�N�X�����@�\���a�n�w
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "")                '�\�����b�Z�[�W
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "")
                                                                    '�����\�����e�i���l�j
    Send_Text.Box_Type(3).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
    
    Send_Text.Box_Type(3).Start_Pos = ""                            '�J�[�\���ʒu
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = ""
    
    Send_Text.Box_Type(3).Max_Size = "00"                           '�ő包��
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
    
    Send_Text.Box_Type(3).MENU = ""                                 '���j���[�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
    '---------------------------------------------------------------
    Send_Text.Box_Type(4).Box_Type = TYPE_REF                       '�{�b�N�X�����@�\���a�n�w
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
    
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")                '�\�����b�Z�[�W
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                                    '�����\�����e�i���l�j
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
    
    Send_Text.Box_Type(4).Start_Pos = ""                            '�J�[�\���ʒu
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = ""
    
    Send_Text.Box_Type(4).Max_Size = "00"                           '�ő包��
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "00"
    
    Send_Text.Box_Type(4).MENU = ""                                 '���j���[�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
    '---------------------------------------------------------------
    Send_Text.CRLF = vbCrLf
    '------------------------------------------ ���M�o�b�t�@�[�֓]��
    Sendbuf = Text_Create_Proc()

    ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M

End Sub


Private Function Normal_Proc(Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   �w�ʏ�e�L�X�g��M�x
'
'-------------------------------------------------------
    Normal_Proc = True
    
    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Start                 '�q�@�d��ON�i�����ɂ͗��Ȃ��j
        
        Case Step_TANTO_REQ             '�S���җv���i�����ɂ͗��Ȃ��j
        
        Case Step_TANTO_RES             '�S���҉�
    
            If Tanto_Check_Proc(Sendbuf) Then
                Exit Function
            End If
    
        Case Step_JGYOBU_REQ            '���ƕ��v���i�����ɂ͗��Ȃ��j
                
        Case Step_JGYOBU_RES            '���ƕ���
            
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
                
        Case Step_NAIGAI_REQ            '�����O�v���i�����ɂ͗��Ȃ��j
                
        Case Step_NAIGAI_RES            '�����O��
            
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
                                        '���j���[�v���i�����ɂ͗��Ȃ��j
'2006.01.30        Case Step_MENU1_REQ, Step_MENU2_REQ, Step_MENU3_REQ
        Case Step_MENU1_REQ, Step_MENU2_REQ
                                        '���j���[��
'2006.01.30        Case Step_MENU1_RES, Step_MENU2_RES, Step_MENU3_RES
        Case Step_MENU1_RES, Step_MENU2_RES
    
            If Menu_Send_Proc(Sendbuf) Then
                Exit Function
            End If
    
    End Select
    
    Normal_Proc = False

End Function

Private Sub Form_Unload(Cancel As Integer)

Dim sts As Integer

'---------------------------------------------- '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
'---------------------------------------------- '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
        End If
    End If
'---------------------------------------------- '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
'---------------------------------------------- '�i�ڃ}�X�^�i���[�N�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
'---------------------------------------------- '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
'---------------------------------------------- '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
'---------------------------------------------- '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
'---------------------------------------------- '���j���[�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_MENU_POS, P_MENUREC, Len(P_MENUREC), K0_P_MENU, Len(K0_P_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���j���[�Ǘ��}�X�^")
        End If
    End If
'---------------------------------------------- '���ԃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ԃ}�X�^")
        End If
    End If
'---------------------------------------------- '�S���ҕʃ��j���[�b�k�n�r�d
    sts = BTRV(BtOpClose, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���ҕʃ��j���[")
        End If
    End If
'---------------------------------------------- '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
'---------------------------------------------- '�݌Ƀf�[�^�i�ړ������p�j�b�k�n�r�d

    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
'---------------------------------------------- '�O�؃f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�O�؃f�[�^")
        End If
    End If
'---------------------------------------------- '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
'---------------------------------------------- '�݌Ɉړ����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ����f�[�^")
        End If
    End If
'---------------------------------------------- '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
'---------------------------------------------- '�����Ϗo�א��b�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo�א�")
        End If
    End If
'---------------------------------------------- '���ޑO���ް��b�k�n�r�d
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޑO���ް�")
        End If
    End If
'---------------------------------------------- '��Ǝ���۸ނb�k�n�r�d
    sts = BTRV(BtOpClose, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޑO���ް�")
        End If
    End If
'---------------------------------------------- '�t�@�C�������Z�b�g
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If



    If Next_Step = 1 Then
        sts = Shell("d:\newsdc\exe\F1100501.bat", vbNormalFocus)
        If sts = 0 Then
            MsgBox "[F110050]�I�������̋N���Ɏ��s���܂����B "
            Call Log_Out(LOG_F, "[F110050]�I�������̋N���Ɏ��s���܂����B")
        End If
    End If


    Set F1100101 = Nothing


    


    End
End Sub
Private Function Tanto_Check_Proc(Sendbuf As String) As Integer
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
Private Sub Err_Send_Proc(Errmsg0 As String, _
                            Errmsg1 As String, _
                            Errmsg2 As String, _
                            Errmsg3 As String, _
                            Errmsg4 As String)
'-------------------------------------------------------
'
'   �w�G���[���b�Z�[�W�d���̍쐬�x
'
'-------------------------------------------------------
    
    Send_Text.sts = Sts_NG                  '�X�e�[�^�X
    Send_Text.Display_Flg = Display_DEF     '�\����ʃt���O
    Send_Text.End_Menu = ""                 '�ŏI���j���[�t���O
    Send_Text.Menu_Suu = ""                 '���j���[��
    Send_Text.fileName = ""                 '�t�@�C����
    Send_Text.Buzzer = Buzzer_CONTI         '�u�U�[��
'-------------------------------------------------------
    Send_Text.Box_Type(0).Box_Type = TYPE_REF               '�s�P�@ BOX����
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, Errmsg0)   '       �\�����e
    Send_Text.Box_Type(0).INIT = ""                         '       ���l�����l
    Send_Text.Box_Type(0).Start_Pos = ""                    '       �J�n�ʒu
    Send_Text.Box_Type(0).Max_Size = ""                     '       ���͌���
    Send_Text.Box_Type(0).MENU = ""                         '       ���j���[���e
'-------------------------------------------------------
    Send_Text.Box_Type(1).Box_Type = TYPE_REF               '�s�Q   BOX����
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Errmsg1)   '       �\�����e
    Send_Text.Box_Type(1).INIT = ""                         '       ���l�����l
    Send_Text.Box_Type(1).Start_Pos = ""                    '       �J�n�ʒu
    Send_Text.Box_Type(1).Max_Size = ""                     '       ���͌���
    Send_Text.Box_Type(1).MENU = ""                         '       ���j���[���e
'-------------------------------------------------------
    Send_Text.Box_Type(2).Box_Type = TYPE_REF               '�s�R�@ BOX����
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, Errmsg2)   '       �\�����e
    Send_Text.Box_Type(2).INIT = ""                         '       ���l�����l
    Send_Text.Box_Type(2).Start_Pos = ""                    '       �J�n�ʒu
    Send_Text.Box_Type(2).Max_Size = ""                     '       ���͌���
    Send_Text.Box_Type(2).MENU = ""                         '       ���j���[���e
'-------------------------------------------------------
    Send_Text.Box_Type(3).Box_Type = TYPE_REF               '�s�S�@ BOX����
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, Errmsg3)   '       �\�����e
    Send_Text.Box_Type(3).INIT = ""                         '       ���l�����l
    Send_Text.Box_Type(3).Start_Pos = ""                    '       �J�n�ʒu
    Send_Text.Box_Type(3).Max_Size = ""                     '       ���͌���
    Send_Text.Box_Type(3).MENU = ""                         '       ���j���[���e
'-------------------------------------------------------
    Send_Text.Box_Type(4).Box_Type = TYPE_REF               '�s�T�@ BOX����
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Errmsg4)   '       �\�����e
    Send_Text.Box_Type(4).INIT = ""                         '       ���l�����l
    Send_Text.Box_Type(4).Start_Pos = ""                    '       �J�n�ʒu
    Send_Text.Box_Type(4).Max_Size = ""                     '       ���͌���
    Send_Text.Box_Type(4).MENU = ""                         '       ���j���[���e
'-------------------------------------------------------
    Send_Text.CRLF = vbCrLf
        
    ID_KANRI_TBL(ING_No).Last_Send = 1                      '�G���[���M

End Sub
Private Function Text_Create_Proc() As String
'-------------------------------------------------------
'
'   �w���M�e�L�X�g�쐬�x
'
'-------------------------------------------------------
Dim i   As Integer
    
    Text_Create_Proc = Send_Text.sts & _
                Send_Text.Display_Flg & _
                Send_Text.End_Menu & _
                Send_Text.Menu_Suu & _
                Send_Text.fileName & _
                Send_Text.Buzzer

    For i = 0 To 4
        Text_Create_Proc = Text_Create_Proc & Send_Text.Box_Type(i).Box_Type & _
                            StrConv(Send_Text.Box_Type(i).LCD, vbUnicode) & _
                            Send_Text.Box_Type(i).INIT & _
                            Send_Text.Box_Type(i).Start_Pos & _
                            Send_Text.Box_Type(i).Max_Size & _
                            Send_Text.Box_Type(i).MENU
    Next i
    
    Text_Create_Proc = Text_Create_Proc & Send_Text.CRLF

End Function
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

Private Function Menu_Send_Proc(Optional Sendbuf As String) As Integer

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


    Menu_Send_Proc = True
'----------------------------------------------- '���ƕ��I������
    If Trim(ID_KANRI_TBL(ING_No).JGYOBU) = "" Then
        Call JGYOBU_MENU_SET

        Sendbuf = Text_Create_Proc()


        Menu_Send_Proc = False
        Exit Function
    End If
'----------------------------------------------- '�����O�I������
    If Trim(ID_KANRI_TBL(ING_No).NAIGAI) = " " Then

        ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        ID_KANRI_TBL(ING_No).MENU_LV2 = ""

        ID_KANRI_TBL(ING_No).PageNo_LV1 = 0
        ID_KANRI_TBL(ING_No).PageNo_LV2 = 0


        Call NAIGAI_MENU_SET
        Sendbuf = Text_Create_Proc
        Menu_Send_Proc = False
        Exit Function
    End If
    '   -------------------------------- ���x���P�@�g�b�v���j���[�̊Ǘ�
    If Trim(ID_KANRI_TBL(ING_No).MENU_LV1) = "" Then
        '�ƭ���ٰ��
        Call UniCode_Conv(K0_P_TMENU.TANTO_CODE, ID_KANRI_TBL(ING_No).TANTO_CODE)
        Erase Menu_Tbl

        sts = BTRV(BtOpGetEqual, P_TMENU_POS, P_TMENUREC, Len(P_TMENUREC), K0_P_TMENU, Len(K0_P_TMENU), 0)
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
        For i = 0 To 29
            If Trim(StrConv(P_TMENUREC.MENU_T(i).MENU_NO, vbUnicode)) = "" Then
                Exit For
            End If
        
            If StrConv(P_TMENUREC.MENU_T(i).JGYOBU, vbUnicode) = ID_KANRI_TBL(ING_No).JGYOBU Or _
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
            If UBound(NAIGAI) = 0 Then
                ID_KANRI_TBL(ING_No).Step = Step_Start
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
            End If
            Menu_Send_Proc = False
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
                Send_Text.End_Menu = Menu_End       '�ŏI�y�[�W
                ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_End
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
        Send_Text.fileName = ""                                         '���M�f�[�^�t�@�C����
        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
        Send_Text.Buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
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
            
            
            
            If P_SAGYO_LOG_OUTPUT_PROC(ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                        ID_KANRI_TBL(ING_No).JGYOBU, _
                                        ID_KANRI_TBL(ING_No).NAIGAI, _
                                        ID_KANRI_TBL(ING_No).MENU_LV1, _
                                        "ST", , , , , , , , , FILE_RETRY) Then
                            
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            
            End If
            
            
            
            
            
            
            
            
            
            
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
            
                Menu_Tbl(Menu_Cnt).Log_Out = StrConv(P_MENUREC.SAGYO(i).Log_Out, vbUnicode)
            
            
            
            
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
                    Send_Text.End_Menu = Menu_End       '�ŏI�y�[�W
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_End
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
            Send_Text.fileName = ""                                         '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
            Send_Text.Buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
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
                                                                        '���j���\�ԍ� & ���Ұ�
                    Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM
                    
''''2006.05.31                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Menu_Tbl(i).PARAM
                
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(j).MENU = Menu_Tbl(i).MENU_NO & Trim(CStr(Dec_To_Bcd(Menu_Tbl(i).PARAM)))
                
                
                
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


Private Sub JGYOBU_MENU_SET()
'-------------------------------------------------------
'
'   �w���ƕ��I��p���j���[�쐬�x
'
'-------------------------------------------------------
Dim i   As Integer

    ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ                     '���ƕ��v��
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_MENU                            '�\����ʃt���O ���j���[���
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
    
    Send_Text.End_Menu = Menu_Only                                  '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = Format(UBound(JGYOBU_T) + 1, "00")         '���j���[���ڐ�
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(UBound(JGYOBU_T) + 1, "00")
    
    Send_Text.fileName = ""                                         '���M�f�[�^�t�@�C����
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    
    '---------------------------------------------------------------
    For i = 0 To M_Gyo - 1
        
        If i > UBound(JGYOBU_T) Then
        
            Send_Text.Box_Type(i).Box_Type = ""                 'BOX����
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
        
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")    '�\�����e
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
            
            Send_Text.Box_Type(i).INIT = ""                     '���l�����l
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                                                
            Send_Text.Box_Type(i).Start_Pos = ""                '�����J�[�\���ʒu
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '���͌���
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
        
        Else
            
            Send_Text.Box_Type(i).Box_Type = TYPE_MENU          'BOX����
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = TYPE_MENU
                                                                '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, JGYOBU_T(i).NAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, JGYOBU_T(i).NAME)
                                                                                
            Send_Text.Box_Type(i).INIT = ""                     '���l�����l
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
            
            Send_Text.Box_Type(i).Start_Pos = ""                '�����J�[�\���ʒu
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '���͌���
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = JGYOBU_T(i).CODE       '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = JGYOBU_T(i).CODE
                                                                                
        
        End If
    
    Next i
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M

End Sub

Private Sub NAIGAI_MENU_SET()
'-------------------------------------------------------
'
'   �w���O�I��p���j���[�쐬�x
'
'-------------------------------------------------------
Dim i   As Integer

    ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ                     '���O�v��
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_MENU                             '�\����ʃt���O ���j���[���
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_MENU
    
    Send_Text.End_Menu = Menu_Only                                  '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = Format(UBound(NAIGAI) + 1, "00")         '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = Format(UBound(NAIGAI) + 1, "00")
    
    Send_Text.fileName = ""                                         '���M�f�[�^�t�@�C����
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    
    '---------------------------------------------------------------
    For i = 0 To M_Gyo - 1
        
        If i > UBound(NAIGAI) Then
        
            Send_Text.Box_Type(i).Box_Type = ""                 'BOX����
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
        
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")    '�\�����e
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
            
            Send_Text.Box_Type(i).INIT = ""                     '���l�����l
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                                                
            Send_Text.Box_Type(i).Start_Pos = ""                '�����J�[�\���ʒu
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '���͌���
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
        
        Else
            
            Send_Text.Box_Type(i).Box_Type = TYPE_MENU          'BOX����
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = TYPE_MENU
                                                                '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(i).LCD, NAIGAI(i).NAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, NAIGAI(i).NAME)
                                                                                
            Send_Text.Box_Type(i).INIT = ""                     '���l�����l
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                                                
            Send_Text.Box_Type(i).Start_Pos = ""                '�����J�[�\���ʒu
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                                                
            Send_Text.Box_Type(i).Max_Size = "00"               '���͌���
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(i).MENU = NAIGAI(i).CODE       '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = NAIGAI(i).CODE
                                                                                
        
        End If
    
    Next i
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M

End Sub

Private Sub Re_Send_Proc(Sendbuf As String)
'-------------------------------------------------------
'
'   �w�G���[���̍đ��x
'
'-------------------------------------------------------
Dim i   As Integer
    
    
    
'    Select Case ID_KANRI_TBL(ING_No).Step
'        Case 2, 4, 6
'            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'    End Select
'-------------------------------------------------------
    Send_Text.sts = ID_KANRI_TBL(ING_No).Send_Text.sts                  '�X�e�[�^�X�@OK
    
    Send_Text.Display_Flg = ID_KANRI_TBL(ING_No).Send_Text.Display_Flg  '�\����ʃt���O ���j���[���
    
    Send_Text.End_Menu = ID_KANRI_TBL(ING_No).Send_Text.End_Menu        '�ŏI���j���[�t���O
    
    Send_Text.Menu_Suu = ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu        '���j���[���ڐ��i05�Œ�j
    
    Send_Text.fileName = ID_KANRI_TBL(ING_No).Send_Text.fileName        '���M�f�[�^�t�@�C����
    
    Send_Text.Buzzer = ID_KANRI_TBL(ING_No).Send_Text.Buzzer            '�u�U�[���@�W��

    For i = 0 To M_Gyo - 1
                                                                        'BOX����
        Send_Text.Box_Type(i).Box_Type = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type
                                                                        '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode))
                                                                        
                                                                    '�����\�����e�i���l�j
        Send_Text.Box_Type(i).INIT = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT
                                                                        '�����J�[�\���ʒu
        Send_Text.Box_Type(i).Start_Pos = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos
                                                                        '���͌���
        Send_Text.Box_Type(i).Max_Size = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size
                                                                        '���j���\�ԍ�
        Send_Text.Box_Type(i).MENU = ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU

    Next
    
    Send_Text.CRLF = vbCrLf
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M
    
    Sendbuf = Text_Create_Proc()


End Sub
Private Sub Send_Err_Proc(Sendbuf As String)
'-------------------------------------------------------
'
'   �w���M�G���[���̍đ��x
'
'-------------------------------------------------------
Dim i   As Integer
    
    
    
'    Select Case ID_KANRI_TBL(ING_No).Step
'        Case 2, 4, 6
'            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'    End Select
'-------------------------------------------------------
    Send_Text.sts = ID_KANRI_TBL(ING_No).Last_Send_Text.sts                  '�X�e�[�^�X�@OK
    
    Send_Text.Display_Flg = ID_KANRI_TBL(ING_No).Last_Send_Text.Display_Flg  '�\����ʃt���O ���j���[���
    
    Send_Text.End_Menu = ID_KANRI_TBL(ING_No).Last_Send_Text.End_Menu        '�ŏI���j���[�t���O
    
    Send_Text.Menu_Suu = ID_KANRI_TBL(ING_No).Last_Send_Text.Menu_Suu        '���j���[���ڐ��i05�Œ�j
    
    Send_Text.fileName = ID_KANRI_TBL(ING_No).Last_Send_Text.fileName        '���M�f�[�^�t�@�C����
    
    Send_Text.Buzzer = ID_KANRI_TBL(ING_No).Last_Send_Text.Buzzer            '�u�U�[���@�W��

    For i = 0 To M_Gyo - 1
                                                                        'BOX����
        Send_Text.Box_Type(i).Box_Type = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).Box_Type
                                                                        '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, StrConv(ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).LCD, vbUnicode))
                                                                        
                                                                    '�����\�����e�i���l�j
        Send_Text.Box_Type(i).INIT = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).INIT
                                                                        '�����J�[�\���ʒu
        Send_Text.Box_Type(i).Start_Pos = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).Start_Pos
                                                                        '���͌���
        Send_Text.Box_Type(i).Max_Size = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).Max_Size
                                                                        '���j���\�ԍ�
        Send_Text.Box_Type(i).MENU = ID_KANRI_TBL(ING_No).Last_Send_Text.Box_Type(i).MENU

    Next
    
    Send_Text.CRLF = vbCrLf
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M
    
    Sendbuf = Text_Create_Proc()


End Sub

Private Function Menu_Recv_Proc(Sendbuf As String) As Integer
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
                                Left(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 7) = (ID_KANRI_TBL(ING_No).MTS_CODE & _
                                                                                        ID_KANRI_TBL(ING_No).SS_CODE) Then
                                
                                Call UniCode_Conv(K0_YOIN.CODE_TYPE, Left(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                Call UniCode_Conv(K0_YOIN.YOIN_CODE, Right(StrConv(P_MENUREC.SAGYO(i).YOIN, vbUnicode), 1))
                                
                                sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
                                Select Case sts
                                
                                    Case BtNoErr
                                        '���ŕ\������
                                        ID_KANRI_TBL(ING_No).YOIN_DNAME = StrConv(P_MENUREC.SAGYO(i).Disp, vbUnicode)
                                
                                        
                                        ID_KANRI_TBL(ING_No).SAGYO_LOG = StrConv(P_MENUREC.SAGYO(i).Log_Out, vbUnicode)
                                                                    
                                        If StrConv(YOINREC.PARAM_F, vbUnicode) = "1" Then   '������Ȃ�i�o�ׁj
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
'2006.01.30                                            ID_KANRI_TBL(ING_No).MTS_CODE = Left(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
'2006.01.30                                            ID_KANRI_TBL(ING_No).SS_CODE = Right(StrConv(P_MENUREC.SAGYO(i).PARAM, vbUnicode), 8)
                                        End If
                                                                                            '���i�i������w��j�Ȃ�
                                        If StrConv(YOINREC.CODE_TYPE, vbUnicode) = ACT_KENPIN_MTS Then
                                            ID_KANRI_TBL(ING_No).CYU_KBN = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        End If
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(YOINREC.CODE_TYPE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(YOINREC.YOIN_CODE, vbUnicode)
                                        ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(YOINREC.Soko_No, vbUnicode)
                                
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




Private Function Sagyo_Send_Proc() As Integer
'-------------------------------------------------------
'
'   �w��Ƃ̑��M�x
'
'-------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer

Dim Found_Flg   As Boolean

Dim sts         As Integer

    Sagyo_Send_Proc = True
    
                        '�v���̓ǂݍ���
    Call UniCode_Conv(K0_YOIN.CODE_TYPE, ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE)
    Call UniCode_Conv(K0_YOIN.YOIN_CODE, ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE)
    sts = BTRV(BtOpGetEqual, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    
    Select Case sts
        Case BtNoErr
        '   -------------------------------- �G���[���b�Z�[�W�쐬
        Case Else
        '�u�v�����o�^�͍l�����Ȃ��G���[�V�X�e����~�Ƃ���v
            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                Sendbuf = Text_Create_Proc()
            Call File_Error(sts, BtOpGetEqual, "�v���}�X�^", 0)
            Exit Function
    End Select
    
    
    
    '   -------------------------------- ���M�p�����[�^�̌���
    Found_Flg = False
    For i = 0 To UBound(WEL_Para_Tbl, 1)
        For j = 0 To UBound(WEL_Para_Tbl, 2)
                                        '�ŏ��͂Q���Ō���
            If StrConv(YOINREC.CODE_TYPE, vbUnicode) & StrConv(YOINREC.YOIN_CODE, vbUnicode) = _
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
               If StrConv(YOINREC.CODE_TYPE, vbUnicode) = Left(WEL_Para_Tbl(i, j).Action, 1) Then
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
        '�M�����Ȃ��G���[
        Call Log_Out(LOG_F, "�v���}�X�^����WLECAT�p�����[�^�iINI�t�@�C���j")
        Exit Function
    End If

    '   -------------------------------- ��ƍ쐬
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ                     '�ʏ��ƊJ�n
    '---------------------------------------------------------------
    Send_Text.sts = Sts_OK                                          '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                             '�\����ʃt���O �ʏ���͉��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                                  '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                                       '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                         '���M�f�[�^�t�@�C����
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                                   '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
    
                                                                    '�p�����[�^��1�s�ڂɂ͗v�����̂��Z�b�g
    WEL_Para_Tbl(i, j).Wel_Para(0).LCD = StrConv(YOINREC.YOIN_DNAME, vbUnicode)
    '---------------------------------------------------------------
    For k = 0 To M_Gyo - 1
        
                                                            'BOX����
        Send_Text.Box_Type(k).Box_Type = WEL_Para_Tbl(i, j).Wel_Para(k).Box_Type
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Box_Type = WEL_Para_Tbl(i, j).Wel_Para(k).Box_Type
                                                            
                                                            '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(k).LCD, WEL_Para_Tbl(i, j).Wel_Para(k).LCD)
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).LCD, WEL_Para_Tbl(i, j).Wel_Para(k).LCD)
                                                            '���l�����\��
        Send_Text.Box_Type(k).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).INIT = ""
                                                            
                                                            '�����J�[�\���ʒu
        If Send_Text.Box_Type(k).Box_Type = "2" Then
            Send_Text.Box_Type(k).Start_Pos = Format(M_Keta - WEL_Para_Tbl(i, j).Wel_Para(k).Keta + 1, "00")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Start_Pos = Format(M_Keta - WEL_Para_Tbl(i, j).Wel_Para(k).Keta, "00")
        Else
            Send_Text.Box_Type(k).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Start_Pos = "01"
        End If
                                                            '���͌���
        Send_Text.Box_Type(k).Max_Size = Format(WEL_Para_Tbl(i, j).Wel_Para(k).Keta, "00")
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).Max_Size = Format(WEL_Para_Tbl(i, j).Wel_Para(k).Keta, "00")
                                                                                
        Send_Text.Box_Type(k).MENU = ""                     '���j���\�ԍ�
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(k).MENU = ""
                                                                                
    
    Next k
    
'    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE = StrConv(MENUREC.CODE_TYPE, vbUnicode)
'    ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE = StrConv(MENUREC.YOIN_CODE, vbUnicode)
    ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM = StrConv(YOINREC.Soko_No, vbUnicode)
    
    ID_KANRI_TBL(ING_No).Last_Send = 0  '�m�[�}���f�[�^���M

    Sagyo_Send_Proc = False


End Function

Private Function Sagyo_Main_Proc(Sendbuf As String) As Integer
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
            
            If Zaitei_Out_Proc(Sendbuf, i, j) Then
                Exit Function
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
        
            If Ido_In_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_IDO_OUT            '�ړ��o��
        
            If Ido_Out_Proc(Sendbuf, i, j) Then
                Exit Function
            End If
        
        
        Case ACT_DENPYO_ID          '�`�[�h�c
        
            If DEN_ID_Dec_Proc(Sendbuf, i, j) Then
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
                
            
            End Select
        
    
    End Select

    Sagyo_Main_Proc = False

End Function

Private Function Zaitei_In_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�ݒ��i�{�j�w�莞�̃`�F�b�N���X�V�����x
'
'-------------------------------------------------------
Dim i           As Integer
Dim Hinban      As String * 13
Dim Tanaban     As String * 8
Dim sts         As Integer

Dim QTY         As Long
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1

Dim MENU_NO     As String * 2

    Zaitei_In_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Tanaban        '�I��
                Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                
                If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                Else
                '------------------ �q�Ƀ}�X�^�Ǎ���
                    Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                            Zaitei_In_Proc = False
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
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Zaitei_In_Proc = False
                            Exit Function
                        End If
                    End If
                    '------------------ �I�}�X�^�Ǎ���
                    Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                    Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                    Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                    Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Zaitei_In_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                            Exit Function
                    End Select
            
                    '------------------ �֎~�I�̃`�F�b�N
                    If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                        Zaitei_In_Proc = False
                        Exit Function
                    End If
            
            
                End If
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ �i�ڃ}�X�^�Ǎ���
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Zaitei_In_Proc = False
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
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Zaitei_In_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                End Select
            
            Case LCD_Suryo          '���ʁi�����͖����j
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_In_Proc = False
                    Exit Function
                End If
                
                QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                If QTY = 0 Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_In_Proc = False
                    Exit Function
                End If
            
            
            Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_In_Proc = False
                    Exit Function
                End If
                
                
                If Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD) = LCD_SUMI_Suryo Then
                    SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                Else
                    MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                End If
        
                If i = M_Gyo - 1 Then
                    If SUMI_QTY = 0 And MI_QTY = 0 Then
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "���i�^�����i���O", "���ʓ��̓~�X", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                        Zaitei_In_Proc = False
                        Exit Function
                    End If
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
    
    
    
    
        sts = Nyuko_Update_Proc(RET_JGYOBU, _
                                RET_NAIGAI, _
                                Hinban, _
                                Format(Now, "YYYYMMDD"), _
                                Tanaban, _
                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                SUMI_QTY, _
                                MI_QTY, _
                                Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                FILE_RETRY, , _
                                StrConv(ITEMREC.G_SHIIRE_TBL(0).CODE, vbUnicode), _
                                StrConv(ITEMREC.G_SHIIRE_TBL(0).TANKA, vbUnicode), , _
                                MENU_NO)
 
        Select Case sts
            Case False
            Case True           '���Ɏ��͔������Ȃ�
            Case SYS_CANCEL
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                Zaitei_In_Proc = False
                GoTo Abort_Tran
            Case SYS_ERR
                Sendbuf = Text_Create_Proc()
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Zaitei_In_Proc = SYS_ERR    '�V�X�e���ُ픭��
                
                GoTo Abort_Tran
        End Select
    
    
    
    
    
    
    Else
                                        
                                            
                                            '���ɍX�V
        sts = Nyuko_Update_Proc(RET_JGYOBU, _
                                RET_NAIGAI, _
                                Hinban, _
                                Format(Now, "YYYYMMDD"), _
                                Tanaban, _
                                (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                SUMI_QTY, _
                                MI_QTY, _
                                Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                FILE_RETRY, , , , , MENU_NO)
        Select Case sts
            Case False
            Case True           '���Ɏ��͔������Ȃ�
            Case SYS_CANCEL
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                Zaitei_In_Proc = False
                GoTo Abort_Tran
            Case SYS_ERR
                Sendbuf = Text_Create_Proc()
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Zaitei_In_Proc = SYS_ERR    '�V�X�e���ُ픭��
                
                GoTo Abort_Tran
        End Select
    End If

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
    
    
    
    Zaitei_In_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function

Private Function Zaitei_Out_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�ݒ��i�{�j�w�莞�̃`�F�b�N���X�V�����x
'
'-------------------------------------------------------
Dim i               As Integer
Dim Hinban          As String * 13
Dim Tanaban         As String * 8
Dim sts             As Integer

Dim QTY             As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    Zaitei_Out_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Tanaban        '�I��
                Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                
                If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                Else
                '------------------ �q�Ƀ}�X�^�Ǎ���
                    Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                            Zaitei_Out_Proc = False
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
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Zaitei_Out_Proc = False
                            Exit Function
                        End If
                    End If
                    '------------------ �I�}�X�^�Ǎ���
                    Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                    Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                    Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                    Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Zaitei_Out_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                            Exit Function
                    End Select
            
                    '------------------ �֎~�I�̃`�F�b�N
                    If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                        Zaitei_Out_Proc = False
                        Exit Function
                    End If
            
                End If
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ �i�ڃ}�X�^�Ǎ���
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Zaitei_Out_Proc = False
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
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Zaitei_Out_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                End Select
            
            Case LCD_Suryo          '���ʁi�����͖����j
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_Out_Proc = False
                    Exit Function
                End If
                
                QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                If QTY = 0 Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_Out_Proc = False
                    Exit Function
                End If
            
            
            Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Zaitei_Out_Proc = False
                    Exit Function
                End If
                
                
                If Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD) = LCD_SUMI_Suryo Then
                    SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                Else
                    MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                End If
        
                If i = M_Gyo - 1 Then
                    If SUMI_QTY = 0 And MI_QTY = 0 Then
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "���i�^�����i���O", "���ʓ��̓~�X", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                        Zaitei_Out_Proc = False
                        Exit Function
                    End If
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
                                        
    If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
        MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                
    Else
                
        MENU_NO = ""
    End If
                                        
                                        
                                        '�o�ɍX�V
    sts = Syuko_Update_Proc(RET_JGYOBU, _
                            RET_NAIGAI, _
                            Hinban, _
                            "", _
                            Tanaban, _
                            (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                            SUMI_QTY, _
                            MI_QTY, _
                            0, _
                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                            FILE_RETRY, , , , , , , MENU_NO)
    Select Case sts
        Case False
        
        Case True       '�݌ɕs�����ɔ���
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "�݌ɐ��s��", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Zaitei_Out_Proc = False
            GoTo Abort_Tran
        Case SYS_CANCEL
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Zaitei_Out_Proc = False
            GoTo Abort_Tran
        Case SYS_ERR
            Sendbuf = Text_Create_Proc()
            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
            Zaitei_Out_Proc = SYS_ERR    '�V�X�e���ُ픭��
            
            GoTo Abort_Tran
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
    
    
    
    Zaitei_Out_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function

Private Function Ido_In_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�ړ����Ɏw�莞�̃`�F�b�N���X�V�����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim From_Tanaban    As String * 8
Dim To_Tanaban      As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    Ido_In_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        To_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(To_Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Ido_In_Proc = False
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
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Ido_In_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(To_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(To_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(To_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Ido_In_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Ido_In_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(To_Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            To_Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Ido_In_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_In_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
                                        'FROM ���z�I��
            From_Tanaban = ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM & "01" & "01" & "01"
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, From_Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           '�����ł͔������Ȃ�
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("�݌Ɏg�p��", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_In_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("�L���݌ɖ���", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                Ido_In_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = To_Tanaban       '�I�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i����
                                                        
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O2006.01.06
                                                        
                                                        
                                                        '���i���p�q�ɋ敪���Z�[�u
            ID_KANRI_TBL(ING_No).GOODS_ON_F = StrConv(SOKOREC.GOODS_ON_F, vbUnicode)
            '���ʕt���̑��M���b�Z�[�W���쐬����
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                            Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2))
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
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
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
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���i�^�����i���ʁj
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '���ʁi�����͖����j
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_In_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_In_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_In_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_In_Proc = False
                                Exit Function
                            
                            End If
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                            If MI_QTY > ID_KANRI_TBL(ING_No).Send_MI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_In_Proc = False
                                Exit Function
                            End If
                        End If
        
                        If i = M_Gyo - 1 Then       '�ŏI�s��������
                            If SUMI_QTY = 0 And MI_QTY = 0 Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "���i�^�����i���O", "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                Ido_In_Proc = False
                                Exit Function
                            End If
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
            '2006.01.30
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
        
        
            sts = IDO_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM & "01" & "01" & "01"), _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , MENU_NO)

    
    
            Select Case sts
                Case False
        
                Case True       '�݌ɕs�����ɔ���
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "�݌ɐ��s��", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                    Ido_In_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_In_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Ido_In_Proc = SYS_ERR    '�V�X�e���ُ픭��
                    GoTo Abort_Tran
            End Select
    
                                        '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '�o�ח\��^�݌ɂ̗\�����
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

    Ido_In_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Ido_Out_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�ړ��o�Ɏw�莞�̃`�F�b�N���X�V�����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim From_Tanaban    As String * 8
Dim To_Tanaban      As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    Ido_Out_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        From_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(From_Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Ido_Out_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(From_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(From_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(From_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Ido_Out_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Ido_Out_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(To_Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            From_Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Ido_Out_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_Out_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            
            '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
                                        'FROM ���I��
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, From_Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           '�����ł͔������Ȃ�
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("�݌Ɏg�p��", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_Out_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("�L���݌ɖ���", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                Ido_Out_Proc = False
                Exit Function
            End If
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = From_Tanaban     '�I�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i�̐���
                                                            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O2006.01.06
                                                            
                                                            
                                                            
                                                            '���i���p�q�ɋ敪���Z�[�u
            ID_KANRI_TBL(ING_No).GOODS_ON_F = StrConv(SOKOREC.GOODS_ON_F, vbUnicode)
            '���ʕt���̑��M���b�Z�[�W���쐬����
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                    Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                    Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))
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
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
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
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���i�^�����i���ʁj
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '���ʁi�����͖����j
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_Out_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_Out_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Ido_Out_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_Out_Proc = False
                                Exit Function
                            End If
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Ido_Out_Proc = False
                                Exit Function
                            End If
                        End If
        
                        If i = M_Gyo - 1 Then       '�ŏI�s��������
                            If SUMI_QTY = 0 And MI_QTY = 0 Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "���i�^�����i���O", "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                Ido_Out_Proc = False
                                Exit Function
                            End If
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
            '2006.01.30
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
            Else
                MENU_NO = ""
            End If
        
            sts = IDO_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.PARAM & "01" & "01" & "01"), _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , MENU_NO)

    
    
            Select Case sts
                Case False
        
                Case True       '�݌ɕs�����ɔ���
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "�݌ɐ��s��", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                    Ido_Out_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Ido_Out_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Ido_Out_Proc = SYS_ERR    '�V�X�e���ُ픭��
                    GoTo Abort_Tran
            End Select
    
                                        '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
            '�o�ח\��^�݌ɂ̗\�����
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

    Ido_Out_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Tanahyoji_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�I�ԕ\�������x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13
Dim Tanaban     As String * 8
Dim ST_Tanaban  As String * 8


Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim Tanahyoji() As Tanahyoji_tag
Dim Tana_Cnt    As Integer


Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1


    Tanahyoji_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    '�i�ڃ}�X�^�ǂݍ��݁i�W���I�Ԃf�d�s�j
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        ST_Tanaban = (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Case BtErrKeyNotFound
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Tanahyoji_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                End Select
                
                
                On Error Resume Next
                Kill (FullPath)             '���M�p�t�@�C���폜
                On Error GoTo 0
        
                Erase Tanahyoji
                Tana_Cnt = -1
        
                If Len(Trim(ST_Tanaban)) = 0 Then
                                            '�W���I�Ԑݒ�Ȃ�
                
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                
                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                Else
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            RET_JGYOBU, _
                                            RET_NAIGAI, _
                                            Hinban, _
                                            ST_Tanaban) Then

            
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Exit Function
                    End If
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                    Tanahyoji(Tana_Cnt).SUMI_QTY = SUMI_QTY
                    Tanahyoji(Tana_Cnt).MI_QTY = MI_QTY
                    
                    
                    
                End If
                
                Call UniCode_Conv(K6_ZAIKO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K6_ZAIKO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K6_ZAIKO.HIN_GAI, Hinban)
                Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                Call UniCode_Conv(K6_ZAIKO.Retu, "")
                Call UniCode_Conv(K6_ZAIKO.Ren, "")
                Call UniCode_Conv(K6_ZAIKO.Dan, "")
                    
                com = BtOpGetGreater
        
                SUMI_QTY = 0
                MI_QTY = 0
    
    
                Do
                    sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "�݌Ƀf�[�^", 0)
                            Exit Function
                    End Select
                
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                        '���ƕ��^�����O�^�i�ԃu���[�N
                        Exit Do
                    End If
                    
                    If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) _
                        = ST_Tanaban Then
                        '�W���I�Ԃ͑ΏۊO
                    Else
                        If Tana_Cnt = (-1) Then
                            Tana_Cnt = Tana_Cnt + 1
                            ReDim Tanahyoji(Tana_Cnt)
                            Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                Tanahyoji(Tana_Cnt).MI_QTY = 0
                            Else
                                Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                            End If
                        
                        Else
                            For j = 0 To UBound(Tanahyoji)
                                If Tanahyoji(j).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                                    Exit For
                                End If
                            Next j
                        
                            If j <= UBound(Tanahyoji) Then
                            
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(j).SUMI_QTY = Tanahyoji(j).SUMI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                Else
                                    Tanahyoji(j).MI_QTY = Tanahyoji(j).MI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            
                            Else
                            
                                Tana_Cnt = Tana_Cnt + 1
                            
                                ReDim Preserve Tanahyoji(Tana_Cnt)
                                Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                                Else
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                    Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            End If
                        
                        End If
                    
                    End If
                    
                    com = BtOpGetNext
                    
                Loop
        
        
        
                FileNo = FreeFile           '���M�p�t�@�C���n�o�d�m
                Open FullPath For Binary As #FileNo
        
        
                SendFileRec.Title = "0"     '�^�C�g���s
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Tana_Cnt + 1, "#0") & "��")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
        
        
                If Tana_Cnt > -1 Then
                                '�W�v�e�[�u�����f�[�^�o��
                    For j = 0 To UBound(Tanahyoji)
                                
                        SendFileRec.Title = "1"
'                        Call UniCode_Conv(SendFileRec.LCD, Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
'                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
'                        SendFileRec.CRLF = vbCrLf
                        
                        If j = 0 Then
                            Call UniCode_Conv(SendFileRec.LCD, "*" & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        Else
                            Call UniCode_Conv(SendFileRec.LCD, " " & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        End If
                        SendFileRec.CRLF = vbCrLf
                                            
                        Put #FileNo, , SendFileRec
                                            
                                            
                        Call UniCode_Conv(SendFileRec.LCD, "  ���F" & Format(Tanahyoji(j).SUMI_QTY, "#0") & "  ���F" & Format(Tanahyoji(j).MI_QTY, "#0"))
                        SendFileRec.CRLF = vbCrLf
                                            
                                            
                        
                        Put #FileNo, , SendFileRec
                            
                    Next j
                End If
        
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '���M���b�Z�[�W���쐬����
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '�\����ʃt���O �Q�Ɖ��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '���M�f�[�^�t�@�C����
    Send_Text.fileName = B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B1_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------�P�`�T�s��
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX����
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            '���l�����\��
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            
                                                            '�����J�[�\���ʒu
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '���͌���
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            '���j���\�ԍ�
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Tanahyoji_Proc = False
    

End Function
Private Function Tanahyoji_Kaso_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�I�ԕ\��(���z�D��)�����x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13
Dim Tanaban     As String * 8
Dim ST_Tanaban  As String * 8


Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim Tanahyoji() As Tanahyoji_tag
Dim Tana_Cnt    As Integer


Dim RET_JGYOBU  As String * 1
Dim RET_NAIGAI  As String * 1


    Tanahyoji_Kaso_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    '�i�ڃ}�X�^�ǂݍ��݁i�W���I�Ԃf�d�s�j
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        ST_Tanaban = (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Case BtErrKeyNotFound
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Tanahyoji_Kaso_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                End Select
                
                
                On Error Resume Next
                Kill (FullPath)             '���M�p�t�@�C���폜
                On Error GoTo 0
        
                Erase Tanahyoji
                Tana_Cnt = -1
        
                If Len(Trim(ST_Tanaban)) = 0 Then
                                            '�W���I�Ԑݒ�Ȃ�
                
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                
                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                Else
                    If Zaiko_Syukei_Proc(SUMI_QTY, _
                                            MI_QTY, _
                                            RET_JGYOBU, _
                                            RET_NAIGAI, _
                                            Hinban, _
                                            ST_Tanaban) Then

            
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Exit Function
                    End If
                    Tana_Cnt = Tana_Cnt + 1
                    ReDim Tanahyoji(Tana_Cnt)
                    Tanahyoji(Tana_Cnt).Tanaban = ST_Tanaban
                    Tanahyoji(Tana_Cnt).SUMI_QTY = SUMI_QTY
                    Tanahyoji(Tana_Cnt).MI_QTY = MI_QTY
                    
                    
                    
                End If
                
                Call UniCode_Conv(K6_ZAIKO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K6_ZAIKO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K6_ZAIKO.HIN_GAI, Hinban)
                Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                Call UniCode_Conv(K6_ZAIKO.Retu, "")
                Call UniCode_Conv(K6_ZAIKO.Ren, "")
                Call UniCode_Conv(K6_ZAIKO.Dan, "")
                    
                com = BtOpGetGreater
        
                SUMI_QTY = 0
                MI_QTY = 0
    
    
                Do
                    sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "�݌Ƀf�[�^", 0)
                            Exit Function
                    End Select
                
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                        '���ƕ��^�����O�^�i�ԃu���[�N
                        Exit Do
                    End If
                    
                    If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) _
                        = ST_Tanaban Then
                        '�W���I�Ԃ͑ΏۊO
                    Else
                        
                        '���z�q�ɂ�D�悷��
                        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_JITU)
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�q��Ͻ�", 0)
                                Exit Function
                        End Select
                        
                        
                        
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                        
                            If Tana_Cnt = (-1) Then
                                Tana_Cnt = Tana_Cnt + 1
                                ReDim Tanahyoji(Tana_Cnt)
                                Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                                Else
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                    Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            Else
                                For j = 0 To UBound(Tanahyoji)
                                    If Tanahyoji(j).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                                        Exit For
                                    End If
                                Next j
                            
                                If j <= UBound(Tanahyoji) Then
                                
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(j).SUMI_QTY = Tanahyoji(j).SUMI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Else
                                        Tanahyoji(j).MI_QTY = Tanahyoji(j).MI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                
                                Else
                                
                                    Tana_Cnt = Tana_Cnt + 1
                                
                                    ReDim Preserve Tanahyoji(Tana_Cnt)
                                    Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                        Tanahyoji(Tana_Cnt).MI_QTY = 0
                                    Else
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                        Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                End If
                            
                            End If
                        
                        End If
                    End If
                    
                    com = BtOpGetNext
                    
                Loop
        
        
                Call UniCode_Conv(K6_ZAIKO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K6_ZAIKO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K6_ZAIKO.HIN_GAI, Hinban)
                Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
                Call UniCode_Conv(K6_ZAIKO.Soko_No, "")
                Call UniCode_Conv(K6_ZAIKO.Retu, "")
                Call UniCode_Conv(K6_ZAIKO.Ren, "")
                Call UniCode_Conv(K6_ZAIKO.Dan, "")
                    
                com = BtOpGetGreater
        
                SUMI_QTY = 0
                MI_QTY = 0
    
    
                Do
                    sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "�݌Ƀf�[�^", 0)
                            Exit Function
                    End Select
                
                    If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                        '���ƕ��^�����O�^�i�ԃu���[�N
                        Exit Do
                    End If
                    
                    If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) _
                        = ST_Tanaban Then
                        '�W���I�Ԃ͑ΏۊO
                    Else
                        
                        '���z�q�ɂ�D�悷��
                        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO)
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�q��Ͻ�", 0)
                                Exit Function
                        End Select
                        
                        
                        
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_JITU Then
                        
                            If Tana_Cnt = (-1) Then
                                Tana_Cnt = Tana_Cnt + 1
                                ReDim Tanahyoji(Tana_Cnt)
                                Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Tanahyoji(Tana_Cnt).MI_QTY = 0
                                Else
                                    Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                    Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                End If
                            
                            Else
                                For j = 0 To UBound(Tanahyoji)
                                    If Tanahyoji(j).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode) Then
                                        Exit For
                                    End If
                                Next j
                            
                                If j <= UBound(Tanahyoji) Then
                                
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(j).SUMI_QTY = Tanahyoji(j).SUMI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    Else
                                        Tanahyoji(j).MI_QTY = Tanahyoji(j).MI_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                
                                Else
                                
                                    Tana_Cnt = Tana_Cnt + 1
                                
                                    ReDim Preserve Tanahyoji(Tana_Cnt)
                                    Tanahyoji(Tana_Cnt).Tanaban = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                                    If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                        Tanahyoji(Tana_Cnt).MI_QTY = 0
                                    Else
                                        Tanahyoji(Tana_Cnt).SUMI_QTY = 0
                                        Tanahyoji(Tana_Cnt).MI_QTY = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                    End If
                                
                                End If
                            
                            End If
                        
                        End If
                    End If
                    
                    com = BtOpGetNext
                    
                Loop
        
        
        
        
                FileNo = FreeFile           '���M�p�t�@�C���n�o�d�m
                Open FullPath For Binary As #FileNo
        
        
                SendFileRec.Title = "0"     '�^�C�g���s
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Tana_Cnt + 1, "#0") & "��")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
        
        
                If Tana_Cnt > -1 Then
                                '�W�v�e�[�u�����f�[�^�o��
                    For j = 0 To UBound(Tanahyoji)
                                
                        SendFileRec.Title = "1"
'                        Call UniCode_Conv(SendFileRec.LCD, Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
'                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
'                        SendFileRec.CRLF = vbCrLf
                        
                        If j = 0 Then
                            Call UniCode_Conv(SendFileRec.LCD, "*" & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        Else
                            Call UniCode_Conv(SendFileRec.LCD, " " & Left(Tanahyoji(j).Tanaban, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 3, 2) & "-" & Mid(Tanahyoji(j).Tanaban, 5, 2) & "-" & Right(Tanahyoji(j).Tanaban, 2) & ":" _
                                            & Format(Tanahyoji(j).SUMI_QTY + Tanahyoji(j).MI_QTY, "#0"))
                        End If
                        SendFileRec.CRLF = vbCrLf
                                            
                        Put #FileNo, , SendFileRec
                                            
                                            
                        Call UniCode_Conv(SendFileRec.LCD, "  ���F" & Format(Tanahyoji(j).SUMI_QTY, "#0") & "  ���F" & Format(Tanahyoji(j).MI_QTY, "#0"))
                        SendFileRec.CRLF = vbCrLf
                                            
                                            
                        
                        Put #FileNo, , SendFileRec
                            
                    Next j
                End If
        
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '���M���b�Z�[�W���쐬����
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '�\����ʃt���O �Q�Ɖ��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '���M�f�[�^�t�@�C����
    Send_Text.fileName = B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B9_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------�P�`�T�s��
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX����
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            '���l�����\��
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            
                                                            '�����J�[�\���ʒu
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '���͌���
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            '���j���\�ԍ�
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Tanahyoji_Kaso_Proc = False
    

End Function
Private Function Ave_Syuka_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�����Ϗo�א��\�������x
'
'-------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim i               As Integer

Dim Hinban          As String * 13
Dim Tanaban         As String * 8

Dim AVE_SYUKA_ED    As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


    Ave_Syuka_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    '�i�ڃ}�X�^�ǂݍ���
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Ave_Syuka_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                End Select
                                '�����Ϗo�א��ǂݍ���
''''''''''''''''2006.01.06 ���ނɑΉ�
'                Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, RET_NAIGAI)
''''''''''''''''2006.01.06 ���ނɑΉ�
                Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, Hinban)
                sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(AVE_SYUKAREC.AVE_SYUKA, "00000000")
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�", 0)
                        Exit Function
                End Select
        
        End Select
    Next i
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '���M���b�Z�[�W���쐬����
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �Q�Ɖ��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '���M�f�[�^�t�@�C����
    Send_Text.fileName = ""
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------�P�s��
                                                            'BOX����
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                            '���l�����\��
    Send_Text.Box_Type(0).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(0).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(0).Max_Size = "00"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                            '���j���\�ԍ�
    Send_Text.Box_Type(0).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).MENU = ""
    
    '-----------------------------------------------�Q�s��
                                                            'BOX����
    Send_Text.Box_Type(1).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Hinban)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Hinban)
                                                            '���l�����\��
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(1).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(1).Max_Size = "13"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "13"
                                                            '���j���\�ԍ�
    Send_Text.Box_Type(1).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '-----------------------------------------------�R�s��
                                                            'BOX����
    Send_Text.Box_Type(2).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            
                                                            
    AVE_SYUKA_ED = Format(CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode)), "#0")
    AVE_SYUKA_ED = "[" & Space(8 - Len(AVE_SYUKA_ED)) & AVE_SYUKA_ED & "]"
                                                             
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, AVE_SYUKA_ED)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, AVE_SYUKA_ED)
                                                            '���l�����\��
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(2).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(2).Max_Size = "08"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "08"
                                                            '���j���\�ԍ�
    Send_Text.Box_Type(2).MENU = ""
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
                                                            '���j���\�ԍ�
    Send_Text.Box_Type(3).MENU = ""
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
                                                            '���j���\�ԍ�
    Send_Text.Box_Type(4).MENU = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        

    Sendbuf = Text_Create_Proc()
    
    
    
    Ave_Syuka_Proc = False
    

End Function


Private Function Host_Zaiko_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�z�X�g�݌Ɂi���_�݌Ɂj�Ɖ���x
'
'-------------------------------------------------------
Dim sts             As Integer
Dim i               As Integer
Dim Hinban          As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1
    
    
    Host_Zaiko_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Host_Zaiko_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                
                End Select
                '------------------ �݌ɏW�v�f�[�^�Ǎ���
                Call UniCode_Conv(K0_SUMZ.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
                Call UniCode_Conv(K0_SUMZ.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
                Call UniCode_Conv(K0_SUMZ.HIN_GAI, Hinban)
                sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�݌ɏW�v�f�[�^", 0)
                        Exit Function
                End Select
            
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '���M���b�Z�[�W���쐬����
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    
    '-----------------------------------------------�P�s��
                                                    'BOX����
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                    '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Hinban)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Hinban)
                                                            '���l�����\��
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(1).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(1).Max_Size = "13"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "13"
                                                                                
    Send_Text.Box_Type(1).MENU = ""                         '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).MENU = ""
    '-----------------------------------------------�R�s��
                                                            'BOX����
    Send_Text.Box_Type(2).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�z�X�g�݌�:" & Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#0"))
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�z�X�g�݌�:" & Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#0"))
                                                            '���l�����\��
    Send_Text.Box_Type(2).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(2).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(2).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
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
    Send_Text.Box_Type(3).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
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
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "")
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "")
                                                            '���l�����\��
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(4).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(4).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(4).MENU = ""                     '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
    
    Sendbuf = Text_Create_Proc()
    
    
    
    Host_Zaiko_Proc = False
    

End Function
Private Function Tanaorosi_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�I���Ɖ���x
'
'-------------------------------------------------------
Dim sts             As Integer
Dim i               As Integer
Dim Hinban          As String
Dim Tanaban         As String
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim Sumi_ED         As String
Dim Mi_ED           As String
Dim Total_ED        As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

    Tanaorosi_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            
            Case LCD_Tanaban        '�I��
                
                Tanaban = Trim(Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1))
                
                
            Case LCD_Hinban         '�i��
                
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
        
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Tanaorosi_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                
                End Select
        
        End Select
    Next i
    
    sts = Zaiko_Syukei_Proc(SUMI_QTY, _
                            MI_QTY, _
                            ID_KANRI_TBL(ING_No).JGYOBU, _
                            ID_KANRI_TBL(ING_No).NAIGAI, _
                            Hinban, _
                            Tanaban)
    If sts Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Exit Function
    End If
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '���M���b�Z�[�W���쐬����
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
    Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
    Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------�P�s��
                                                        'BOX����
    Send_Text.Box_Type(0).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                        '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
    Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2))
                                                            '���l�����\��
    Send_Text.Box_Type(1).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(1).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(1).Max_Size = "13"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "13"
                                                                                
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
    Send_Text.Box_Type(2).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(2).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
    '-----------------------------------------------�S�s��
                                                             'BOX����
    Send_Text.Box_Type(3).Box_Type = TYPE_REF
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '�\�����e
    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "�� �i/�� �i/�� �v")
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "�� �i/�� �i/�� �v")
                                                            '���l�����\��
    Send_Text.Box_Type(3).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(3).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
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
    Sumi_ED = Space(5 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0") & "/"
    Mi_ED = Space(5 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0") & "/"
    Total_ED = Space(5 - Len(Format(SUMI_QTY + MI_QTY, "#0"))) & Format(SUMI_QTY + MI_QTY, "#0")
    
    Call UniCode_Conv(Send_Text.Box_Type(4).LCD, Sumi_ED & Mi_ED & Total_ED)
    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, Sumi_ED & Mi_ED & Total_ED)
                                                            '���l�����\��
    Send_Text.Box_Type(4).INIT = ""
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '�����J�[�\���ʒu
    Send_Text.Box_Type(4).Start_Pos = "01"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                            '���͌���
    Send_Text.Box_Type(4).Max_Size = "05"
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
    Send_Text.Box_Type(4).MENU = ""                     '���j���\�ԍ�
    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
    
    Sendbuf = Text_Create_Proc()
    
    
    
    Tanaorosi_Proc = False
    

End Function
Private Function St_Tanaban_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�W���I�Ԑݒ菈���x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim i           As Integer
Dim Hinban      As String
Dim Tanaban     As String

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


    St_Tanaban_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            
            Case LCD_Tanaban        '�I��
                Tanaban = Trim(Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1))
                                '------------------ �q�Ƀ}�X�^�Ǎ���
                Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                        St_Tanaban_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                        Exit Function
                End Select
                    
                '------------------ ���ڃ`�F�b�N    2006.01.06 �i��������Ɉړ���
'                If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
'                    If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
'                        StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
'                        Sendbuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                        St_Tanaban_Proc = False
'                        Exit Function
'                    End If
'                End If
                '------------------ �I�}�X�^�Ǎ���
                Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        St_Tanaban_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                        Exit Function
                    End Select
   
            Case LCD_Hinban         '�i��
                
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ �i�ڃ}�X�^�Ǎ���
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        St_Tanaban_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                End Select
        
        
                '------------------ ���ڃ`�F�b�N
                If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
                    If StrConv(SOKOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                        StrConv(SOKOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Then
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        St_Tanaban_Proc = False
                        Exit Function
                    End If
                End If
        
        
        
        End Select
    Next i
    '----------------------------------- �f�[�^�X�V�����J�n -----------
    Call UniCode_Conv(K0_ITEM.JGYOBU, RET_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, RET_NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
    
    Do
        '------------------ �i�ڃ}�X�^�Ǎ���
        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
            '   -------------------------------- �G���[���b�Z�[�W�쐬
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
        
                St_Tanaban_Proc = False
                Exit Function
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            '   -------------------------------- �G���[���b�Z�[�W�쐬
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "���[���Ŏg�p��", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                St_Tanaban_Proc = False
                Exit Function
            Case Else
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                Exit Function
        End Select
    
    Loop
                                        '�W���I�Ԑݒ�
    Call UniCode_Conv(ITEMREC.ST_SOKO, Left(Tanaban, 2))
    Call UniCode_Conv(ITEMREC.ST_RETU, Mid(Tanaban, 3, 2))
    Call UniCode_Conv(ITEMREC.ST_REN, Mid(Tanaban, 5, 2))
    Call UniCode_Conv(ITEMREC.ST_DAN, Right(Tanaban, 2))
    Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
                                        
    Do
        '------------------ �i�ڃ}�X�^�X�V
        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            '   -------------------------------- �G���[���b�Z�[�W�쐬
                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "���[���Ŏg�p��", "", "")
                    
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                St_Tanaban_Proc = False
                Exit Function
            Case Else
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                Exit Function
        End Select
    Loop
                                        
                                        
                                        
                                        
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
    
    
    
    St_Tanaban_Proc = False
    

End Function
Private Function Rireki_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�����o�ɐ��ڕ\�������x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1



Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag

Dim Data_Cnt    As Integer


    Rireki_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Rireki_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                
                End Select
                                    
                                    
                                    
                                    
                                    '��ǂ݂��Č����J�E���g
                
                
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Format(Now, "YYYYMMDD"))
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                Data_Cnt = 0
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '���ƕ��^���O�^�i�ԃu���[�N�H
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                Exit Do
                            End If
                        
                            '���t�H
                            If StrConv(IDOREC.JITU_DT, vbUnicode) <> Format(Now, "YYYYMMDD") Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "�݌Ɉړ���", 0)
                            Exit Function
                    End Select
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_ZAITEI_OUT Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                    
                    
                        Data_Cnt = Data_Cnt + 1
                    
                    End If
                    
                    com = BtOpGetPrev
                
                Loop
                
                
                On Error Resume Next
                Kill (FullPath)             '���M�p�t�@�C���폜
                On Error GoTo 0
        
                FileNo = FreeFile           '���M�p�t�@�C���n�o�d�m
                Open FullPath For Binary As #FileNo
        
                SendFileRec.Title = "0"     '�^�C�g���s
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Data_Cnt, "#0") & "��")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                SendFileRec.Title = "0"     '�i��
                Call UniCode_Conv(SendFileRec.LCD, Hinban)
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                    
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Format(Now, "YYYYMMDD"))
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '���ƕ��^���O�^�i�ԃu���[�N�H
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                Exit Do
                            End If
                        
                            '���t�H
                            If StrConv(IDOREC.JITU_DT, vbUnicode) <> Format(Now, "YYYYMMDD") Then
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "�݌Ɉړ���", 0)
                            Exit Function
                    End Select
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_ZAITEI_OUT Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                    
                        '���������o��
                        SendFileRec.Title = "1"
                        SUMI_QTY = CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode))
                        MI_QTY = CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
                        Call UniCode_Conv(SendFileRec.LCD, StrConv(IDOREC.RIRK_NAME, vbUnicode) & _
                                            Space(10 - Len(Format(SUMI_QTY + MI_QTY, "#0"))) & _
                                            Format(SUMI_QTY + MI_QTY, "#0"))
                        SendFileRec.CRLF = vbCrLf
                        Put #FileNo, , SendFileRec
                    End If
                    
                    com = BtOpGetPrev
                
                Loop
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '���M���b�Z�[�W���쐬����
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '�\����ʃt���O �Q�Ɖ��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '���M�f�[�^�t�@�C����
    Send_Text.fileName = B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B6_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------�P�`�T�s��
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX����
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            
                                                            '���l�����\��
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            
                                                            '�����J�[�\���ʒu
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '���͌���
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            '���j���\�ԍ�
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Rireki_Proc = False
    

End Function
Private Function Suii_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�o�א��ڕ\�������x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim j           As Integer

Dim Hinban      As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


Dim SUMI_QTY    As Long
Dim MI_QTY      As Long




Dim Start_YMD   As String * 8
Dim End_YMD     As String * 8
Dim Save_YMD    As String * 6
Dim SYUKA_QTY   As Long


Dim FileNo      As Integer

Dim FullPath    As String

Dim SendFileRec As SendFileRec_Tag


Dim Data_Cnt    As Integer

    Suii_Proc = True

    If Right(CtrsWsk1.SendFolder, 1) <> "\" Then
        FullPath = CtrsWsk1.SendFolder & "\" & B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    Else
        FullPath = CtrsWsk1.SendFolder & B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    End If

    For i = 31 To 28 Step -1
        Start_YMD = Left(Format(DateAdd("m", -1, Now), "YYYYMMDD"), 6) & Format(i, "00")
        If IsDate(Left(Start_YMD, 4) & "/" & Mid(Start_YMD, 5, 2) & "/" & Right(Start_YMD, 2)) Then
            Exit For
        End If
    Next i

'    Start_YMD = Left(Format(DateAdd("m", -1, Now), "YYYYMMDD"), 6) & "31"
    
    End_YMD = Left(Format(DateAdd("m", -11, (Left(Start_YMD, 4) & "/" & Mid(Start_YMD, 5, 2) & "/" & Right(Start_YMD, 2))), "YYYYMMDD"), 6) & "01"

    Save_YMD = ""

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                                    
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Suii_Proc = False
                        Exit Function
                    
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                
                End Select
                                    
                                    '��ǂ݂��ăf�[�^�����l��
                
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Start_YMD)
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                Data_Cnt = 0
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '���ƕ��^���O�^�i�ԃu���[�N�H
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    Data_Cnt = Data_Cnt + 1
                                End If
                                
                                Exit Do
                            
                            End If
                            '���t�H
                            If StrConv(IDOREC.JITU_DT, vbUnicode) < End_YMD Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    Data_Cnt = Data_Cnt + 1
                                End If
                                
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            If Len(Trim(Save_YMD)) <> 0 Then
                                Data_Cnt = Data_Cnt + 1
                            End If
                            
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "�݌Ɉړ���", 0)
                            Exit Function
                    End Select
                    
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                        
                        
                        If Len(Trim(Save_YMD)) = 0 Then
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                        End If
                        If Save_YMD <> Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6) Then
                            Data_Cnt = Data_Cnt + 1
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                        End If
                    
                    
                    
                    End If
                    
                    com = BtOpGetPrev
                Loop
                
                
                On Error Resume Next
                Kill (FullPath)             '���M�p�t�@�C���폜
                On Error GoTo 0
        
                FileNo = FreeFile           '���M�p�t�@�C���n�o�d�m
                Open FullPath For Binary As #FileNo
        
                SendFileRec.Title = "0"     '�^�C�g���s
                Call UniCode_Conv(SendFileRec.LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME & Format(Data_Cnt, "#0") & "��")
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                SendFileRec.Title = "0"     '�i��
                Call UniCode_Conv(SendFileRec.LCD, Hinban)
                SendFileRec.CRLF = vbCrLf
                Put #FileNo, , SendFileRec
                    
                    
                    
                Save_YMD = ""
                    
                Call UniCode_Conv(K1_IDO.JGYOBU, RET_JGYOBU)
                Call UniCode_Conv(K1_IDO.NAIGAI, RET_NAIGAI)
                Call UniCode_Conv(K1_IDO.HIN_GAI, Hinban)
                Call UniCode_Conv(K1_IDO.JITU_DT, Start_YMD)
                Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzz")
                
                com = BtOpGetLess
                
                
                Do
                    
                    sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
                    Select Case sts
                        Case BtNoErr
                            '���ƕ��^���O�^�i�ԃu���[�N�H
                            If StrConv(IDOREC.JGYOBU, vbUnicode) <> RET_JGYOBU Or _
                                StrConv(IDOREC.NAIGAI, vbUnicode) <> RET_NAIGAI Or _
                                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Hinban) Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    SendFileRec.Title = "1"
                                    Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                        Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                        Format(SYUKA_QTY, "#0"))
                                    SendFileRec.CRLF = vbCrLf
                                    Put #FileNo, , SendFileRec
                                End If
                                
                                Exit Do
                            
                            End If
                            '���t�H
                            If StrConv(IDOREC.JITU_DT, vbUnicode) < End_YMD Then
                                
                                If Len(Trim(Save_YMD)) <> 0 Then
                                    SendFileRec.Title = "1"
                                    Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                        Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                        Format(SYUKA_QTY, "#0"))
                                    SendFileRec.CRLF = vbCrLf
                                    Put #FileNo, , SendFileRec
                                End If
                                
                                Exit Do
                            End If
                        
                        Case BtErrEOF
                            If Len(Trim(Save_YMD)) <> 0 Then
                                SendFileRec.Title = "1"
                                Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                    Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                    Format(SYUKA_QTY, "#0"))
                                SendFileRec.CRLF = vbCrLf
                                Put #FileNo, , SendFileRec
                            End If
                            
                            Exit Do
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, com, "�݌Ɉړ���", 0)
                            Exit Function
                    End Select
                    
                    
                    If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_KEI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_HYO Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_SYUKA_GAI Or _
                        Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = ACT_DENPYO_ID Then
                        
                        
                        If Len(Trim(Save_YMD)) = 0 Then
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                            SYUKA_QTY = 0
                        End If
                        If Save_YMD <> Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6) Then
                                
                            SendFileRec.Title = "1"
                            Call UniCode_Conv(SendFileRec.LCD, Left(Save_YMD, 4) & "/" & Right(Save_YMD, 2) & _
                                                Space(13 - Len(Format(SYUKA_QTY, "#0"))) & _
                                                Format(SYUKA_QTY, "#0"))
                            SendFileRec.CRLF = vbCrLf
                            Put #FileNo, , SendFileRec
                            
                            Save_YMD = Left(StrConv(IDOREC.JITU_DT, vbUnicode), 6)
                            SYUKA_QTY = 0
                                                
                        End If
                    
                        SYUKA_QTY = SYUKA_QTY + (CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                    
                    
                    End If
                    
                    com = BtOpGetPrev
                Loop
        
                Close #FileNo
        
        
        End Select
    Next i
    
    
    ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            
    '���M���b�Z�[�W���쐬����
    Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
    Send_Text.Display_Flg = Display_REF                     '�\����ʃt���O �Q�Ɖ��
    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_REF
    
    Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
    Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                                                            '���M�f�[�^�t�@�C����
    Send_Text.fileName = B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    ID_KANRI_TBL(ING_No).Send_Text.fileName = B7_SendFile & "." & Format(ID_KANRI_TBL(ING_No).ID, "000")
    
    Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
    '-----------------------------------------------�P�`�T�s��
                                                            
    For i = 0 To M_Gyo - 1
                                                            'BOX����
        Send_Text.Box_Type(i).Box_Type = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Box_Type = ""
                                                            '�\�����e
        Call UniCode_Conv(Send_Text.Box_Type(i).LCD, "")
        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, "")
                                                            '���l�����\��
        Send_Text.Box_Type(i).INIT = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).INIT = ""
                                                            '�����J�[�\���ʒu
        Send_Text.Box_Type(i).Start_Pos = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Start_Pos = ""
                                                            '���͌���
        Send_Text.Box_Type(i).Max_Size = "00"
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size = "00"
                                                            '���j���\�ԍ�
        Send_Text.Box_Type(i).MENU = ""
        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).MENU = ""
        
    Next i

    Sendbuf = Text_Create_Proc()
    
    
    
    Suii_Proc = False
    

End Function
Private Function Hin_Shogo_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�i�ԕʍ݌ɏƍ������x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MEMO            As String

Dim MENU_NO         As String

    Hin_Shogo_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then     '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Hin_Shogo_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ ���ڃ`�F�b�N
'                            If StrConv(SOKOREC.KONS_KBN, vbUnicode) = KONS_KBN_NG Then
'                                If StrConv(SOKOREC.JGYOBU, vbUnicode) <> ID_KANRI_TBL(ING_No).JGYOBU Or _
'                                    StrConv(SOKOREC.NAIGAI, vbUnicode) <> ID_KANRI_TBL(ING_No).NAIGAI Then
'                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
'                                    SendBuf = Text_Create_Proc()
'                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'                                    Hin_Shogo_Proc = False
'                                    Exit Function
'                                End If
'                            End If
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Hin_Shogo_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ �֎~�I�̃`�F�b�N
'                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
'
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�g�p�s��", "", "")
'
'                                SendBuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                Ido_In_Proc = False
'                                Exit Function
'                            End If
            
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Hin_Shogo_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Hin_Shogo_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            '   -------------------------------- �݌ɐ��W�v
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, Hinban) Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            '   -------------------------------- ���M�e�L�X�g�쐬
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '�I�Ԃ��Z�[�u
            
            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O2006.01.06
            
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i����
                                                        
            '���ʕt���̑��M���b�Z�[�W���쐬����
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
            Send_Text.Box_Type(1).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
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
            Send_Text.Box_Type(2).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                            '���͌���
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
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
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���i�^�����i���ʁj
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '���ʁi�����͖����j
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
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
                                        '�i�ڃ}�X�^�Ǎ���
            Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).S_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).S_NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).Hinban)
            Do
                sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "���Ŏg�p��", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Hin_Shogo_Proc = False
                        GoTo Abort_Tran
                    Case BtErrKeyNotFound
                    '   -------------------------------- �G���[���b�Z�[�W�쐬
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "�i�Ԗ��o�^", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Hin_Shogo_Proc = False
                        GoTo Abort_Tran
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)
                        Hin_Shogo_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
        
            Loop
                                        '�ŏI�ƍ����t
            Call UniCode_Conv(ITEMREC.LAST_CHK_DT, Format(Date, "yyyymmdd"))
                                        '�ŏI�ƍ��݌ɐ�
            Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, Format(SUMI_QTY + MI_QTY, "00000000"))
                                        '�i�ڃ}�X�^��������
            Do
                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "���Ŏg�p��", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Hin_Shogo_Proc = False
                        GoTo Abort_Tran
            
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^", 0)
                        Hin_Shogo_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
        
            If (SUMI_QTY + MI_QTY) = (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                MEMO = B2_MEMO & StrConv(Format((SUMI_QTY + MI_QTY), "#0"), vbWide) & "[OK]"
            Else
                MEMO = B2_MEMO & StrConv(Format((ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_SUMI_QTY), "#0"), vbWide) & "[" & StrConv(Format(SUMI_QTY + MI_QTY, "#0"), vbWide) & "]"
            End If
            '2006.01.30
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
            Else
                MENU_NO = ""
            End If
                                
                                
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
                                        MEMO, , , , , MENU_NO)
            Select Case sts
                Case False      '����I��
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Hin_Shogo_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
    
    
                
    
    
    
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

    Hin_Shogo_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Tana_Hin_Shogo_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�I�ԕʕi�ԕʍ݌ɏƍ������x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MEMO            As String

Dim MENU_NO         As String * 2
    
    Tana_Hin_Shogo_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then     '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Tana_Hin_Shogo_Proc = False
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
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Tana_Hin_Shogo_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Tana_Hin_Shogo_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ �֎~�I�̃`�F�b�N
'                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
'
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�g�p�s��", "", "")
'
'                                SendBuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                Ido_In_Proc = False
'                                Exit Function
'                            End If
            
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Tana_Hin_Shogo_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Tana_Hin_Shogo_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            '   -------------------------------- �݌ɐ��W�v
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    RET_JGYOBU, _
                                    RET_NAIGAI, _
                                    Hinban, _
                                    Tanaban) Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            '   -------------------------------- ���M�e�L�X�g�쐬
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '�I�Ԃ��Z�[�u
            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O2006.01.06
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i����
                                                        
            '���ʕt���̑��M���b�Z�[�W���쐬����
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
            Send_Text.Box_Type(1).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
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
            Send_Text.Box_Type(2).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                            '���͌���
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
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
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���i�^�����i���ʁj
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '���ʁi�����͖����j
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Tana_Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Tana_Hin_Shogo_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
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
                                        '�i�ڃ}�X�^�Ǎ���
'            Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'            Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'            Call UniCode_Conv(K0_ITEM.HIN_GAI, ID_KANRI_TBL(ING_No).Hinban)
'            Do
'                sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                        Exit Do
'                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'                    '   -------------------------------- �G���[���b�Z�[�W�쐬
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "���Ŏg�p��", "", "")
'
'                        SendBuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                        Hin_Shogo_Proc = False
'                        GoTo Abort_Tran
'                    Case BtErrKeyNotFound
'                    '   -------------------------------- �G���[���b�Z�[�W�쐬
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "�i�Ԗ��o�^", "", "")
'
'                        SendBuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                        Hin_Shogo_Proc = False
'                        GoTo Abort_Tran
'                    Case Else
'                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                        SendBuf = Text_Create_Proc()
'                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)
'                        Hin_Shogo_Proc = SYS_ERR
'                        GoTo Abort_Tran
'                End Select
'
'            Loop
'                                        '�ŏI�ƍ����t
'            Call UniCode_Conv(ITEMREC.LAST_CHK_DT, Format(Date, "yyyymmdd"))
'                                        '�ŏI�ƍ��݌ɐ�
'            Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, Format(SUMI_QTY + MI_QTY, "00000000"))
'                                        '�i�ڃ}�X�^��������
'            Do
'                sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                        Exit Do
'                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'
'                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "���Ŏg�p��", "", "")
'
'                        SendBuf = Text_Create_Proc()
'                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                        Hin_Shogo_Proc = False
'                        GoTo Abort_Tran
'
'                    Case Else
'                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
'                        SendBuf = Text_Create_Proc()
'                        Call File_Error(sts, BtOpUpdate, "�i�ڃ}�X�^", 0)
'                        Hin_Shogo_Proc = SYS_ERR
'                        GoTo Abort_Tran
'                End Select
'            Loop
        
            If (SUMI_QTY + MI_QTY) = (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                MEMO = B8_MEMO & StrConv(Format((SUMI_QTY + MI_QTY), "#0"), vbWide) & "[OK]"
            Else
                MEMO = B8_MEMO & StrConv(Format((ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_SUMI_QTY), "#0"), vbWide) & "[" & StrConv(Format(SUMI_QTY + MI_QTY, "#0"), vbWide) & "]"
            End If
                                
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
                                
                                
            sts = IDOREKI_OUTPUT_PROC(ID_KANRI_TBL(ING_No).Tanaban, _
                                        "", _
                                        ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                        ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                        ID_KANRI_TBL(ING_No).Hinban, _
                                        "", _
                                        (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                        0, _
                                        0, _
                                        (Format(ID_KANRI_TBL(ING_No).ID, "000")), _
                                        ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                        FILE_RETRY, _
                                        , _
                                        MEMO, , , , , MENU_NO)
            Select Case sts
                Case False      '����I��
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Tana_Hin_Shogo_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
    
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

    Tana_Hin_Shogo_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function MTS_Dec_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w������錾�ł̏o�׏����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    MTS_Dec_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    MTS_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    MTS_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                MTS_Dec_Proc = False
                                Exit Function
                            End If
                        End If
                            
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            MTS_Dec_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                MTS_Dec_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
            '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
            ID_NO = ""
            sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                    ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    Hinban, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE, _
                                    ID_KANRI_TBL(ING_No).SS_CODE, _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    Y_SYU_CNT, _
                                    ID_NO, _
                                    SYUKA_QTY, _
                                    DEN_NO, _
                                    KAN_KBN)
            Select Case sts
                Case False          '����
                    If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�o�ח\�薳��", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        MTS_Dec_Proc = False
                        Exit Function
                    End If
                
                Case True
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�o�ח\��g�p��", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    MTS_Dec_Proc = False
                    Exit Function
            End Select
            '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
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
                    MTS_Dec_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("�L���݌ɖ���", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                MTS_Dec_Proc = False
                Exit Function
            End If
        
                
        
        
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '�I�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Y_SYU_CNT = Y_SYU_CNT      '�Y���`�[����
            ID_KANRI_TBL(ING_No).ID_NO = ID_NO              '�`�[��
            ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO            '�`�[��
            ID_KANRI_TBL(ING_No).YUKO_SUMI_QTY = SUMI_QTY   '�g�p�\�ȏ��i���ςݍ݌�
            ID_KANRI_TBL(ING_No).YUKO_MI_QTY = MI_QTY       '�g�p�\�Ȗ����i�݌�
        
        
            Select Case Y_SYU_CNT
                Case 1              '�Ώۓ`�[���P��
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                    ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                    '���ʕt���̑��M���b�Z�[�W���쐬����
                    Send_Text.sts = Sts_OK                                      '�X�e�[�^�X�@OK
                    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                    Send_Text.Display_Flg = Display_DEF                         '�\����ʃt���O �ʏ���͉��
                    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                    Send_Text.End_Menu = Menu_Only                              '�ŏI���j���[�t���O
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                    Send_Text.Menu_Suu = "05"                                   '���j���[���ڐ��i05�Œ�j
                    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                    Send_Text.fileName = ""                                     '���M�f�[�^�t�@�C����
                    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
                    Send_Text.Buzzer = Buzzer_DEF                               '�u�U�[���@�W��
                    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                
                
                    '-----------------------------------------------�P�s��
                                                                                'BOX����
                    Send_Text.Box_Type(0).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '���l�����\��
                    Send_Text.Box_Type(0).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                                                '�����J�[�\���ʒu
                    Send_Text.Box_Type(0).Start_Pos = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '���͌���
                    Send_Text.Box_Type(0).Max_Size = "00"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                    Send_Text.Box_Type(0).MENU = ""                             '���j���\�ԍ�
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
                    Send_Text.Box_Type(2).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                            '���͌���
                    Send_Text.Box_Type(2).Max_Size = "13"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                    Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                    '-----------------------------------------------�S�s��
                                                                            'BOX����
                    Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                                            
                    If SYUKA_QTY > (SUMI_QTY + MI_QTY) Then
                        SYUKA_QTY = SUMI_QTY + MI_QTY           '�݌ɐ������Ȃ����͍݌ɐ��𑗐M
                    End If
                                                                            
                                                                            '�\�����e
                    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            
                                                                            
                                                                            '���l�����\��
                    Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
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
                
                
                Case Else           '�Ώۓ`�[��������
            
                    ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                    
                    '���ʕt���̑��M���b�Z�[�W���쐬����
                    Send_Text.sts = Sts_OK                                      '�X�e�[�^�X�@OK
                    ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                    Send_Text.Display_Flg = Display_DEF                         '�\����ʃt���O �ʏ���͉��
                    ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                    Send_Text.End_Menu = Menu_Only                              '�ŏI���j���[�t���O
                    ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                    Send_Text.Menu_Suu = "05"                                   '���j���[���ڐ��i05�Œ�j
                    ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                    Send_Text.fileName = ""                                     '���M�f�[�^�t�@�C����
                    ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
                    Send_Text.Buzzer = Buzzer_DEF                               '�u�U�[���@�W��
                    ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                
                
                    '-----------------------------------------------�P�s��
                                                                                'BOX����
                    Send_Text.Box_Type(0).Box_Type = TYPE_REF
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                    Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                '���l�����\��
                    Send_Text.Box_Type(0).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).INIT = ""
                                                            
                                                                                '�����J�[�\���ʒu
                    Send_Text.Box_Type(0).Start_Pos = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Start_Pos = ""
                                                                                '���͌���
                    Send_Text.Box_Type(0).Max_Size = "00"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Max_Size = "00"
                                                                                
                    Send_Text.Box_Type(0).MENU = ""                             '���j���\�ԍ�
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
                    Send_Text.Box_Type(2).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                            '���͌���
                    Send_Text.Box_Type(2).Max_Size = "13"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                    Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                    '-----------------------------------------------�S�s��
                                                                            'BOX����
                    Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCANK
                                                                            '�\�����e
                    Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_ID_No)
                    Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_ID_No)
                                                                            '���l�����\��
                    Send_Text.Box_Type(3).INIT = ""
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                                            '�����J�[�\���ʒu
                    Send_Text.Box_Type(3).Start_Pos = "01"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                                            '���͌���
                    Send_Text.Box_Type(3).Max_Size = "08"
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "08"
                                                                                
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
                                                                                
                    Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                    ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
                    Sendbuf = Text_Create_Proc()
            
            End Select
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�o�א��^�`�[�h�c�j
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Syuka      '�o�׎c��
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
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
                                    
                                    '�o�ɏ���
                        sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "���[���Ŏg�p��", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MTS_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
                                        '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                        '���̍�Ɨv��
                        
                        
                        '�o�ח\��^�݌ɂ̗\�����
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
            
            
            
                    Case LCD_ID_No      '�`�[�h�c
                
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                "", _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN)
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "�o�ח\�薳��", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    MTS_Dec_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                                Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Hinban, "�o�ח\��g�p��", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MTS_Dec_Proc = False
                                Exit Function
                        End Select
                
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO
                
                        '------------------ �m�肵���o�ח\��̗\�萔�𑗐M����
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo3_REQ
                        ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                        '���ʕt���̑��M���b�Z�[�W���쐬����
                        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                    
                        Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
        
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                                                                                
                        Send_Text.Box_Type(1).MENU = ""                         '���j���\�ԍ�
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
                        Send_Text.Box_Type(2).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                            'BOX����
                        Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                                            
                        If SYUKA_QTY > (ID_KANRI_TBL(ING_No).YUKO_SUMI_QTY + ID_KANRI_TBL(ING_No).YUKO_MI_QTY) Then
                            SYUKA_QTY = ID_KANRI_TBL(ING_No).YUKO_SUMI_QTY + ID_KANRI_TBL(ING_No).YUKO_MI_QTY
                        End If
                                                                            
                                                                            '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            '���l�����\��
                        Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
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
        Case Step_Sagyo3_RES        '�R��ڂ̎�M�i�o�א��j
            For i = 0 To M_Gyo - 1
            
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Syuka      '�o�׎c��
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            MTS_Dec_Proc = False
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
                                    '�o�ɏ���
                        sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "���[���Ŏg�p��", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                MTS_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
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

    MTS_Dec_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function
Private Function DEN_ID_Dec_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�`�[�h�c�ł̏o�׏����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8


Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim KAN_KBN         As String * 1

Dim MENU_NO         As String * 2

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1


    DEN_ID_Dec_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�`�[�h�c�j
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No      '�`�[�h�c
    
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            RET_JGYOBU = Left(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), 1)
                            ID_NO = Right(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), 12)
                        Else
                            RET_JGYOBU = ID_KANRI_TBL(ING_No).JGYOBU
'''                            If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                                ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "00000000")
'''                            Else
                                ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                            End If
                        End If
                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                                RET_JGYOBU, _
                                                RET_NAIGAI, _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN)
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    DEN_ID_Dec_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                                Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                DEN_ID_Dec_Proc = False
                                Exit Function
                        End Select
                
                
                        If KAN_KBN <> KAN_KBN_UN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ɏ����ς�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                                                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�2006.01.06
                        ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O2006.01.06
                           
                        
                        
                        
                        '-----------------------------------------------�w�b�_�[
                        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        
                        '-----------------------------------------------�P�s��
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "09"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "09"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
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
                        Send_Text.Box_Type(3).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                                'BOX����
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            
                                                                                '���l�����\��
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                                '���͌���
                         Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�I�ԁ^�i�ԁ^���ʁj
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban    '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    DEN_ID_Dec_Proc = False
                                    Exit Function
                    
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    DEN_ID_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                DEN_ID_Dec_Proc = False
                                Exit Function
                            End If
                        End If
                            
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                        If ID_KANRI_TBL(ING_No).JGYOBU = ID_KANRI_TBL(ING_No).S_JGYOBU And _
                            ID_KANRI_TBL(ING_No).NAIGAI = ID_KANRI_TBL(ING_No).S_NAIGAI Then
                        
                            sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        
                        Else
                        
                            sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, ID_KANRI_TBL(ING_No).S_NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        
                        End If
                        
                        
                        
                        
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            DEN_ID_Dec_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                DEN_ID_Dec_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
                    
                    
                    
                        If Hinban <> ID_KANRI_TBL(ING_No).Hinban Then
                        
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                            DEN_ID_Dec_Proc = False
                            Exit Function
                        
                        End If
                    
                    
                    Case LCD_Syuka      '�o�׎c��
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                        
                        
                        
                        '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
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
                                DEN_ID_Dec_Proc = False
                                Exit Function
                        End Select
                    
                        If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                            Call Err_Send_Proc("�L���݌ɖ���", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            DEN_ID_Dec_Proc = False
                            Exit Function
                        End If
                        
                        
                        
                        If SYUKA_QTY > (SUMI_QTY + MI_QTY) Then
                            Call Err_Send_Proc("�o�א��s��", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            DEN_ID_Dec_Proc = False
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
                                    '�o�ɏ���
                        sts = Syuko_Update_Proc(RET_JGYOBU, _
                                    RET_NAIGAI, _
                                    Hinban, _
                                    "", _
                                    Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, _
                                    MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "���[���Ŏg�p��", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                DEN_ID_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
                                        '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                        
                                        
                        '�o�ח\��^�݌ɂ̗\�����
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
    DEN_ID_Dec_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If



End Function
Private Function SYUKO_HYO_Dec_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�o�ɕ\�ł̏o�׏����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim i               As Integer

Dim Y_SYU_CNT       As Integer
Dim ID_NO           As String * 12
Dim DEN_NO          As String * 6
Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8
Dim CYU_KBN         As String * 1
Dim KAN_KBN         As String * 1

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    SYUKO_HYO_Dec_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�o�ɕ\���j
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_SYUKO_HYO_No   '�o�ɕ\��
    
                        
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ɕ\�g�p�s��", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                        
'''                        If IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
'''                            ID_NO = Format(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "000000000000")
'''                        Else
                            ID_NO = Trim(ID_KANRI_TBL(ING_No).Recv_text(i))
'''                        End If
                        '------------------ �g�p�\�ȏo�ח\��̗\����s���A�o�ח\�萔���l������
                        sts = Y_Syuka_Chek_Proc(KAN_KBN_UN, _
                                                ID_KANRI_TBL(ING_No).JGYOBU, _
                                                "", _
                                                Hinban, _
                                                MTS_CODE, _
                                                SS_CODE, _
                                                CYU_KBN, _
                                                Y_SYU_CNT, _
                                                ID_NO, _
                                                SYUKA_QTY, _
                                                DEN_NO, _
                                                KAN_KBN)
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    SYUKO_HYO_Dec_Proc = False
                                    Exit Function
                                End If
                
                            Case True
                                Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Exit Function
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                        End Select
                
                        If KAN_KBN <> KAN_KBN_UN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ɏ����ς�", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                
                        ID_KANRI_TBL(ING_No).Hinban = Hinban
                        
                        ID_KANRI_TBL(ING_No).ID_NO = ID_NO
                        ID_KANRI_TBL(ING_No).DEN_NO = DEN_NO
                        ID_KANRI_TBL(ING_No).MTS_CODE = MTS_CODE
                        ID_KANRI_TBL(ING_No).SS_CODE = SS_CODE
                        ID_KANRI_TBL(ING_No).CYU_KBN = CYU_KBN
                        
                        
                        
                        '-----------------------------------------------�w�b�_�[
                        Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
                        ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
                
                        Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
                        ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
                
                        Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
                        ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
                
                        Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
                        ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
                
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                        ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
                        ID_KANRI_TBL(ING_No).Send_Syuka_QTY = SYUKA_QTY
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                                                                                
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                                                                                '���l�����\��
                        Send_Text.Box_Type(2).INIT = ""
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(2).Start_Pos = "01"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                                '���͌���
                        Send_Text.Box_Type(2).Max_Size = "09"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "09"
                                                                                
                        Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
                        '-----------------------------------------------�S�s��
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
                        Send_Text.Box_Type(3).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                
                        Send_Text.Box_Type(3).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
                        '-----------------------------------------------�S�s��
                                                                                'BOX����
                        Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_Syuka & Space(M_Keta - (Len(LCD_Syuka) * 2) - 5 + (5 - Len(Format(SYUKA_QTY, "#0")))) & Format(SYUKA_QTY, "#0"))
                                                                            
                                                                                '���l�����\��
                        Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SYUKA_QTY, "#0"))) & Format(SYUKA_QTY, "#0")
                                                                                '�����J�[�\���ʒu
                        Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                                                '���͌���
                         Send_Text.Box_Type(4).Max_Size = "05"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
                        Send_Text.Box_Type(4).MENU = ""                         '���j���\�ԍ�
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""

                        Sendbuf = Text_Create_Proc()
    
                End Select
            Next i
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�I�ԁ^�i�ԁ^���ʁj
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_Tanaban    '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    SYUKO_HYO_Dec_Proc = False
                                    Exit Function
                    
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    SYUKO_HYO_Dec_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                            End If
                        End If
                            
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            SYUKO_HYO_Dec_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
                    
                    
                    
                        If Hinban <> ID_KANRI_TBL(ING_No).Hinban Then
                        
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        
                        End If
                    
                    
                    Case LCD_Syuka      '�o�׎c��
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                
                        SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If SYUKA_QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
            
                        If SYUKA_QTY > ID_KANRI_TBL(ING_No).Send_Syuka_QTY Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                        
                        '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
                        sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
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
                                SYUKO_HYO_Dec_Proc = False
                                Exit Function
                        End Select
                    
                        If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                            Call Err_Send_Proc("�L���݌ɖ���", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
                            Exit Function
                        End If
                        
                        
                        If SYUKA_QTY > (SUMI_QTY + MI_QTY) Then
                            Call Err_Send_Proc("�o�א��s��", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            SYUKO_HYO_Dec_Proc = False
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
                                    
                                    
                                    
                                    '�o�ɏ���
                        sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                    ID_KANRI_TBL(ING_No).NAIGAI, _
                                    Hinban, _
                                    "", _
                                    Tanaban, _
                                    ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE, _
                                    0, _
                                    0, _
                                    SYUKA_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).CYU_KBN, _
                                    ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).DEN_NO, _
                                    ID_KANRI_TBL(ING_No).ID_NO, MENU_NO)
                        Select Case sts
                            Case False
                            Case SYS_CANCEL
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "���[���Ŏg�p��", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                SYUKO_HYO_Dec_Proc = False
                                GoTo Abort_Tran
                            Case SYS_ERR
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpEndTransaction, "", 0)
                                GoTo Abort_Tran
                        End Select
    
                                        '�g�����U�N�V�����I��
                        sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        If sts <> BtNoErr Then
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpEndTransaction, "", 0)
                            GoTo Abort_Tran
                        End If
                                        
                                        
                        '�o�ח\��^�݌ɂ̗\�����
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
    SYUKO_HYO_Dec_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If



End Function
Private Function Out_Plan_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�v��O�i�f�[�^�����j�̏o�׏����x
'
'-------------------------------------------------------
Dim i               As Integer
Dim Hinban          As String * 13
Dim Tanaban         As String * 8
Dim sts             As Integer

Dim SYUKA_QTY       As Long
Dim SUMI_QTY        As Long
Dim MI_QTY          As Long

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2
    
    Out_Plan_Proc = True

    For i = 0 To M_Gyo - 1
        Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
            Case LCD_Tanaban        '�I��
                Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                
                If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                Else
                '------------------ �q�Ƀ}�X�^�Ǎ���
                    Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                            Out_Plan_Proc = False
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
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                             
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                             
                             Out_Plan_Proc = False
                            Exit Function
                        End If
                    End If
                    '------------------ �I�}�X�^�Ǎ���
                    Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                    Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                    Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                    Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                
                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            Out_Plan_Proc = False
                            Exit Function
                        Case Else
                            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                            Sendbuf = Text_Create_Proc()
                            Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                            Exit Function
                    End Select
            
                    '------------------ �֎~�I�̃`�F�b�N
                    If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                        Out_Plan_Proc = False
                        Exit Function
                    End If
            
                End If
            Case LCD_Hinban         '�i��
                Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                '------------------ �i�ڃ}�X�^�Ǎ���
'                Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                Select Case sts
                    Case BtNoErr
                        If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Out_Plan_Proc = False
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
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                        Out_Plan_Proc = False
                        Exit Function
                    Case Else
                        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                        Exit Function
                End Select
            
            Case LCD_Suryo          '����
                If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Out_Plan_Proc = False
                    Exit Function
                End If
                
                SYUKA_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                If SYUKA_QTY = 0 Then
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                    Out_Plan_Proc = False
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
                                        
    If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
        MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                
    Else
                
        MENU_NO = ""
    End If
                                        
                                        '�o�ɍX�V
    sts = Syuko_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                            ID_KANRI_TBL(ING_No).NAIGAI, _
                            Hinban, _
                            "", _
                            Tanaban, _
                            (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                            SUMI_QTY, _
                            MI_QTY, _
                            SYUKA_QTY, _
                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                            FILE_RETRY, _
                            "", _
                            ID_KANRI_TBL(ING_No).CYU_KBN, _
                            ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE, _
                            Format(Now, "YYYYMMDD"), _
                            "", _
                            "", MENU_NO)

    Select Case sts
        Case False
        
        Case True       '�݌ɕs�����ɔ���
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "�݌ɐ��s��", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Out_Plan_Proc = False
            GoTo Abort_Tran
        Case SYS_CANCEL
            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
            Sendbuf = Text_Create_Proc()
            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
            Out_Plan_Proc = False
            GoTo Abort_Tran
        Case SYS_ERR
            Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
            Sendbuf = Text_Create_Proc()
            Out_Plan_Proc = SYS_ERR    '�V�X�e���ُ픭��
            
            GoTo Abort_Tran
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
    
    
    
    Out_Plan_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Inspe_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
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


Dim MENU_NO         As String * 2

    Inspe_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�`�[�h�c�j
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_ID_No      '�`�[�h�c
    
    
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���i��ƕs��", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
    
    
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
                                                JITU_QTY)
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc = False
                                Exit Function
                        End Select
                
                        '------------------ ������̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "������G���[", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ �����敪�̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�����敪�~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc = False
                            Exit Function
                        End If
                        '------------------ �o�Ɋ����̃`�F�b�N
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "��Ɩ�����", "", "")
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
                
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
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
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Hinban     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
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
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")
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
                
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                        Call UniCode_Conv(Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
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
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")
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
                                            
                                            '�o�ח\�菑����
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")
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
                                        ID_KANRI_TBL(ING_No).ID_NO)
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
Private Function Inspe_Proc_MTS(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i�����i�l�s�r�ǂݍ��݂���j�x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13
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

Dim MENU_NO         As String * 2

    Inspe_Proc_MTS = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i������j
        
            For i = 0 To M_Gyo - 1
                
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_MTS    '������
                                
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) < 16 Then
                                    '������i���Ӑ�j�݂̂Ō�����}�X�^�ǂݍ���
                            Call UniCode_Conv(K2_MTS.MUKE_CODE, ID_KANRI_TBL(ING_No).Recv_text(i))
                            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
                            Select Case sts
                                Case BtNoErr
                                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                                    
                                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "")
                    
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
                                        
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "�o�א�G���[", "", "")
                    
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
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, MTS_CODE & SS_CODE, "�o�א�G���[", "", "")
                    
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
                
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), _
                            Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - _
                            CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    Case LCD_ID_No      '�`�[�h�c
    
    
                        If Len(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) = 13 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).Recv_text(i), "���i��ƕs��", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
    
    
    
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
                                                JITU_QTY)
                        Select Case sts
                            Case False          '����
                                If Y_SYU_CNT = 0 Then   '�Ώۃf�[�^�Ȃ�
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\�薳��", "", "")
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�o�ח\��g�p��", "", "")
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                Sendbuf = Text_Create_Proc()
                                Inspe_Proc_MTS = False
                                Exit Function
                        End Select
                
                        '------------------ ������̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).MTS_CODE <> MTS_CODE Or _
                            ID_KANRI_TBL(ING_No).SS_CODE <> SS_CODE Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "������G���[", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ �����敪�̃`�F�b�N
                        If ID_KANRI_TBL(ING_No).CYU_KBN <> CYU_KBN Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "�����敪�~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
                        End If
                        '------------------ �o�Ɋ����̃`�F�b�N
                        If Inspection_Flg = 0 Then
                            If KAN_KBN <> KAN_KBN_FIN Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_NO, "��Ɩ�����", "", "")
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
                
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                                                                                'BOX����
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
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
                        Send_Text.Box_Type(3).Max_Size = "13"
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "13"
                                                                                
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
            
                
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
                    
                    Case LCD_Hinban     '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                        
                        
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                    
                            Case BtErrKeyNotFound
                            '   -------------------------------- �G���[���b�Z�[�W�쐬
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
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
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�i�ԃG���[", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                            Inspe_Proc_MTS = False
                            Exit Function
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
                
                        Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                        ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
                
                        Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                        ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                        
                        '-----------------------------------------------�P�s��
                        Send_Text.Box_Type(0).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, ID_KANRI_TBL(ING_No).MTS_CODE & ID_KANRI_TBL(ING_No).SS_CODE)
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
                        Send_Text.Box_Type(2).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, "�`�[ID:" & ID_KANRI_TBL(ING_No).ID_NO)
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
                        Send_Text.Box_Type(4).Box_Type = TYPE_REF
                        ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_REF
                                                                                '�\�����e
                        Call UniCode_Conv(Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
                        Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, "�o�א��F" & StrConv(Format(ID_KANRI_TBL(ING_No).SYUKA_QTY, "#0"), vbWide))
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
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)     '���ƕ�
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_KANRI_TBL(ING_No).ID_NO)   'ID��
    
            Do
            
                sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrKeyNotFound
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��s��", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Inspe_Proc_MTS = False
                        GoTo Abort_Tran
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")
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
                                            
                                            '�o�ח\�菑����
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, ID_KANRI_TBL(ING_No).ID_NO, "�o�ח\��g�p��", "", "")
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
                                        ID_KANRI_TBL(ING_No).ID_NO)
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
        
                Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
                ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
        
                Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
                ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                
                '-----------------------------------------------�P�s��
                                                                'BOX����
                Send_Text.Box_Type(0).Box_Type = TYPE_REF
                ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                                '�\�����e
                Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
Private Function Zaiko_Reserve_Proc(ID_NO As Integer, FROM_LOCATION As String, JGYOBU As String, NAIGAI As String, Hinban As String, SUMI_QTY As Long, MI_QTY As Long) As Integer
'-------------------------------------------------------
'
'   �w�݌Ƀf�[�^�̎g�p�\��x
'
'-------------------------------------------------------
Dim sts             As Integer

    Zaiko_Reserve_Proc = True
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Zaiko_Reserve_Proc = SYS_ERR
        Exit Function
    End If

    sts = Zaiko_Lock_Proc(FROM_LOCATION, JGYOBU, NAIGAI, Hinban, Format(ID_NO, "000"), SUMI_QTY, MI_QTY, FILE_RETRY)
    If sts Then
        Zaiko_Reserve_Proc = sts
        GoTo Abort_Tran
    End If
End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Zaiko_Reserve_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    Zaiko_Reserve_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Y_Syuka_Chek_Proc(Mode As String, _
                                        JGYOBU As String, _
                                        NAIGAI As String, _
                                        Hinban As String, _
                                        MTS_CODE As String, _
                                        SS_CODE As String, _
                                        CYU_KBN As String, _
                                        Y_SYU_CNT As Integer, _
                                        ID_NO As String, _
                                        SYUKA_QTY As Long, _
                                        DEN_NO As String, _
                                        KAN_KBN As String, _
                                        Optional JITU_QTY) As Integer
'-------------------------------------------------------
'
'   �w�P��o�ח\��̎g�p�\��x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer


Dim ans         As Integer

Dim RETRY_CNT   As Integer

Dim WK_ID_NO    As String * 12
Dim WK_DEN_NO   As String * 6




    Y_Syuka_Chek_Proc = True
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Y_Syuka_Chek_Proc = SYS_ERR
        Exit Function
    End If


    WK_ID_NO = ""
    WK_DEN_NO = ""

    Y_SYU_CNT = 0
    
    If Len(Trim(ID_NO)) <> 0 Then
        '�`�[�h�c�w��ł̏���
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)              '���ƕ�
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, ID_NO)            'ID��
    
        RETRY_CNT = 0
    
        Do
            DoEvents
            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    '�f�[�^�Ȃ�
                    Y_Syuka_Chek_Proc = False
                    GoTo Abort_Tran
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                        Y_Syuka_Chek_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                    Y_Syuka_Chek_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
                                                
                                                
        If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) = 0 And _
            Len(Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode))) = 0 Then
        Else
            If StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> Format(ID_KANRI_TBL(ING_No).ID, "000") Or _
                Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) <> StrConv(App.EXEName, vbUpperCase) Then
                                '���Ŏg�p��
                Y_Syuka_Chek_Proc = SYS_CANCEL
                GoTo Abort_Tran
            End If
        End If
            
        Call UniCode_Conv(Y_SYUREC.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
        Call UniCode_Conv(Y_SYUREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
                                            
        RETRY_CNT = 0
                                            
                                            '�o�ח\�菑����
        Do
            DoEvents
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                        Y_Syuka_Chek_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                
                Case Else
                    Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                    Y_Syuka_Chek_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
                                
                                '1���̂ݓ`�[�����o�א�KEEP
        SYUKA_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
        
        WK_ID_NO = StrConv(Y_SYUREC.ID_NO, vbUnicode)
        WK_DEN_NO = Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6)
        Y_SYU_CNT = 1
        
        NAIGAI = StrConv(Y_SYUREC.NAIGAI, vbUnicode)
        Hinban = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
        MTS_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)
        SS_CODE = StrConv(Y_SYUREC.SS_CODE, vbUnicode)
        CYU_KBN = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
        KAN_KBN = StrConv(Y_SYUREC.KAN_KBN, vbUnicode)
        JITU_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
'        SYUKA_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode))
    
    
    
    Else
        '�����敪�^������^�i�Ԃł̏���
        Call UniCode_Conv(K3_Y_SYU.JGYOBU, JGYOBU)              '���ƕ�
        Call UniCode_Conv(K3_Y_SYU.KEY_CYU_KBN, CYU_KBN)        '�����敪
        Call UniCode_Conv(K3_Y_SYU.KEY_MUKE_CODE, MTS_CODE)     '���Ӑ�R�[�h
        Call UniCode_Conv(K3_Y_SYU.KEY_SS_CODE, SS_CODE)        '���Ӑ�R�[�h
        Call UniCode_Conv(K3_Y_SYU.NAIGAI, NAIGAI)              '�����O
        Call UniCode_Conv(K3_Y_SYU.KEY_HIN_NO, Hinban)          '�i��
        Call UniCode_Conv(K3_Y_SYU.KEY_ID_NO, "")               'ID��
    
        com = BtOpGetGreaterEqual
    
    
        Do
            RETRY_CNT = 0
            Do
                DoEvents
                sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                Select Case sts
                    Case BtNoErr
                        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> JGYOBU Or _
                            StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> CYU_KBN Or _
                            Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) <> Trim(MTS_CODE) Or _
                            Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) <> Trim(SS_CODE) Or _
                            StrConv(Y_SYUREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                            Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) <> Trim(Hinban) Then

                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "�o�ח\��", 0)
                                Y_Syuka_Chek_Proc = SYS_ERR
                                GoTo Abort_Tran
                            End If
                        
                            sts = BtErrEOF
                    
                        End If
                                        
                    
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > FILE_RETRY Then
                            Y_Syuka_Chek_Proc = SYS_CANCEL
                            GoTo Abort_Tran
                        End If
                   Case Else
                        Call File_Error(sts, com + BtSNoWait, "�o�ח\��", 0)
                        Y_Syuka_Chek_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
                                        
            If sts = BtErrEOF Then
                Exit Do
            End If
        
        
'            If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) <> KAN_KBN_FIN And _
'                  Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) = 0 Then
                                            
            If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) = 0 And _
                Len(Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode))) = 0 Then
            Else
                If StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> Format(ID_KANRI_TBL(ING_No).ID, "000") Or _
                    Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) <> StrConv(App.EXEName, vbUpperCase) Then
                                            '���Ŏg�p��
                    Y_Syuka_Chek_Proc = SYS_CANCEL
                    GoTo Abort_Tran
                End If
            End If
                    
            Call UniCode_Conv(Y_SYUREC.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
            Call UniCode_Conv(Y_SYUREC.PRG_ID, StrConv(App.EXEName, vbUpperCase))
                                            
            RETRY_CNT = 0
                                            
                                            '�o�ח\�菑����
            Do
                DoEvents
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K3_Y_SYU, Len(K3_Y_SYU), 3)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > FILE_RETRY Then
                            Y_Syuka_Chek_Proc = SYS_CANCEL
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                        Y_Syuka_Chek_Proc = SYS_ERR
                        GoTo Abort_Tran
                End Select
            Loop
                                
            
            If StrConv(Y_SYUREC.KAN_KBN, vbUnicode) = Mode Then
                        
            
                Y_SYU_CNT = Y_SYU_CNT + 1
                If Y_SYU_CNT > 1 Then
                                            '�����`�[����
                    Y_Syuka_Chek_Proc = False
                    GoTo Abort_Tran
            
                End If
                                '1���̂ݓ`�[�����o�א�KEEP
                SYUKA_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode))
        
                WK_ID_NO = StrConv(Y_SYUREC.ID_NO, vbUnicode)
                WK_DEN_NO = Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6)
                                   
            End If
            
            com = BtOpGetNext
    
        Loop
    
    End If


    If Len(Trim(WK_ID_NO)) <> 0 Then
        ID_NO = WK_ID_NO
        DEN_NO = WK_DEN_NO
    End If

End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Y_Syuka_Chek_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    Y_Syuka_Chek_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If



End Function
Private Function Cancel_Proc(Sendbuf As String) As Integer
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
            ID_KANRI_TBL(ING_No).NAIGAI = ""
            
'            ID_KANRI_TBL(ING_No).MENU_GRP = ""
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            



        Case Step_NAIGAI_REQ    '�����O�v��
            
            ID_KANRI_TBL(ING_No).Step = Step_JGYOBU_REQ
            ID_KANRI_TBL(ING_No).JGYOBU = ""
            ID_KANRI_TBL(ING_No).NAIGAI = ""
            
'            ID_KANRI_TBL(ING_No).MENU_GRP = ""
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            ID_KANRI_TBL(ING_No).MENU_LV2 = ""
'            ID_KANRI_TBL(ING_No).MENU_LV3 = ""
            
            
            Call Menu_Send_Proc(Sendbuf)


        Case Step_MENU1_REQ     '���j���[�P�v��
        
            If UBound(NAIGAI) = 0 Then
                '�����O�̐؂蕪���Ȃ�
'                ID_KANRI_TBL(ING_No).Step = Step_TANTO_REQ
'                ID_KANRI_TBL(ING_No).JGYOBU = ""
'                Call Start_Proc(Sendbuf)
            
                ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                ID_KANRI_TBL(ING_No).MENU_LV1 = ""
            
            
                Call Menu_Send_Proc(Sendbuf)
            
            Else
                ID_KANRI_TBL(ING_No).Step = Step_NAIGAI_REQ
                ID_KANRI_TBL(ING_No).NAIGAI = ""
                Call Menu_Send_Proc(Sendbuf)
            End If
        
        Case Step_MENU2_REQ     '���j���[�Q�v��
        
            ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
            ID_KANRI_TBL(ING_No).MENU_LV1 = ""
        
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
                If Len(Trim(ID_KANRI_TBL(ING_No).MENU_LV2)) <> 0 Then
                    ID_KANRI_TBL(ING_No).Step = Step_MENU2_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV2 = ""
                    Call Menu_Send_Proc(Sendbuf)
                Else
                    ID_KANRI_TBL(ING_No).Step = Step_MENU1_REQ
                    ID_KANRI_TBL(ING_No).MENU_LV1 = ""
                    Call Menu_Send_Proc(Sendbuf)
                End If
'2006.01.30            End If
                                                    '��ƂQ�^��ƂR�^��ƂS�v��
        Case Step_Sagyo2_REQ, Step_Sagyo3_REQ, Step_Sagyo4_REQ
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo1_REQ
            If Sagyo_Send_Proc() Then
                Sendbuf = Text_Create_Proc()
                Exit Function
            End If
            
            Sendbuf = Text_Create_Proc()
        
    
    End Select
    
    Cancel_Proc = False


End Function

Private Function Data_Clear_Proc(Mode As Integer, Sendbuf As String) As Integer
'-------------------------------------------------------
'
'   �w�o�ח\��^�݌ɂ̗\��L�����Z���x
'
'-------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim RETRY_CNT   As Integer
    
    
    Data_Clear_Proc = True
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
        Data_Clear_Proc = SYS_ERR
        Exit Function
    End If
    
    If Mode = 0 Then
                                        '�o�ח\��̊J��
        Call UniCode_Conv(K4_Y_SYU.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
        Call UniCode_Conv(K4_Y_SYU.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        com = BtOpGetGreaterEqual
    Else
        Call UniCode_Conv(K4_Y_SYU.WEL_ID, "")
        Call UniCode_Conv(K4_Y_SYU.PRG_ID, "")
        com = BtOpGetGreater
    End If

    Do
        DoEvents
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
            Select Case sts
                                
                Case BtNoErr
                    If Mode = 0 Then
                        If Format(ID_KANRI_TBL(ING_No).ID, "000") <> StrConv(Y_SYUREC.WEL_ID, vbUnicode) Or _
                             StrConv(App.EXEName, vbUpperCase) <> Trim(StrConv(Y_SYUREC.PRG_ID, vbUnicode)) Then
                            sts = BtErrEOF
                        
                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
                            If sts Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                                Data_Clear_Proc = SYS_ERR
                                GoTo Abort_Tran
                            End If
                        
                        End If
                    End If
                    
                    Exit Do
                
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("�o�׎g�p��", "", "", "", "")
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        RETRY_CNT = 0
        
        Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
        Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
        Do
        
            sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("�o�׎g�p��", "", "", "", "")
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�o�ח\��", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop


    If Mode = 0 Then
                                        '�݌ɂ̊J��
        Call UniCode_Conv(K3_ZAIKO.WEL_ID, Format(ID_KANRI_TBL(ING_No).ID, "000"))
        Call UniCode_Conv(K3_ZAIKO.PRG_ID, StrConv(App.EXEName, vbUpperCase))
        com = BtOpGetGreaterEqual
    Else
        Call UniCode_Conv(K3_ZAIKO.WEL_ID, "")
        Call UniCode_Conv(K3_ZAIKO.PRG_ID, "")
        com = BtOpGetGreater
    End If
    
    Do
        DoEvents
        
        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    
                    If Mode = 0 Then
                        If Format(ID_KANRI_TBL(ING_No).ID, "000") <> StrConv(ZAIKOREC.WEL_ID, vbUnicode) Or _
                             StrConv(App.EXEName, vbUpperCase) <> Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)) Then
                            sts = BtErrEOF
                        
                        
                            sts = BTRV(BtOpUnlock, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
                            If sts Then
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                                Data_Clear_Proc = SYS_ERR
                                GoTo Abort_Tran
                            End If
                        
                        
                        End If
                    End If
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("�݌Ɏg�p��", "", "", "", "")
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Sendbuf = Text_Create_Proc()
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
    
        
        RETRY_CNT = 0
        
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")
        Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)
        Do
        
            sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    RETRY_CNT = RETRY_CNT + 1
                    If RETRY_CNT > FILE_RETRY Then
                        Call Err_Send_Proc("�݌Ɏg�p��", "", "", "", "")
                        Sendbuf = Text_Create_Proc()
                        ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                        Data_Clear_Proc = SYS_CANCEL
                        GoTo Abort_Tran
                    End If
                        
                    DoEvents
                
                Case Else
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�݌Ƀf�[�^", 0)
                    Data_Clear_Proc = SYS_ERR
                    GoTo Abort_Tran
            End Select
        
        Loop
    
        com = BtOpGetNext
    
    Loop
End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "", 0)
        Data_Clear_Proc = SYS_ERR
        GoTo Abort_Tran
    End If
    
    Data_Clear_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function
Private Function tmpZaiko_Clear_Proc() As Integer
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
        DoEvents
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
                        
                    DoEvents
                
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
                        
                    DoEvents
                
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

Private Function Item_Read_Proc(JGYOBU As String, NAIGAI As String, Hinban As String, RET_JGYOBU As String, RET_NAIGAI As String) As Integer
'-------------------------------------------------------
'
'   �w�i�ڃ}�X�^�x�̓ǂݍ��ݏ���
'
'   �u�O���i�ԁv��[�i����]�ˁu�ǂݑւ��R�[�h�v�����ɓǂݍ���
'
'   �Ԃ�l
'       BtNoErr             :����I��
'       BtErrKeyNotFound    :���o�^
'       ��L�ȊO            :Pervasive ���^�[���R�[�h
'
'-------------------------------------------------------
Dim sts As Integer

    
    '--------------------------------------------------�O���i��
    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case BtErrKeyNotFound
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select
    '--------------------------------------------------�i�����R�[�h
    Call UniCode_Conv(K4_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K4_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K4_ITEM.JAN_CODE, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K4_ITEM, Len(K4_ITEM), 4)
    Select Case sts
        Case BtNoErr
            Hinban = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case BtErrKeyNotFound
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select
    '--------------------------------------------------�Ǒւ��R�[�h
    Call UniCode_Conv(K5_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K5_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K5_ITEM.HIN_CHANGE, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K5_ITEM, Len(K5_ITEM), 5)
    Select Case sts
        Case BtNoErr
            Hinban = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case BtErrKeyNotFound   '2006.01.06 '���ޕi�Ԃł̓ǂݑւ���ǉ�
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select
    '--------------------------------------------------���ޕi�Ԃœǂݑւ�
    Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
    
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Hinban = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            RET_JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
            RET_NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
            Item_Read_Proc = sts
            Exit Function
        Case Else
            Item_Read_Proc = sts
            Exit Function
    End Select

End Function


Private Function GOODS_ONOFF_Ono_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i���ς݁������i�̐؂�ւ��i����p�j�x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim MENU_NO         As String * 2

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

    GOODS_ONOFF_Ono_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    GOODS_ONOFF_Ono_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                GOODS_ONOFF_Ono_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            GOODS_ONOFF_Ono_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                GOODS_ONOFF_Ono_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
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
                    GOODS_ONOFF_Ono_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("�L���݌ɖ���", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                GOODS_ONOFF_Ono_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '�I�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i����
                                                        
            '���ʕt���̑��M���b�Z�[�W���쐬����
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, "���݌ɁF" & Format((SUMI_QTY + MI_QTY), "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, "���݌ɁF" & Format((SUMI_QTY + MI_QTY), "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(3).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = ""
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(3).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Start_Pos = "01"
                                                            '���͌���
            Send_Text.Box_Type(3).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(3).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).MENU = ""
            
            
            '-----------------------------------------------�T�s��
                                                            'BOX����
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���i�^�����i���ʁj
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '���ʁi�����͖����j
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Ono_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Ono_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Ono_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                        End If
        
                        If i = M_Gyo - 1 Then       '�ŏI�s��������
'                            If SUMI_QTY = 0 And MI_QTY = 0 Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "���i�^�����i���O", "���ʓ��̓~�X", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                GOODS_ONOFF_Siga_Proc = False
'                                Exit Function
'                            End If
                
                
                            MI_QTY = (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) - SUMI_QTY
                
                
                            If (SUMI_QTY + MI_QTY) <> (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�����ʕύX�s��", "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                
                                
                                
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                GOODS_ONOFF_Ono_Proc = False
                                Exit Function
                    
                            End If
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
        
        
        
            '���i�����������i�̐؂�ւ��X�V
            sts = GOODS_ONOFF_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                            ID_KANRI_TBL(ING_No).NAIGAI, _
                                            ID_KANRI_TBL(ING_No).Hinban, _
                                            ID_KANRI_TBL(ING_No).Tanaban, _
                                            SUMI_QTY, _
                                            MI_QTY, _
                                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                            FILE_RETRY)
            
            Select Case sts
                Case False

                Case True       '�݌ɕs�����ɔ���
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "�݌ɐ��s��", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                    GOODS_ONOFF_Ono_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GOODS_ONOFF_Ono_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GOODS_ONOFF_Ono_Proc = SYS_ERR    '�V�X�e���ُ픭��
                    GoTo Abort_Tran
            End Select
   
                                        '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '�o�ח\��^�݌ɂ̗\�����
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

    GOODS_ONOFF_Ono_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


End Function

Private Function GOODS_ONOFF_Siga_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w���i���ς݁������i�̐؂�ւ��i����p�j�x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

    GOODS_ONOFF_Siga_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    GOODS_ONOFF_Siga_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                GOODS_ONOFF_Siga_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            GOODS_ONOFF_Siga_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                GOODS_ONOFF_Siga_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, Tanaban, ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, SUMI_QTY, MI_QTY)
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
                    GOODS_ONOFF_Siga_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("�L���݌ɖ���", Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                GOODS_ONOFF_Siga_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = Tanaban          '�I�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i����
                                                        
            '���ʕt���̑��M���b�Z�[�W���쐬����
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
            Send_Text.Box_Type(2).Max_Size = "13"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "13"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#0")))) & Format(SUMI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#0"))) & Format(SUMI_QTY, "#0")
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
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#0")))) & Format(MI_QTY, "#0"))
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#0"))) & Format(MI_QTY, "#0")
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���i�^�����i���ʁj
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_Suryo          '���ʁi�����͖����j
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Siga_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Siga_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            GOODS_ONOFF_Siga_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                            
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
                        End If
        
                        If i = M_Gyo - 1 Then       '�ŏI�s��������
'                            If SUMI_QTY = 0 And MI_QTY = 0 Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "���i�^�����i���O", "���ʓ��̓~�X", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                GOODS_ONOFF_Siga_Proc = False
'                                Exit Function
'                            End If
                
                
                
                            If (SUMI_QTY + MI_QTY) <> (ID_KANRI_TBL(ING_No).Send_SUMI_QTY + ID_KANRI_TBL(ING_No).Send_MI_QTY) Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�����ʕύX�s��", "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                
                                
                                
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                GOODS_ONOFF_Siga_Proc = False
                                Exit Function
                    
                            End If
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
        
            '���i�����������i�̐؂�ւ��X�V
            sts = GOODS_ONOFF_Update_Proc(ID_KANRI_TBL(ING_No).JGYOBU, _
                                            ID_KANRI_TBL(ING_No).NAIGAI, _
                                            ID_KANRI_TBL(ING_No).Hinban, _
                                            ID_KANRI_TBL(ING_No).Tanaban, _
                                            SUMI_QTY, _
                                            MI_QTY, _
                                            Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                            ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                            FILE_RETRY)
            
            Select Case sts
                Case False

                Case True       '�݌ɕs�����ɔ���
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "�݌ɐ��s��", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1

                    GOODS_ONOFF_Siga_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    GOODS_ONOFF_Siga_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    GOODS_ONOFF_Siga_Proc = SYS_ERR    '�V�X�e���ُ픭��
                    GoTo Abort_Tran
            End Select
   
                                        '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '�o�ח\��^�݌ɂ̗\�����
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

    GOODS_ONOFF_Siga_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Private Function GOODS_ONOFF_Update_Proc(JGYOBU As String, _
                                        NAIGAI As String, _
                                        HIN_GAI As String, _
                                        LOCATION As String, _
                                        SUMI_JITU_QTY As Long, _
                                        MI_JITU_QTY As Long, _
                                        ID As String, _
                                        TANTO_CODE As String, _
                                        Optional RETRY As Integer = 10) As Integer
'****************************************************
'*      �u���i���^�����i�؂�ւ������v�݌Ƀf�[�^�X�V
'*
'*  �݌Ƀf�[�^�̍X�V���s���B
'*  (�����̐ݒ�~�X�͂�����ł̓`�F�b�N���Ȃ�)
'*  �g�p̧��    :   �݌Ƀf�[�^
'*                  �݌Ƀf�[�^(�ꎞ�t�@�C��)
'*  �����F  ���ƕ��i�ȗ��s�j
'*          �����O�i�ȗ��s�j
'*          �i�ԊO��(�ȗ��s��)
'*          �I�ԁiXXXXXXXX(�q�ɇ�+��+�A+�i)�ȗ��s�j
'*          ���i���ςݎ��ѐ��i���ꂩ����K�{�j
'*          �����i���ѐ��@�@�i�@�@�V�@�@�@�@�j
'*          ID(�ȗ��s��)
'*          �S���ҁi�ȗ��s�j
'*          ���g���C(�ȗ��� �P����:1=��ʃ��b�Z�[�W�L 0:���C�Q����:���g���C��(0�`9 0:����))
'*  �߂�l: false       :����
'*          true        :�p���\�Ȉُ�
'*          SYS_ERR     :�p���ł��Ȃ��ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts         As Integer
Dim com         As Integer
Dim Upd_com     As Integer


Dim RETRY_CNT   As Integer
Dim MESG_FLG    As Integer
Dim RETRY_SU    As Integer
    
Dim Zan_Qty     As Long
Dim WK_Qty      As Long
    
    

    GOODS_ONOFF_Update_Proc = True
                                                                      
                                                                      
    MESG_FLG = CInt(Mid(Format(RETRY, "00"), 1, 1))
    RETRY_SU = CInt(Mid(Format(RETRY, "00"), 2, 1))
    
'============================================================ �Ώۍ݌Ƀf�[�^���݌Ɉꎞ�f�[�^�ɑS���ړ�����B
    
    Call UniCode_Conv(K4_ZAIKO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K4_ZAIKO.Soko_No, Left(LOCATION, 2))
    Call UniCode_Conv(K4_ZAIKO.Retu, Mid(LOCATION, 3, 2))
    Call UniCode_Conv(K4_ZAIKO.Ren, Mid(LOCATION, 5, 2))
    Call UniCode_Conv(K4_ZAIKO.Dan, Right(LOCATION, 2))
    
    
    com = BtOpGetGreaterEqual
    
    RETRY_CNT = 0
    
    
    Do
        DoEvents
'------- ���݌ɓǍ���
        Do
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
            Select Case sts
                Case BtNoErr
                    If JGYOBU <> StrConv(ZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Or _
                        LOCATION <> (StrConv(ZAIKOREC.Soko_No, vbUnicode) & _
                                        StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                        StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                        StrConv(ZAIKOREC.Dan, vbUnicode)) Then
                        sts = BtErrEOF
                    
                    End If
                
                    Exit Do
                
                Case BtErrEOF
    
                    Exit Do
                
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function
    
                        End If
    
                    End If
    
                    DoEvents
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function
            End Select
    
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
'------- �݌�(�ꎞ�f�[�^)�Ǎ���
        Call UniCode_Conv(K0_tmpZAIKO.Soko_No, StrConv(ZAIKOREC.Soko_No, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.Retu, StrConv(ZAIKOREC.Retu, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.Ren, StrConv(ZAIKOREC.Ren, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.Dan, StrConv(ZAIKOREC.Dan, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K0_tmpZAIKO.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
    
        
        Do
        
            sts = BTRV(BtOpGetEqual + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Upd_com = BtOpUpdate
                    Exit Do
                
                Case BtErrKeyNotFound
                    Upd_com = BtOpInsert
    
                    Exit Do
                
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then
    
                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function
    
                        End If
    
                    End If
    
                    DoEvents
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function
            End Select
        
        
        Loop
'------- �݌�(�ꎞ�f�[�^)���o��
        Select Case Upd_com
        
            Case BtOpInsert
            '------- �V�K�ǉ�
                Do
                    sts = BTRV(BtOpInsert, tmpZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                
                
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then
        
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                    Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j", 0)
                                    GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                    Exit Function
        
                                End If
        
                            End If
        
                            DoEvents
                        Case Else
                            Call File_Error(sts, Upd_com, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                            GOODS_ONOFF_Update_Proc = SYS_ERR
                            Exit Function
        
                    End Select
                
                Loop
    
        
            Case BtOpUpdate
            '------- �݌ɐ����Z�i�X�V�j
                Call UniCode_Conv(tmpZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode)) + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "00000000"))
        
        
                Do
                    sts = BTRV(BtOpUpdate, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                
                
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                        '���g���C�񐔃`�F�b�N
                            If RETRY_SU <> 0 Then
        
                                RETRY_CNT = RETRY_CNT + 1
                                If RETRY_CNT > RETRY_SU Then
                                        '�񐔃I�[�o�[
                                    Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j", 0)
                                    GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                    Exit Function
        
                                End If
        
                            End If
        
                            DoEvents
                        Case Else
                            Call File_Error(sts, Upd_com, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                            GOODS_ONOFF_Update_Proc = SYS_ERR
                            Exit Function
        
                    End Select
                
                Loop
        
        End Select
   
'------- ���݌ɍ폜
        Do
            sts = BTRV(BtOpDelete, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
        
        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, BtOpDelete, "�݌Ƀf�[�^", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    DoEvents
                Case Else
                    Call File_Error(sts, Upd_com, "�݌Ƀf�[�^")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        
        Loop
    
    
        com = BtOpGetNext
    
    Loop
    
'============================================================ ���i���ς݂̏���(�Â����t���������Ă�)
    If SUMI_JITU_QTY <> 0 Then
    
        Zan_Qty = SUMI_JITU_QTY




        Call UniCode_Conv(K0_tmpZAIKO.Soko_No, Left(LOCATION, 2))
        Call UniCode_Conv(K0_tmpZAIKO.Retu, Mid(LOCATION, 3, 2))
        Call UniCode_Conv(K0_tmpZAIKO.Ren, Mid(LOCATION, 5, 2))
        Call UniCode_Conv(K0_tmpZAIKO.Dan, Right(LOCATION, 2))
        Call UniCode_Conv(K0_tmpZAIKO.JGYOBU, JGYOBU)
        Call UniCode_Conv(K0_tmpZAIKO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K0_tmpZAIKO.HIN_GAI, HIN_GAI)
        Call UniCode_Conv(K0_tmpZAIKO.NYUKA_DT, "")

        
        com = BtOpGetGreaterEqual
        
        Do

            RETRY_CNT = 0
'------- �݌Ɂi�ꎞ�f�[�^�j�Ǎ���
            Do
                sts = BTRV(com + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                    
                    
                        If JGYOBU <> StrConv(tmpZAIKOREC.JGYOBU, vbUnicode) Or _
                            NAIGAI <> StrConv(tmpZAIKOREC.NAIGAI, vbUnicode) Or _
                            Trim(HIN_GAI) <> Trim(StrConv(tmpZAIKOREC.HIN_GAI, vbUnicode)) Or _
                            LOCATION <> (StrConv(tmpZAIKOREC.Soko_No, vbUnicode) & _
                                            StrConv(tmpZAIKOREC.Retu, vbUnicode) & _
                                            StrConv(tmpZAIKOREC.Ren, vbUnicode) & _
                                            StrConv(tmpZAIKOREC.Dan, vbUnicode)) Then
                            sts = BtErrEOF
                        
                        End If
                    
                    
                        If Zan_Qty < CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode)) Then
                            Upd_com = BtOpUpdate
                            WK_Qty = Zan_Qty
                        Else
                            Upd_com = BtOpDelete
                            WK_Qty = CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        End If

                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j", 0)
                                GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                Exit Function

                            End If

                        End If

                        DoEvents
                    
                    Case Else
                        Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                        GOODS_ONOFF_Update_Proc = SYS_ERR
                        Exit Function
                End Select

            Loop

            If sts = BtErrEOF Then
                Exit Do
            End If

            If Upd_com = BtOpUpdate Then
                                                                            '�L���݌ɐ�
                Call UniCode_Conv(tmpZAIKOREC.YUKO_Z_QTY, Format(CLng(StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode)) - WK_Qty, "00000000"))
            
            End If


            RETRY_CNT = 0
'------- �݌Ɂi�ꎞ�f�[�^�j��������
            Do
                sts = BTRV(Upd_com, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j", 0)
                                GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                Exit Function

                            End If

                        End If
                        DoEvents
                    Case Else
                        Call File_Error(sts, Upd_com, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                        GOODS_ONOFF_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            
            Loop
'============================================================ ���ۂ̍݌Ƀf�[�^�Ɉړ�
                                                '��ɐV�K�ǉ�
            Call UniCode_Conv(ZAIKOREC.Soko_No, Left(LOCATION, 2))          '�q�ɇ�
            Call UniCode_Conv(ZAIKOREC.Retu, Mid(LOCATION, 3, 2))           '��
            Call UniCode_Conv(ZAIKOREC.Ren, Mid(LOCATION, 5, 2))            '�A
            Call UniCode_Conv(ZAIKOREC.Dan, Right(LOCATION, 2))             '�i
            Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                      '���ƕ�
            Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                      '���O
            Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
            Call UniCode_Conv(ZAIKOREC.GOODS_ON, "0")                       '���i�^�����i
                                                                            '���ד�
            Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(tmpZAIKOREC.NYUKA_DT, vbUnicode))
                                                                            '���ɓ�
            Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(tmpZAIKOREC.NYUKO_DT, vbUnicode))
                                                                            '�i�ԁi�����j
'            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(tmpZAIKOREC.NYUKO_DT, vbUnicode))  2005.09.03
            Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(tmpZAIKOREC.HIN_NAI, vbUnicode))
                                                                            '�L���݌ɐ�
            Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, Format(WK_Qty, "00000000"))
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                   '�r���t���O
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                         '�g�p���q�@ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                         '�g�p����۸���
            Call UniCode_Conv(ZAIKOREC.GOODS_YMD, Format(Now, "YYYYMMDD"))  '���i�����t


            Call UniCode_Conv(ZAIKOREC.FILLER, "")

            RETRY_CNT = 0
'*------------------------------------------------------'�݌Ƀf�[�^�o��
            Do
                sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                    '���g���C�񐔃`�F�b�N
                        If RETRY_SU <> 0 Then

                            RETRY_CNT = RETRY_CNT + 1
                            If RETRY_CNT > RETRY_SU Then
                                    '�񐔃I�[�o�[
                                Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                                GOODS_ONOFF_Update_Proc = SYS_CANCEL
                                Exit Function

                            End If

                        End If
                        DoEvents
                    Case Else
                        Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^")
                        GOODS_ONOFF_Update_Proc = SYS_ERR
                        Exit Function

                End Select
            Loop

            Zan_Qty = Zan_Qty - WK_Qty

            If Zan_Qty <= 0 Then
                Exit Do                     '�������Ƃ��I���i���i���ςݕ��j
            End If

        Loop
                
    End If
'================================================================================
    '*
    '*--------------------  �����i���̏���(�ꎞ�݌ɂɎc���Ă��镪�͑S�Ė����i�Ƃ��Čv��)
    
    Call UniCode_Conv(K0_tmpZAIKO.Soko_No, Left(LOCATION, 2))
    Call UniCode_Conv(K0_tmpZAIKO.Retu, Mid(LOCATION, 3, 2))
    Call UniCode_Conv(K0_tmpZAIKO.Ren, Mid(LOCATION, 5, 2))
    Call UniCode_Conv(K0_tmpZAIKO.Dan, Right(LOCATION, 2))
    Call UniCode_Conv(K0_tmpZAIKO.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_tmpZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_tmpZAIKO.HIN_GAI, HIN_GAI)
    Call UniCode_Conv(K0_tmpZAIKO.NYUKA_DT, "")
    
    
    com = BtOpGetGreaterEqual
    
    Do


        DoEvents


        RETRY_CNT = 0
        Do
            sts = BTRV(com + BtSNoWait, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
            Select Case sts
                Case BtNoErr
                                        '�I�{�i�u���[�N
                    If LOCATION <> (StrConv(tmpZAIKOREC.Soko_No, vbUnicode) & _
                                        StrConv(tmpZAIKOREC.Retu, vbUnicode) & _
                                        StrConv(tmpZAIKOREC.Ren, vbUnicode) & _
                                        StrConv(tmpZAIKOREC.Dan, vbUnicode)) Or _
                        JGYOBU <> StrConv(tmpZAIKOREC.JGYOBU, vbUnicode) Or _
                        NAIGAI <> StrConv(tmpZAIKOREC.NAIGAI, vbUnicode) Or _
                        Trim(HIN_GAI) <> Trim(StrConv(tmpZAIKOREC.HIN_GAI, vbUnicode)) Then


                        sts = BtErrEOF

                    End If
                    Exit Do
                Case BtErrEOF
                    
                    Exit Do

                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    DoEvents
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function
            End Select

        Loop


        If sts = BtErrEOF Then
            Exit Do
        End If

'============================================================ ���ۂ̍݌Ƀf�[�^�Ɉړ�
                                                '��ɐV�K�ǉ�
        Call UniCode_Conv(ZAIKOREC.Soko_No, Left(LOCATION, 2))          '�q�ɇ�
        Call UniCode_Conv(ZAIKOREC.Retu, Mid(LOCATION, 3, 2))           '��
        Call UniCode_Conv(ZAIKOREC.Ren, Mid(LOCATION, 5, 2))            '�A
        Call UniCode_Conv(ZAIKOREC.Dan, Right(LOCATION, 2))          '�i
        Call UniCode_Conv(ZAIKOREC.JGYOBU, JGYOBU)                      '���ƕ�
        Call UniCode_Conv(ZAIKOREC.NAIGAI, NAIGAI)                      '���O
        Call UniCode_Conv(ZAIKOREC.HIN_GAI, HIN_GAI)                    '�i�ԁi�O���j
        Call UniCode_Conv(ZAIKOREC.GOODS_ON, "1")                       '���i�^�����i
                                                                        '���ד�
        Call UniCode_Conv(ZAIKOREC.NYUKA_DT, StrConv(tmpZAIKOREC.NYUKA_DT, vbUnicode))
                                                                        '���ɓ�
        Call UniCode_Conv(ZAIKOREC.NYUKO_DT, StrConv(tmpZAIKOREC.NYUKO_DT, vbUnicode))
                                                                        '�i�ԁi�����j
        Call UniCode_Conv(ZAIKOREC.HIN_NAI, StrConv(tmpZAIKOREC.HIN_NAI, vbUnicode))
                                                                        '�L���݌ɐ�
        Call UniCode_Conv(ZAIKOREC.YUKO_Z_QTY, StrConv(tmpZAIKOREC.YUKO_Z_QTY, vbUnicode))
        Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)                   '�r���t���O
        Call UniCode_Conv(ZAIKOREC.WEL_ID, "")                         '�g�p���q�@ID
        Call UniCode_Conv(ZAIKOREC.PRG_ID, "")                         '�g�p����۸���
        Call UniCode_Conv(ZAIKOREC.GOODS_YMD, Format(Now, "YYYYMMDD"))  '���i�����t


        Call UniCode_Conv(ZAIKOREC.FILLER, "")

        RETRY_CNT = 0
'*------------------------------------------------------'�݌Ƀf�[�^�o��
        Do
            sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, Upd_com, "�݌Ƀf�[�^", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If
                    DoEvents
                Case Else
                    Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        Loop

'*------------------------------------------------------'�݌Ɂi�ꎞ�f�[�^�j�폜
        Do
            sts = BTRV(BtOpDelete, tmpZAIKO_POS, tmpZAIKOREC, Len(tmpZAIKOREC), K0_tmpZAIKO, Len(K0_tmpZAIKO), 0)
        
        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                '���g���C�񐔃`�F�b�N
                    If RETRY_SU <> 0 Then

                        RETRY_CNT = RETRY_CNT + 1
                        If RETRY_CNT > RETRY_SU Then
                                '�񐔃I�[�o�[
                            Call File_Error(sts, BtOpDelete, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j", 0)
                            GOODS_ONOFF_Update_Proc = SYS_CANCEL
                            Exit Function

                        End If

                    End If

                    DoEvents
                Case Else
                    Call File_Error(sts, BtOpDelete, "�݌Ƀf�[�^�i�ꎞ�f�[�^�j")
                    GOODS_ONOFF_Update_Proc = SYS_ERR
                    Exit Function

            End Select
        
        Loop


        com = BtOpGetNext
    
    Loop
                
    
    
    GOODS_ONOFF_Update_Proc = False

End Function

Private Function RETURNED_GOODS_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�Ǖi�ԕi�x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim Tanaban         As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2

    RETURNED_GOODS_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                RETURNED_GOODS_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
        
            SUMI_QTY = 0
            MI_QTY = 0
        
        
            '-----------------------------------------------���M�e�L�X�g�쐬
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban                    '�i�Ԃ��Z�[�u
                                                        
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O2006.01.06
                                                        
                                                        
                                                        
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, Left(ID_KANRI_TBL(ING_No).YOIN_DNAME, 2) & "[" & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode) & "]")
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, Left(ID_KANRI_TBL(ING_No).YOIN_DNAME, 2) & "[" & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode) & "]")
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
            Call UniCode_Conv(Send_Text.Box_Type(1).LCD, Hinban)

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, Hinban)
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
            Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Box_Type = TYPE_BCANK
                                                                    '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(2).LCD, LCD_Tanaban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).LCD, LCD_Tanaban)
                                                                    '���l�����\��
            Send_Text.Box_Type(2).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).INIT = ""
                                                                    '�����J�[�\���ʒu
            Send_Text.Box_Type(2).Start_Pos = "01"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = "01"
                                                                    '���͌���
            Send_Text.Box_Type(2).Max_Size = "09"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "09"
                                                                    
            Send_Text.Box_Type(2).MENU = ""                         '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#")))) & Format(SUMI_QTY, "#"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_SUMI_Suryo & Space(M_Keta - (Len(LCD_SUMI_Suryo) * 2) - 5 + (5 - Len(Format(SUMI_QTY, "#")))) & Format(SUMI_QTY, "#"))
                                                            '���l�����\��
            Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#"))) & Format(SUMI_QTY, "#")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).INIT = Space(10 - Len(Format(SUMI_QTY, "#"))) & Format(SUMI_QTY, "#")
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
            Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCNUM
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#")))) & Format(MI_QTY, "#"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_MI_Suryo & Space(M_Keta - (Len(LCD_MI_Suryo) * 2) - 5 + (5 - Len(Format(MI_QTY, "#")))) & Format(MI_QTY, "#"))
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#"))) & Format(MI_QTY, "#")
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = Space(10 - Len(Format(MI_QTY, "#"))) & Format(MI_QTY, "#")
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")      '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = Format(M_Keta - 4, "00")
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "05"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "05"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i���i�^�����i���ʁj
            
            QTY = 0
            SUMI_QTY = 0
            MI_QTY = 0
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
                    Case LCD_Tanaban        '�I��
                        Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        
                        If Trim(Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                        '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                    RETURNED_GOODS_Proc = False
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
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    RETURNED_GOODS_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                        
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�ԃG���[", "", "")
                            
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                        
                                    RETURNED_GOODS_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
                    
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                                    
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(Tanaban, 2) & "-" & Mid(Tanaban, 3, 2) & "-" & Mid(Tanaban, 5, 2) & "-" & Right(Tanaban, 2), "�I�g�p�s��", "", "")
                        
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                
                                RETURNED_GOODS_Proc = False
                                Exit Function
                            End If
                    
                    
                        End If
            
            
            
                    Case LCD_Suryo          '���ʁi�����͖����j
                        If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            RETURNED_GOODS_Proc = False
                            Exit Function
                        End If
                
                        QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        If QTY = 0 Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            RETURNED_GOODS_Proc = False
                            Exit Function
                        End If
            
            
                    Case LCD_SUMI_Suryo, LCD_MI_Suryo    '���ʁi���i���ςݐ��ʁ^�����i���ʁj
                
                       If Not IsNumeric(Trim(ID_KANRI_TBL(ING_No).Recv_text(i))) Then
                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
                            Sendbuf = Text_Create_Proc()
                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                            RETURNED_GOODS_Proc = False
                            Exit Function
                        End If
                
                
                        If Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size))) = LCD_SUMI_Suryo Then
                            SUMI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))

'                            If SUMI_QTY > ID_KANRI_TBL(ING_No).Send_SUMI_QTY Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                RETURNED_GOODS_Proc = False
'                                Exit Function
'
'                            End If
                        Else
                            MI_QTY = CLng(Trim(ID_KANRI_TBL(ING_No).Recv_text(i)))
                        
'                            If MI_QTY > ID_KANRI_TBL(ING_No).Send_MI_QTY Then
'                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Trim(ID_KANRI_TBL(ING_No).Recv_text(i)), "���ʓ��̓~�X", "", "")
'                                Sendbuf = Text_Create_Proc()
'                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
'
'                                RETURNED_GOODS_Proc = False
'                                Exit Function
'                            End If
                        End If
        
                        If i = M_Gyo - 1 Then       '�ŏI�s��������
                            If SUMI_QTY = 0 And MI_QTY = 0 Then
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "���i�^�����i���O", "���ʓ��̓~�X", "", "")
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                                RETURNED_GOODS_Proc = False
                                Exit Function
                            End If
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
                                    Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , , , , MENU_NO)
            Select Case sts
                Case False
                Case True           '���Ɏ��͔������Ȃ�
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    
                    RETURNED_GOODS_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Sendbuf = Text_Create_Proc()
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    RETURNED_GOODS_Proc = SYS_ERR    '�V�X�e���ُ픭��
                    
                    GoTo Abort_Tran
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
    
    End Select

    RETURNED_GOODS_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
Private Function Location_Move_Proc(Sendbuf As String, Ti As Integer, Tj As Integer) As Integer
'-------------------------------------------------------
'
'   �w�I�ړ��w�莞�̃`�F�b�N���X�V�����x
'
'-------------------------------------------------------
Dim sts             As Integer

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim QTY             As Long


Dim i               As Integer

Dim From_Tanaban    As String * 8
Dim To_Tanaban      As String * 8
Dim Hinban          As String * 13

Dim RET_JGYOBU      As String * 1
Dim RET_NAIGAI      As String * 1

Dim MENU_NO         As String * 2
    
    Location_Move_Proc = True

    Select Case ID_KANRI_TBL(ING_No).Step
        Case Step_Sagyo1_RES        '�P��ڂ̎�M�i�I�ԁ^�i�ԁj
            For i = 0 To M_Gyo - 1
                Select Case Trim(WEL_Para_Tbl(Ti, Tj).Wel_Para(i).LCD)
                    Case LCD_Tanaban        '�I��
                        From_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(From_Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Location_Move_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^", 0)
                                    Exit Function
                            End Select
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(From_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(From_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(From_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(From_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Location_Move_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Location_Move_Proc = False
                                Exit Function
                            End If
            
                        End If
                    Case LCD_Hinban         '�i��
                        Hinban = ID_KANRI_TBL(ING_No).Recv_text(i)
                
                        '------------------ �i�ڃ}�X�^�Ǎ���
'                        Call UniCode_Conv(K0_ITEM.JGYOBU, ID_KANRI_TBL(ING_No).JGYOBU)
'                        Call UniCode_Conv(K0_ITEM.NAIGAI, ID_KANRI_TBL(ING_No).NAIGAI)
'                        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
'                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        sts = Item_Read_Proc(ID_KANRI_TBL(ING_No).JGYOBU, ID_KANRI_TBL(ING_No).NAIGAI, Hinban, RET_JGYOBU, RET_NAIGAI)
                        Select Case sts
                            Case BtNoErr
                                If Trim(From_Tanaban) = Loc_OK_Para Then
                                    '�I��OK���̒I�ԃ`�F�b�N
                                    Call UniCode_Conv(K0_TANA.Soko_No, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Retu, StrConv(ITEMREC.ST_RETU, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Ren, StrConv(ITEMREC.ST_REN, vbUnicode))
                                    Call UniCode_Conv(K0_TANA.Dan, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                    sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            From_Tanaban = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                                        Case BtErrKeyNotFound
                                        '   -------------------------------- �G���[���b�Z�[�W�쐬
                                            Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                            Sendbuf = Text_Create_Proc()
                                            ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                            Location_Move_Proc = False
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
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Hinban, "�i�ԃG���[", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                Location_Move_Proc = False
                                Exit Function
                            Case Else
                                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                Sendbuf = Text_Create_Proc()
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^", 0)
                                Exit Function
                        End Select
        
                End Select
            Next i
        
            '------------------ �݌ɂ̎g�p�\����s���A�L���݌ɐ����l������
            sts = Zaiko_Reserve_Proc(ID_KANRI_TBL(ING_No).ID, From_Tanaban, RET_JGYOBU, RET_NAIGAI, Hinban, SUMI_QTY, MI_QTY)
            Select Case sts
                Case False
                Case True           '�����ł͔������Ȃ�
                    Exit Function
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Exit Function
                Case SYS_CANCEL
                    Call Err_Send_Proc("�݌Ɏg�p��", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Location_Move_Proc = False
                    Exit Function
            End Select
                    
            If (SUMI_QTY = 0) And (MI_QTY = 0) Then
                Call Err_Send_Proc("�L���݌ɖ���", Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2), Hinban, "", "")
                Sendbuf = Text_Create_Proc()
                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                Location_Move_Proc = False
                Exit Function
            End If
        
        
            ID_KANRI_TBL(ING_No).Step = Step_Sagyo2_REQ
            
            ID_KANRI_TBL(ING_No).Tanaban = From_Tanaban       '�I�Ԃ��Z�[�u
            
            ID_KANRI_TBL(ING_No).S_JGYOBU = RET_JGYOBU      '���ޑΉ��̎��ƕ�2006.01.06
            ID_KANRI_TBL(ING_No).S_NAIGAI = RET_NAIGAI      '���ޑΉ��̍����O2006.01.06
            
            
            ID_KANRI_TBL(ING_No).Hinban = Hinban            '�i�Ԃ��Z�[�u
            ID_KANRI_TBL(ING_No).Send_SUMI_QTY = SUMI_QTY   '���M���鏤�i���ςݐ���
            ID_KANRI_TBL(ING_No).Send_MI_QTY = MI_QTY       '���M���関���i����
            
            '���ʕt���̑��M���b�Z�[�W���쐬����
            Send_Text.sts = Sts_OK                                  '�X�e�[�^�X�@OK
            ID_KANRI_TBL(ING_No).Send_Text.sts = Sts_OK
    
            Send_Text.Display_Flg = Display_DEF                     '�\����ʃt���O �ʏ���͉��
            ID_KANRI_TBL(ING_No).Send_Text.Display_Flg = Display_DEF
    
            Send_Text.End_Menu = Menu_Only                          '�ŏI���j���[�t���O
            ID_KANRI_TBL(ING_No).Send_Text.End_Menu = Menu_Only
    
            Send_Text.Menu_Suu = "05"                               '���j���[���ڐ��i05�Œ�j
            ID_KANRI_TBL(ING_No).Send_Text.Menu_Suu = "05"
    
            Send_Text.fileName = ""                                 '���M�f�[�^�t�@�C����
            ID_KANRI_TBL(ING_No).Send_Text.fileName = ""
    
            Send_Text.Buzzer = Buzzer_DEF                           '�u�U�[���@�W��
            ID_KANRI_TBL(ING_No).Send_Text.Buzzer = Buzzer_DEF
                                                                        
            '-----------------------------------------------�P�s��
                                                            'BOX����
            Send_Text.Box_Type(0).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(0).LCD, ID_KANRI_TBL(ING_No).YOIN_DNAME)
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
                            Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))

            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).LCD, _
                            Left(From_Tanaban, 2) & "-" & Mid(From_Tanaban, 3, 2) & "-" & Mid(From_Tanaban, 5, 2) & "-" & Right(From_Tanaban, 2))
                                                            '���l�����\��
            Send_Text.Box_Type(1).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).INIT = ""
                                                            
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(1).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Start_Pos = ""
                                                            '���͌���
            Send_Text.Box_Type(1).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(1).Max_Size = "00"
                                                                                
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
            Send_Text.Box_Type(2).Start_Pos = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Start_Pos = ""
                                                            '���͌���
            Send_Text.Box_Type(2).Max_Size = "00"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).Max_Size = "00"
                                                                                
            Send_Text.Box_Type(2).MENU = ""                     '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(2).MENU = ""
            '-----------------------------------------------�S�s��
                                                            'BOX����
            Send_Text.Box_Type(3).Box_Type = TYPE_REF
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).Box_Type = TYPE_REF
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(3).LCD, LCD_Suryo & ":" & Format(SUMI_QTY + MI_QTY, "#0"))
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(3).LCD, LCD_Suryo & ":" & Format(SUMI_QTY + MI_QTY, "#0"))
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
            Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Box_Type = TYPE_BCANK
                                                            '�\�����e
            Call UniCode_Conv(Send_Text.Box_Type(4).LCD, LCD_To_Tanaban)
            Call UniCode_Conv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).LCD, LCD_To_Tanaban)
                                                            '���l�����\��
            Send_Text.Box_Type(4).INIT = ""
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).INIT = ""
                                                            '�����J�[�\���ʒu
            Send_Text.Box_Type(4).Start_Pos = "01"          '���l�͂T���Œ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Start_Pos = "01"
                                                            '���͌���
            Send_Text.Box_Type(4).Max_Size = "09"
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).Max_Size = "09"
                                                                                
            Send_Text.Box_Type(4).MENU = ""                 '���j���\�ԍ�
            ID_KANRI_TBL(ING_No).Send_Text.Box_Type(4).MENU = ""
        
        
            Sendbuf = Text_Create_Proc()
        
        
        Case Step_Sagyo2_RES        '�Q��ڂ̎�M�i�ړ���I�ԁj
            
            
            
            For i = 0 To M_Gyo - 1
            
                
                Select Case Trim(Left(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode), Len(StrConv(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).LCD, vbUnicode)) - CInt(ID_KANRI_TBL(ING_No).Send_Text.Box_Type(i).Max_Size)))
            
            
                    Case LCD_To_Tanaban         '�ړ���I��
                    
                    
                        To_Tanaban = Right(ID_KANRI_TBL(ING_No).Recv_text(i), Len(ID_KANRI_TBL(ING_No).Recv_text(i)) - 1)
                        If Trim(From_Tanaban) = Loc_OK_Para Then '�I��OK
                        Else
                            '------------------ �q�Ƀ}�X�^�Ǎ���
                            Call UniCode_Conv(K0_SOKO.Soko_No, Left(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2), "�q�ɃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Location_Move_Proc = False
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
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Recv_text(i), 2), "���ڃG���[", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                                    Location_Move_Proc = False
                                    Exit Function
                                End If
                            End If
                            '------------------ �I�}�X�^�Ǎ���
                            Call UniCode_Conv(K0_TANA.Soko_No, Left(To_Tanaban, 2))
                            Call UniCode_Conv(K0_TANA.Retu, Mid(To_Tanaban, 3, 2))
                            Call UniCode_Conv(K0_TANA.Ren, Mid(To_Tanaban, 5, 2))
                            Call UniCode_Conv(K0_TANA.Dan, Right(To_Tanaban, 2))
                            sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                
                                '   -------------------------------- �G���[���b�Z�[�W�쐬
                                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�ԃG���[", "", "")
                    
                                    Sendbuf = Text_Create_Proc()
                                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                
                                    Location_Move_Proc = False
                                    Exit Function
                                Case Else
                                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                                    Sendbuf = Text_Create_Proc()
                                    Call File_Error(sts, BtOpGetEqual, "�I�}�X�^", 0)
                                    Exit Function
                            End Select
            
            
                            '------------------ �֎~�I�̃`�F�b�N
                            If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_NG Then
                            
                                Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(To_Tanaban, 2) & "-" & Mid(To_Tanaban, 3, 2) & "-" & Mid(To_Tanaban, 5, 2) & "-" & Right(To_Tanaban, 2), "�I�g�p�s��", "", "")
                    
                                Sendbuf = Text_Create_Proc()
                                ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                                Location_Move_Proc = False
                                Exit Function
                            End If
            
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
        
            If ID_KANRI_TBL(ING_No).SAGYO_LOG = "1" Then
                MENU_NO = ID_KANRI_TBL(ING_No).MENU_LV1
                        
            Else
                        
                MENU_NO = ""
            End If
        
        
            sts = IDO_Update_Proc(ID_KANRI_TBL(ING_No).S_JGYOBU, _
                                    ID_KANRI_TBL(ING_No).S_NAIGAI, _
                                    ID_KANRI_TBL(ING_No).Hinban, _
                                    "", _
                                    ID_KANRI_TBL(ING_No).Tanaban, _
                                    To_Tanaban, _
                                    (ID_KANRI_TBL(ING_No).Sagyo_Code.CODE_TYPE & ID_KANRI_TBL(ING_No).Sagyo_Code.YOIN_CODE), _
                                    ID_KANRI_TBL(ING_No).Send_SUMI_QTY, _
                                    ID_KANRI_TBL(ING_No).Send_MI_QTY, _
                                    Format(ID_KANRI_TBL(ING_No).ID, "000"), _
                                    ID_KANRI_TBL(ING_No).TANTO_CODE, _
                                    FILE_RETRY, , MENU_NO)

    
    
            Select Case sts
                Case False
        
                Case True       '�݌ɕs�����ɔ���
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, Left(ID_KANRI_TBL(ING_No).Tanaban, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 3, 2) & "-" & Mid(ID_KANRI_TBL(ING_No).Tanaban, 5, 2) & "-" & Right(ID_KANRI_TBL(ING_No).Tanaban, 2), ID_KANRI_TBL(ING_No).Hinban, "�݌ɐ��s��", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
            
                    Location_Move_Proc = False
                    GoTo Abort_Tran
                Case SYS_CANCEL
                    Call Err_Send_Proc(ID_KANRI_TBL(ING_No).YOIN_DNAME, "�������f", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    ID_KANRI_TBL(ING_No).Step = ID_KANRI_TBL(ING_No).Step - 1
                    Location_Move_Proc = False
                    GoTo Abort_Tran
                Case SYS_ERR
                    Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                    Sendbuf = Text_Create_Proc()
                    Location_Move_Proc = SYS_ERR    '�V�X�e���ُ픭��
                    GoTo Abort_Tran
            End Select
    
                                        '�g�����U�N�V�����I��
            sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            If sts <> BtNoErr Then
                Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
                Sendbuf = Text_Create_Proc()
                Call File_Error(sts, BtOpEndTransaction, "", 0)
                GoTo Abort_Tran
            End If
                                        
                                        
                            
            '�o�ח\��^�݌ɂ̗\�����
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

    Location_Move_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call Err_Send_Proc("�V�X�e���ُ픭��", "", "", "", "")
        Sendbuf = Text_Create_Proc()
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function


Private Function Dec_To_Bcd(DecStr As String) As Variant
Dim i           As Long
Dim BCDChr      As Variant

    Dec_To_Bcd = ""

    For i = 1 To Len(DecStr) Step 2
        BCDChr = Chr(Val(Mid(DecStr, i, 1)) * 16 Or Val(Mid(DecStr, i + 1, 1)))
        Dec_To_Bcd = Dec_To_Bcd & BCDChr
    Next i

End Function

Private Function Bcd_To_Dec(BcdStr As String) As Variant
Dim i           As Long
Dim DecLow      As Long

    Bcd_To_Dec = ""

    For i = 1 To Len(BcdStr)
        DecLow = Asc(Mid(BcdStr, i, 1)) Mod 16
        Bcd_To_Dec = Bcd_To_Dec & CStr((Asc(Mid(BcdStr, i, 1)) - DecLow) / 16) & CStr(DecLow)
    Next i

End Function

