VERSION 5.00
Begin VB.Form F1020151 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���o�ח\��f�[�^�捞�� (F102015 2016.03.08 09�F30) "
   ClientHeight    =   4170
   ClientLeft      =   1905
   ClientTop       =   2385
   ClientWidth     =   8580
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   8580
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ListBox LBox_Hin 
      Height          =   300
      Left            =   1560
      TabIndex        =   25
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6360
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6360
      TabIndex        =   23
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   22
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   21
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblINCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   20
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   5760
      TabIndex        =   19
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   5760
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   5760
      TabIndex        =   17
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblOUTCNT 
      Alignment       =   1  '�E����
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblJGYOBU 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "F1020151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WS_NO       As String * 3           'ܰ��ð��ݔԍ�

Private FileName    As String               '�e�L�X�g�t�@�C����
Private FileNo      As Integer              '�t�@�C����

Private KASO_NYUKA_SOKO      As String * 2  '���z���בq�ɔԍ�
Private KASO_SMODOSHI_SOKO   As String * 2  '���z�x���߂��q�ɔԍ�

Private Proc_F      As Integer              '�i�ԁ��݌ɗL���@����t���O
Private Last_Proc_F As Integer              '���������ް��폜�����@���s�L���t���O
                                            
Private Type YUKO_SOKO_TBL                  '�L��νđq�Ɏ�荞�݃e�[�u��
    HS_SOKO             As String * 8
    NAIGAI              As String * 1
End Type

Dim Soko_T()            As YUKO_SOKO_TBL  '�q�ɏ��

'-                                          2005.12.30
Private Type SHIMUKE_TBL
    SHIMUKE_CODE            As String * 2   '�d������
    JGYOBU                  As String * 1   '���ƕ�
    NAIGAI                  As String * 1   '�����O
End Type

Private SHIMUKE_T()         As SHIMUKE_TBL

Private SHIMUKE_Flg         As Boolean

'-                                          2005.12.30


Private New_HS_IN_SIJ   As String           '���Ƀf�[�^�t�@�C����
Private New_HS_OUT_SIJ  As String           '�Vڲ��ďo�Ƀf�[�^�t�@�C����


Private In_Cnt      As Integer              '�f�[�^�ǂݍ��݌���
Private Out_Cnt     As Integer              '�f�[�^�o�͌���

Private Const In_Mode% = 1                  '���׏���
Private Const Out_Mode% = 2                 '�o�׏���

                                            
Dim NormalFont As New StdFont               '����t�H���g

Private Const LMAX% = 46                    '�œ��ő�s��
Private Const MGN_L% = 1                    '���׈���J�n���ʒu�i�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j
Private Pdate As String                     '����J�n���t�iͯ�ް�p�j
Private Ptime As String                     '����J�n�����iͯ�ް�p�j


Private Const NAI_CHANGE% = 1
Private Const GAI_CHANGE% = 2
Private Const NOT_GAI_CHANGE% = 3


Private ETC_MTS_NAI As String * 8             '���̑�������(����)
Private ETC_SS_NAI  As String * 8             '���̑�������(����)

Private ETC_MTS_GAI As String * 8             '���̑�������(�C�O)
Private ETC_SS_GAI  As String * 8             '���̑�������(�C�O)

Dim DUP_SYUKA_DATA  As String                 '�o�׃f�[�^�t���p�X

                                        
Dim MyCenter        As String

Dim Err_FLg         As Boolean

Dim TANA_SPACE      As Boolean          '2009.03.07
                                        
Private MENU_NO     As String * 2       '���у��O�o�͗p�ƭ���   2007.11.06
                                        
                                        
                                        
Dim RYOHEN_TANA     As String * 8       '�Ǖi�ԕi���ɒI��       2011.01.18
                                        
'���i���v��x�� 2011.07.07
Dim NOT_Hin_Name    As Variant          '���O�i��
Dim NOT_Hin_Name_F  As Boolean          '���O�i���L��
'���i���v��x�� 2011.07.07
Dim GOODS_F         As String * 1       '���i���L���@��̫�� 2012.12.20
                                        
                                        
Private Function New_Nyuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   �u���ח\��f�[�^�v�X�V����
'----------------------------------------------------------------------------

''''''''''''''''''''''''''''''' �S�Ĕp�~ 2009.06.18
'''Dim i           As Integer
'''Dim j           As Integer
'''Dim Skip_Flg    As Boolean
    
'''Dim WK_Y_QTY    As Long     '�o�א����[�N
'''Dim WK_Qty      As Long     '�O�؎c���[�N
'''Dim WK_E_QTY    As Long     '��s�o�א����[�N

'''Dim SUMI_QTY    As Long     '���i���ς݂Ƃ��ēo�^
'''Dim MI_QTY      As Long     '�����i�Ƃ��ēo�^

'''Dim Work_SOKO     As String * 2
    
'''Dim sts         As Integer
'''Dim ans         As Integer
'''Dim Not_SHUSI   As Boolean
    
''''�o�ח\�� �ҏW�O���� ################################################################# 2005/05/16 Add ��
'''Dim Fast_Flg        As Boolean
'''Dim DUP_SYUKANo     As Integer
'''Dim fileName        As String
'''Dim Ret             As Integer
'''Dim INS_NOW         As String * 14
'''Dim wkStr           As String
    
'''Dim wkMUKE_CODE     As String
    
    
'''Dim NAIGAI          As String * 1
    
'''''''''''''''''''''''''''''   �S�Ă̓��׏�����p�~    2007.06.22
'''    Fast_Flg = True
'''
'''
'''
'''    DUP_SYUKANo = FreeFile
'''    fileName = DUP_SYUKA_DATA
'''
'''    Ret = InStr(1, Trim(fileName), ".") - 1
'''    fileName = Left(Trim(fileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
'''
'''    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
''''#################################################################################### 2005/05/16 Add ��
'''
'''    New_Nyuka_Update_Proc = True
'''
'''
'''    Do
'''        Get #FileNo, , New_HS_IN_SIJREC
'''        If Left(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), 1) < " " Then
'''            Exit Do
'''        End If
'''
'''        If StrConv(New_HS_IN_SIJREC.CR_LF, vbUnicode) <> vbCrLf Then
'''
'''            Call NG_File_Make_Proc
'''
'''
'''            Exit Do
'''        End If
'''
'''
'''        In_Cnt = In_Cnt + 1
'''        lblINCNT(i).Caption = Format(In_Cnt, "#0")
'''        DoEvents
'''
'''
''''        Skip_Flg = True
''''        For i = 0 To UBound(JGYOBU_T)               '���x�敪�̃`�F�b�N
''''            If JGYOBU = JGYOBU_T(i).CODE Then
''''                For j = 0 To UBound(Soko_T, 2)
''''                    If StrConv(HS_IN_SIJREC.HOST_SOKO, vbUnicode) = Soko_T(i, j).HS_SOKO Then
''''                        Skip_Flg = False
''''                        Exit For
''''                    End If
''''                Next j
''''                Exit For
''''            End If
''''        Next i
'''
''''--     2005.12.30
'''        Skip_Flg = True
'''        Not_SHUSI = False
'''        For i = 0 To UBound(JGYOBU_T)               '���x�敪�̃`�F�b�N
'''            If JGYOBU = JGYOBU_T(i).CODE Then
'''                For j = 0 To UBound(Soko_T, 2)
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = Soko_T(i, j).HS_SOKO Then
'''                        Skip_Flg = False
'''                        Exit For
'''                    End If
'''                Next j
'''                Exit For
'''            End If
'''        Next i
'''
'''        If Skip_Flg Then
'''            Not_SHUSI = True
'''        End If
''''--     2005.12.30
'''
'''
'''
'''
'''
'''        If StrConv(New_HS_IN_SIJREC.IO_KBN, vbUnicode) <> "1" Then
'''            Skip_Flg = True
'''        End If
'''
'''
'''        If StrConv(New_HS_IN_SIJREC.PM_KBN, vbUnicode) = "-" Then
'''            Skip_Flg = True
'''        End If
'''
'''        'NOPOS  2006.05.01
'''        If Trim(StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) = "NOPOS" Then
'''            Skip_Flg = True
'''        End If
'''
'''
'''
'''
'''
'''
'''        Work_SOKO = KASO_NYUKA_SOKO
'''
'''
'''
'''        Select Case JGYOBU
'''
'''''            Case SENTAKU                        '����@
'''''
'''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
'''''                    Skip_Flg = True
'''''                End If
'''''
'''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
'''''                    Skip_Flg = True
'''''                End If
'''
'''
'''
'''
'''            Case SOJIKI                         '�|���@
'''
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KM" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "GG" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "SS" Then
'''                    Skip_Flg = True
'''                End If
'''
'''                '2005.04.07 ���x�ǉ�
'''                If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 5) = "0090K" Then
'''                    Skip_Flg = True
'''                End If
'''
'''
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KM" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KK" Then
''''                    Skip_Flg = True
''''                End If
'''
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KM" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "GG" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KK" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "GG" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "KK" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "SS" Then
''''                    Skip_Flg = True
''''                End If
'''
''''                If Left(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "SS" And _
''''                    Left(StrConv(HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "KK" Then
''''                    Skip_Flg = True
''''                End If
'''
'''
'''                If Trim(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
'''                    Work_SOKO = KASO_SMODOSHI_SOKO
'''                End If
'''
'''
'''
'''            Case DENKA, SUIHAN, SENTAKU         '�d���A���сA����@�i�A�C�����j
'''
'''
'''                Select Case MyCenter
'''
'''                    Case "O"
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "01" Then
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "H33" Then    '2004.07.16
'''                            Skip_Flg = True
'''                        End If
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "H22" Then    '2004.07.16
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "05" Then
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "G22" Then
'''                            Work_SOKO = "80"
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) = "G11" Then
'''                            Work_SOKO = "81"
'''                        End If
'''
'''                        '2006.04.29�p
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "S1" And _
'''                            Left(StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "S3" Then
'''                            Work_SOKO = "87"
'''                        End If
'''                        '2006.05.01
'''                        If Trim(StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) = "POS87" Then
'''                            Work_SOKO = "87"
'''                        End If
'''
'''
'''                    Case "F"
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "P2" Then
'''                            Skip_Flg = True
'''                        End If
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 3) <> "904" Then
'''                            If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 1) = "9" Then
'''                              Skip_Flg = True
'''                            End If
'''                        End If
'''
'''
'''                        If Left(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode), 2) = "S1" And _
'''                            Left(StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode), 2) = "S2" Then
'''                            Work_SOKO = "88"
'''                        End If
'''
'''                        '2006.05.01
'''                        If Trim(StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) = "POS88" Then
'''                            Work_SOKO = "88"
'''                        End If
'''
'''
'''                End Select
'''             Case AIRCON                     '�G�A�R��
'''
'''                If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "J4" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "JG" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "JW" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "JV" Or _
'''                    StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "HY" Then
'''                    Skip_Flg = True
'''                End If
'''
'''
'''                If Trim(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) = "SH" Then
'''                    Skip_Flg = True
'''                End If
'''
'''
'''                If Trim(StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) = "S1" Then
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "OS" Then
'''                      Skip_Flg = True
'''                    End If
'''                End If
'''
'''                If Not Skip_Flg Then
''''---------------    2005.06.14
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S8" Then
'''                        Work_SOKO = "80"
'''                    Else
'''                        If StrConv(New_HS_IN_SIJREC.CYU_KBN, vbUnicode) = "A" Then
'''                            If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "SH" Then
'''                                Work_SOKO = "92"
'''                            Else
'''                                Work_SOKO = "95"
'''                            End If
'''                        Else
'''                            If StrConv(New_HS_IN_SIJREC.CYU_KBN, vbUnicode) = "D" Then
'''                                Work_SOKO = "70"
'''                            Else
''''---------------    2005.06.14
'''                                If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "SH" Then
'''                                    Work_SOKO = "92"
'''                                Else
'''
''''                                    If StrConv(HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S8" Then
''''                                        Work_SOKO = "80"
''''                                    Else
'''                                        Work_SOKO = "95"
''''                                    End If
'''                                End If
'''                            End If
'''                        End If
'''                    End If
'''                End If
'''
'''
'''
'''        End Select
'''
'''
'''
'''
'''
'''
'''
'''        If Not Skip_Flg Then
'''
'''
'''
'''
'''                                        '���ח\��d���`�F�b�N
'''            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'''            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''            Select Case sts
'''                Case BtNoErr
'''                    Call Log_Out(LOG_F, "Y_NYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''                    Skip_Flg = True
'''                Case BtErrKeyNotFound
'''                Case Else
'''                    Call File_Error(sts, BtOpGetEqual, "���ח\��")
'''                    Exit Function
'''            End Select
'''
'''            If Not Skip_Flg Then
'''                                                '�g�����U�N�V�����J�n
'''                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                If sts <> BtNoErr Then
'''                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
'''                    Exit Function
'''                End If
'''                                            '�i�ڃ}�X�^�`�F�b�N
'''                If Item_Check_Proc(In_Mode, JGYOBU, Soko_T(i, j).NAIGAI, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode)) Then
'''                    GoTo Abort_Tran
'''                End If
'''
'''
'''                                            '���׃f�[�^�쐬
'''                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
'''                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
'''                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
'''                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
'''                Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''
'''
'''
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
'''                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.KAIKEI_JGYOBA, "")
'''                Call UniCode_Conv(Y_NYUREC.SHISAN_JGYOBA, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000"))
'''                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
'''                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.SHISAN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.HOJYO_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.TANKA, "")
'''                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
'''                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
'''                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
'''                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
'''                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
'''                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
'''                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
'''                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
'''                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
'''                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.S_SHISAN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.S_HOJYO_SYUSI, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
'''                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
'''                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.ZAN_SHISAN_SYUSI, "")
'''                Call UniCode_Conv(Y_NYUREC.ZAN_HOJYO_SYUSI, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
'''                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
'''                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
'''
'''                Call UniCode_Conv(Y_NYUREC.SS_CODE, "")
'''                Call UniCode_Conv(Y_NYUREC.KEPIN_KAIJYO, "")
'''
'''
'''                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
'''
'''
'''                Last_Proc_F = True              '���������ް��폜�����@���s�L��
'''
'''
'''                '���������ް��X�V
'''                Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU)
'''                Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
'''                Call UniCode_Conv(K0_J_NYU.HIN_GAI, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''
'''                WK_Y_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''
'''
'''                Do
'''                    sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
'''                    Select Case sts
'''                        Case BtNoErr
'''                            If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
'''                                WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
'''                                Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
'''
'''                                Do
'''
'''                                    sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpUpdate, "���������ް�")
'''                                            Exit Function
'''                                    End Select
'''
'''                                Loop
'''                                WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            Else
'''                                Do
'''                                    sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpDelete, "���������ް�")
'''                                            Exit Function
'''                                    End Select
'''                                Loop
'''                                WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
'''                            End If
'''
'''                            Exit Do
'''                        Case BtErrKeyNotFound
'''                            WK_E_QTY = 0
'''                            Exit Do
'''                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                            Beep
'''                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'''                            If ans = vbCancel Then
'''                                Exit Function
'''                            End If
'''                        Case Else
'''                            Call File_Error(sts, BtOpGetEqual, "���������ް�")
'''                            Exit Function
'''                    End Select
'''                Loop
'''                                    '��s���א��i���׎��ѐ��j
'''                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
'''
'''                                    '�\�Z�P�ʌ�
'''                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode))
'''                                    '�\�Z�P�ʐ�
'''                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'''                                    '�W���I��
'''                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
'''                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'''                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
'''                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
'''                Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'''
'''                Call UniCode_Conv(Y_NYUREC.FILLER, "")
'''
'''                Do
'''                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''                    Select Case sts
'''                        Case BtNoErr
'''                            Exit Do
'''                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                            Beep
'''                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'''                            If ans = vbCancel Then
'''                                Exit Function
'''                            End If
'''                        Case Else
'''                            Call File_Error(sts, BtOpInsert, "���ח\��")
'''                            Exit Function
'''                    End Select
'''                Loop
'''
''''------------ 2005.12.30
'''                Select Case JGYOBU
'''                    Case AIRCON, SENTAKU
'''                        Call UniCode_Conv(K0_SOKO.Soko_No, Work_SOKO)
'''                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'''                        Select Case sts
'''                            Case BtNoErr
'''                            Case Else
'''                                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
'''                                Exit Function
'''                        End Select
'''
'''                        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
'''
'''                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            MI_QTY = 0
'''                        Else
'''
'''                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'''                                MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                                SUMI_QTY = 0
'''                            Else
'''                                SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                                MI_QTY = 0
'''                            End If
'''                        End If
'''
''''------------ 2005.12.30
'''
'''                    Case Else
'''
'''
'''                        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'''                            MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            SUMI_QTY = 0
'''                        Else
'''                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'''                            MI_QTY = 0
'''                        End If
'''                End Select
'''
'''
''''                Wk_SOKO = KASO_NYUKA_SOKO
''''                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
''''                    Wk_SOKO = KASO_SMODOSHI_SOKO
''''
''''                End If
'''
'''                '���א��ō݌Ƀf�[�^�X�V�i�{�j
'''                If Nyuko_Update_Proc(JGYOBU, _
'''                                    Soko_T(i, j).NAIGAI, _
'''                                    StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode), _
'''                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
'''                                    (Work_SOKO & "01" & "01" & "01"), _
'''                                    YOIN_TU_NYUKA, _
'''                                    SUMI_QTY, MI_QTY, _
'''                                    WS_NO, WS_NO, , _
'''                                    StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode) & " �`��:" & StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode)) Then
'''                    Exit Function
'''
'''                End If
'''
'''                '�O�؂萔�ō݌Ƀf�[�^�X�V�i�|�j
'''                If WK_E_QTY <> 0 Then
'''                '�݌Ƀf�[�^LOCK
'''                    If Zaiko_Lock_Proc((Work_SOKO & "01" & "01" & "01"), _
'''                                        JGYOBU, _
'''                                        Soko_T(i, j).NAIGAI, _
'''                                        StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode), _
'''                                        WS_NO) Then
'''                        Exit Function
'''
'''                    End If
'''
'''                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'''                        MI_QTY = WK_E_QTY
'''                    Else
'''                        SUMI_QTY = WK_E_QTY
'''                    End If
'''
'''
'''                    If Syuko_Update_Proc(JGYOBU, _
'''                                        Soko_T(i, j).NAIGAI, _
'''                                        StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode), _
'''                                        StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode), _
'''                                        (Work_SOKO & "01" & "01" & "01"), _
'''                                        YOIN_MAE_SOUSAI, _
'''                                        SUMI_QTY, MI_QTY, 0, _
'''                                        WS_NO, WS_NO) Then
'''                        Exit Function
'''
'''                    End If
'''
'''
'''
'''
'''
'''
'''                End If
'''
'''
'''                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                If sts <> BtNoErr Then
'''                    GoTo Abort_Tran
'''                End If
'''
'''
'''                Out_Cnt = Out_Cnt + 1
'''                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
'''                DoEvents
'''
'''            End If
'''
'''
'''
'''
'''
''''�o�ח\��ϊ�################################################## 2005/05/16 Add ���ꕨ����
'''        Else
'''
'''            If JGYOBU = AIRCON Then
'''
'''                If Not_SHUSI Then
'''                Else
'''                    If StrConv(New_HS_IN_SIJREC.IO_KBN, vbUnicode) = "2" Then
'''
'''
'''                        wkMUKE_CODE = ""
'''
'''                        If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S8" Then
'''                            wkMUKE_CODE = "S8"
'''                        Else
'''                            If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "SH" Then
'''                            Else
'''                                Select Case Trim(StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'''
'''                                    Case "Z0014"
'''                                        wkMUKE_CODE = "LM"
'''                                    Case "B0070"
'''                                        If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = "S2" Then
'''                                            wkMUKE_CODE = "S2"
'''                                        Else
'''                                            wkMUKE_CODE = "AC"
'''                                        End If
'''                                    Case Else
'''                                        wkMUKE_CODE = "AC"
'''                                End Select
'''                            End If
'''                        End If
'''
'''
'''                        If wkMUKE_CODE = "" Then
'''                        Else
'''                            Skip_Flg = False
'''                                                        '���ח\��d���`�F�b�N
'''                            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU)
'''                            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''                            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''                            Select Case sts
'''                                Case BtNoErr
'''                                    Call Log_Out(LOG_F, "Y_NYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''                                    Skip_Flg = True
'''                                Case BtErrKeyNotFound
'''                                Case Else
'''                                    Call File_Error(sts, BtOpGetEqual, "���ח\��")
'''                                    Exit Function
'''                            End Select
'''
'''
'''
'''
'''
'''        ''                                    '�o�ח\��d���`�F�b�N
'''        ''                    Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
'''        ''                    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'''        ''
'''        ''                    sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
'''        ''                    Select Case sts
'''        ''                        Case BtNoErr
'''        ''                            Call Log_Out(LOG_F, "Y_SYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�`�[�h�c��" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'''        ''                            Skip_Flg = True
'''        ''
'''        ''                            If Fast_Flg Then
'''        ''                                Open (fileName) For Output As DUP_SYUKANo
'''        ''                                Write #DUP_SYUKANo, , , "�o�׏d�����X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS")
'''        ''                                Write #DUP_SYUKANo, "�o�ד�", "�`�[��", "�x���溰��", "�q��/�r�r����", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"
'''        ''                                Fast_Flg = False
'''        ''                            End If
'''        ''
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.SYUKA_YMD, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.DEN_NO, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.MUKE_NAME, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.CHU_KBN, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.CHU_KBN_NAME, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.HIN_NO, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.SURYO, vbUnicode),
'''        ''                            Write #DUP_SYUKANo, StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode)
'''        ''
'''        ''                        Case BtErrKeyNotFound
'''        ''                        Case Else
'''        ''                            Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
'''        ''                            Exit Function
'''        ''                    End Select
'''
'''                            If Not Skip_Flg Then
'''
'''                                                                '�g�����U�N�V�����J�n
'''                                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                                If sts <> BtNoErr Then
'''                                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
'''                                    Exit Function
'''                                End If
'''                                                                '�i�ڃ}�X�^�`�F�b�N
'''                                If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode)) Then
'''                                    GoTo Abort_Tran
'''                                End If
'''
'''        '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
'''                                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
'''                                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
'''                                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU)
'''                                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
'''                                Call UniCode_Conv(Y_NYUREC.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''
'''
'''
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
'''                                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.KAIKEI_JGYOBA, "")
'''                                Call UniCode_Conv(Y_NYUREC.SHISAN_JGYOBA, "")
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000"))
'''                                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
'''                                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.HOJYO_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.TANKA, "")
'''                                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
'''                                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
'''                                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
'''                                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
'''                                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
'''                                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
'''                                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
'''                                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
'''                                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
'''                                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.S_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.S_HOJYO_SYUSI, "")
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
'''                                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
'''                                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.ZAN_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_NYUREC.ZAN_HOJYO_SYUSI, "")
'''
'''
'''                                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
'''                                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
'''                                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
'''
'''                                Call UniCode_Conv(Y_NYUREC.SS_CODE, "")
'''                                Call UniCode_Conv(Y_NYUREC.KEPIN_KAIJYO, "")
'''                                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
'''                                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")
'''                                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
'''                                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
'''                                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
'''                                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
'''                                Call UniCode_Conv(Y_NYUREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_NYUREC.FILLER, "")
'''
'''                                Do
'''                                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpInsert, "���ח\��")
'''                                            Exit Function
'''                                    End Select
'''                                Loop
'''        '-------------------------------------------------------'���׃f�[�^�̂ݓo�^����i�Ď捞�ݎ��������̂��߁j
'''
'''                                Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
'''                                Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
'''                                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
'''                                Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
'''                                Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
'''                                Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
'''
'''                                If Den_No_Set_Proc(21, JGYOBU, wkStr) Then
'''                                    GoTo Abort_Tran
'''                                Else
'''                                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, wkStr)
'''                                End If
'''
'''                                Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
'''                                Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, wkMUKE_CODE)
'''                                Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
'''                                Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
'''                                Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(Y_SYUREC.KEY_ID_NO, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, "")
'''                                Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, "")
'''
'''                                Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'''                                wkStr = Format(Val(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000")
'''                                Call UniCode_Conv(Y_SYUREC.SURYO, wkStr)
'''                                Call UniCode_Conv(Y_SYUREC.MUKE_CODE, wkMUKE_CODE)
'''                                Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
'''
'''                                Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.TANKA, "")
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
'''                                Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
'''
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
'''                                Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
'''                                Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
'''
'''                                Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(New_HS_IN_SIJREC.SYUK_NAME, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
'''                                Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, "")
'''                                Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
'''                                Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
'''                                Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
'''                                Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(New_HS_IN_SIJREC.CYOK_KBN, vbUnicode))
'''
'''                                Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
'''                                Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
'''                                Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
'''                                Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, "")
'''                                Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
'''                                Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
'''                                Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
'''                                Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(New_HS_IN_SIJREC.HOST_TANA, vbUnicode))
'''                                Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
'''                                Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
'''                                Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
'''                                Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
'''
'''                                Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
'''                                Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
'''                                Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
'''                                Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")               '2006.07.20
'''
'''
'''                                Call UniCode_Conv(Y_SYUREC.FILLER, "")
'''
'''                                Do
'''                                    sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
'''                                    Select Case sts
'''                                        Case BtNoErr
'''                                            Exit Do
'''                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'''                                            Beep
'''                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'''                                            If ans = vbCancel Then
'''                                                Exit Function
'''                                            End If
'''                                        Case Else
'''                                            Call File_Error(sts, BtOpInsert, "�o�ח\��")
'''                                            Exit Function
'''                                    End Select
'''                                Loop
'''
'''                                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''                                If sts <> BtNoErr Then
'''                                    GoTo Abort_Tran
'''                                End If
'''
'''                                Out_Cnt = Out_Cnt + 1
'''                                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
'''                                DoEvents
'''
'''                                If SYUKA_LOG_ON Then
'''                                    Call SYUKA_LOG_OUT_PROC("INS", "���ׂ��琶��")
'''                                End If
'''
'''                                If Not Fast_Flg Then
'''                                    Close #DUP_SYUKANo
'''                                End If
'''                            End If
'''                        End If
'''                    End If
'''                End If
'''            End If
''''#################################################################################### 2005/05/16 Add ��
'''
'''
'''
'''
'''
'''        End If
'''
'''    Loop
'''    New_Nyuka_Update_Proc = False
'''
'''    Exit Function
'''
'''Abort_Tran:
'''
'''    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'''    If sts <> BtNoErr Then
'''        Call File_Error(sts, BtOpAbortTransaction, "")
'''    End If


    
''''''''''''''''''''''''''''''' �S�Ĕp�~    2009.06.18
'''    New_Nyuka_Update_Proc = True
'''
'''    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
'''
'''
'''
'''
'''    Do
'''        Get #FileNo, , New_HS_IN_SIJREC
'''        If Left(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), 1) < " " Then
'''            Exit Do
'''        End If
'''
'''        If StrConv(New_HS_IN_SIJREC.CR_LF, vbUnicode) <> vbCrLf Then
'''
'''            Call NG_File_Make_Proc
'''
'''
'''            Exit Do
'''        End If
'''
'''
'''        In_Cnt = In_Cnt + 1
'''        lblINCNT(i).Caption = Format(In_Cnt, "#0")
'''        DoEvents
'''
'''
'''        Skip_Flg = True
'''        Not_SHUSI = False
'''        For i = 0 To UBound(JGYOBU_T)               '���x�敪�̃`�F�b�N
'''            If StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode) = JGYOBU_T(i).CODE Then
'''                For j = 0 To UBound(Soko_T, 2)
'''                    If StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode) = Soko_T(i, j).HS_SOKO Then
'''                        Skip_Flg = False
'''                        Exit For
'''                    End If
'''                Next j
'''                Exit For
'''            End If
'''        Next i
'''
'''        If Skip_Flg Then
'''            Not_SHUSI = True
'''        End If

''''-----------------------------------------  �ƍ��p���ח\��̏o�͏���    2007.06.15
'''        '�ƍ��p���ח\��d���`�F�b�N
'''        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode))
'''        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'''        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''
'''        sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'''        Select Case sts
'''            Case BtNoErr
'''                Call Log_Out(LOG_F, "Y_GLICS.DAT DUP ���ƕ�=" & StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode) & "�s�d�w�s�h�c��" & StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'''            Case BtErrKeyNotFound
'''            Case Else
'''                Call File_Error(sts, BtOpGetEqual, "�ƍ��p���ח\��")
'''                Exit Function
'''        End Select
'''
'''
'''        If Not_SHUSI Then
'''            NAIGAI = "1"
'''        Else
'''            NAIGAI = Soko_T(i, j).NAIGAI
'''        End If
'''
'''        If sts = BtErrKeyNotFound Then
'''
'''            If Y_GLICS_PUT_PROC(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), NAIGAI, INS_NOW) Then
'''                Exit Function
'''            End If
'''
'''        End If
'''
'''
'''    Loop
''''-----------------------------------------  �ƍ��p���ח\��̏o�͏���    2007.06.15
'''
'''
'''
'''
'''    New_Nyuka_Update_Proc = False
    





'----------------------------------------------------------------------------
'                   �u���ח\��f�[�^�v�X�V����  F102010���ڍs
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim Skip_Flg    As Boolean
    
Dim WK_Y_QTY    As Long     '�o�א����[�N
Dim WK_Qty      As Long     '�O�؎c���[�N
Dim WK_E_QTY    As Long     '��s�o�א����[�N

Dim SUMI_QTY    As Long     '���i���ς݂Ƃ��ēo�^
Dim MI_QTY      As Long     '�����i�Ƃ��ēo�^

Dim WORK_SOKO   As String * 2
    
Dim sts         As Integer
Dim ans         As Integer
Dim Not_SHUSI   As Boolean
    
Dim wkText      As String
Dim Length      As Integer
    
    
Dim NAIGAI      As String * 1   '2007.06.15
    
    
Dim TEXT_NO     As String * 9           '÷�ć�
Dim JGYOBU_Code As String * 1           '���ƕ��敪
Dim CYOK_KBN    As String * 1           '�����敪
Dim DEN_DT      As String * 8           '�`�[���t
Dim IO_KBN      As String * 1           '���o�ɋ敪
Dim PM_KBN      As String * 1           '�ԍ��敪
Dim DEN_SYU     As String * 1           '�`�[���
Dim DEN_NO      As String * 6           '�`�[��
Dim CYU_KBN     As String * 1           '�����敪
'Dim HIN_GAI     As String * 13          '�i�ԁi�O���j  '13-->20 2016.03.07
Dim HIN_GAI13     As String * 13          '�i�ԁi�O���j   '13-->20 2016.03.07
Dim HIN_GAI20     As String * 20          '�i�ԁi�O���j   '13-->20 2016.03.07
Dim HIN_GAI     As String * 20          '�i�ԁi�O���j   '13-->20 2016.03.07


Dim HIN_NAI     As String * 13          '�i�ԁi�����j
Dim HIN_NAME    As String * 25          '�i��
Dim YOTEI_QTY   As String * 6           '����
Dim YOSAN_FROM  As String * 5           '�\�Z�P�ʁi���j
Dim YOSAN_TO    As String * 5           '�\�Z�P�ʁi��j
Dim HOST_SOKO   As String * 8           '�q�ɋ敪�iνāj
Dim HOST_TANA   As String * 8           '�I�ԁiνāj
Dim SYUK_CODE   As String * 5           '�x����^�o�א�
Dim SYUK_NAME   As String * 20          '�x����^�o�א於
Dim REC_END     As String * 1           'ں��ޏI�[ϰ�(@)
    
    
    
    
'2011.01.18
Dim GENSANKOKU          As String * 20  '���Y����
Dim GEN_GENSANKOKU      As String * 20  '�����\�����Y����
Dim SHIIRE_WORK_CENTER  As String * 8   '���ގd����ܰ�����
Dim KANKYO_KBN          As String * 3   '����ދ敪
Dim KANKYO_KBN_ST       As String * 8   '����ދ敪�K�p�J�n
Dim KANKYO_KBN_SURYO    As String * 10  '����ދ敪����
Dim ID_NO2              As String * 12  'ID_NO
Dim AITESAKI_CODE       As String * 16  '����溰��
Dim JYUCHU_YMD          As String * 8   '�󒍔N����
Dim SHITEI_NOUKI_YMD    As String * 8   '�w��[���N����


Dim GENSAN_CNT          As Integer

Dim com                 As Integer
'2011.01.18
    
    
    
    
'�o�ח\�� �ҏW�O���� ################################################################# 2005/05/16 Add ��
Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
Dim INS_NOW         As String * 14
Dim wkStr           As String
    
Dim wkMUKE_CODE     As String
    
    
Dim Loop_Cnt        As Integer          '2011.01.15

'2011.03.23
Dim MOTO_TEXT_NO    As String * 9
'2011.03.23
Dim DUP_FLG         As Boolean

    
    Fast_Flg = True

    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
'#################################################################################### 2005/05/16 Add ��
    
    New_Nyuka_Update_Proc = True


    Do Until EOF(FileNo)
        Line Input #FileNo, wkText
    
    
    
    
    
'        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 Then    '138-->145 2016.03.07
        If LenB(StrConv(wkText, vbFromUnicode)) <> 138 And _
            LenB(StrConv(wkText, vbFromUnicode)) <> 145 Then     '138-->145 2016.03.07
            
            
            
'            Call NG_File_Make_Proc
             Err_FLg = True
           
            
            Exit Do
        End If
    
    
    
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
    
    
    
    
    
    
                                                                    '÷�ć�
        Length = 1
        TEXT_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(TEXT_NO)), vbUnicode)
                                                                    '���ƕ��敪
        Length = Length + Len(TEXT_NO)
        JGYOBU_Code = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(JGYOBU_Code)), vbUnicode)
                                                                    '�����敪
        Length = Length + Len(JGYOBU_Code)
        CYOK_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYOK_KBN)), vbUnicode)
                                                                    '�`�[���t
        Length = Length + Len(CYOK_KBN)
        DEN_DT = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_DT)), vbUnicode)
                                                                    '���o�ɋ敪
        Length = Length + Len(DEN_DT)
        IO_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(IO_KBN)), vbUnicode)
                                                                    '�ԍ��敪
        Length = Length + Len(IO_KBN)
        PM_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(PM_KBN)), vbUnicode)
                                                                    '�`�[���
        Length = Length + Len(PM_KBN)
        DEN_SYU = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_SYU)), vbUnicode)
                                                                    '�`�[��
        Length = Length + Len(DEN_SYU)
        DEN_NO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(DEN_NO)), vbUnicode)
                                                                    '�����敪
        Length = Length + Len(DEN_NO)
        CYU_KBN = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(CYU_KBN)), vbUnicode)
                                                                    '�i�ԁi�O���j
        Length = Length + Len(CYU_KBN)
'        HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI)), vbUnicode)
                                                                    '�i�ԁi�����j
        If LenB(StrConv(wkText, vbFromUnicode)) = 138 Then
            HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI13)), vbUnicode)
            Length = Length + Len(HIN_GAI13)
        Else
            HIN_GAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_GAI20)), vbUnicode)
            Length = Length + Len(HIN_GAI20)
        End If
        
        
        HIN_NAI = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAI)), vbUnicode)
                                                                    '�i��
        Length = Length + Len(HIN_NAI)
        HIN_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HIN_NAME)), vbUnicode)
                                                                    '����
        Length = Length + Len(HIN_NAME)
        YOTEI_QTY = Trim(StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOTEI_QTY)), vbUnicode))
                                                                    '�\�Z�P�ʁi���j
        Length = Length + Len(YOTEI_QTY)
        YOSAN_FROM = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_FROM)), vbUnicode)
                                                                    '�\�Z�P�ʁi��j
        Length = Length + Len(YOSAN_FROM)
        YOSAN_TO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(YOSAN_TO)), vbUnicode)
                                                                    '�q�ɋ敪�iνāj
        Length = Length + Len(YOSAN_TO)
        HOST_SOKO = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_SOKO)), vbUnicode)
                                                                    '�I�ԁiνāj
        Length = Length + Len(HOST_SOKO)
        HOST_TANA = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(HOST_TANA)), vbUnicode)
                                                                    '�x����^�o�א�
        Length = Length + Len(HOST_TANA)
        SYUK_CODE = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_CODE)), vbUnicode)
                                                                    '�x����^�o�א於
        Length = Length + Len(SYUK_CODE)
        SYUK_NAME = StrConv(MidB(StrConv(wkText, vbFromUnicode), Length, Len(SYUK_NAME)), vbUnicode)
    
    
    
        '2011.01.18
        GENSANKOKU = ""             '���Y����
        GEN_GENSANKOKU = ""         '�����\�����Y����
        SHIIRE_WORK_CENTER = ""     '���ގd����ܰ�����
        KANKYO_KBN = ""             '����ދ敪
        KANKYO_KBN_ST = ""          '����ދ敪�K�p�J�n
        KANKYO_KBN_SURYO = ""       '����ދ敪����
        ID_NO2 = TEXT_NO            'ID_NO
        AITESAKI_CODE = ""          '����溰��
        JYUCHU_YMD = ""             '�󒍔N����
        SHITEI_NOUKI_YMD = ""       '�w��[���N����
        '2011.01.18
    
        Skip_Flg = True
        Not_SHUSI = False
        For i = 0 To UBound(JGYOBU_T)               '���x�敪�̃`�F�b�N
            If JGYOBU_Code = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(HOST_SOKO) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
        If Skip_Flg Then
            Not_SHUSI = True
        End If
    
    
'-----------------------------------------  �ƍ��p���ח\��̏o�͏���    2007.06.15
        '�ƍ��p���ח\��d���`�F�b�N
        Call UniCode_Conv(K0_Y_GLICS.JGYOBU, JGYOBU_Code)
        Call UniCode_Conv(K0_Y_GLICS.SYUKA_YMD, DEN_DT)
        Call UniCode_Conv(K0_Y_GLICS.TEXT_NO, TEXT_NO)

        DUP_FLG = False                 '2011.03.23

        sts = BTRV(BtOpGetEqual, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Call LOG_OUT(LOG_F, "Y_GLICS.DAT DUP ���ƕ�=" & JGYOBU_Code & "�s�d�w�s�h�c��" & TEXT_NO)
                DUP_FLG = True          '2011.03.23
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�ƍ��p���ח\��", 0)
                Exit Function
        End Select


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' �������ɂ��f�[�^�m�F  2011.03.23
        MOTO_TEXT_NO = ""
        
        If DUP_FLG Then
            If StrConv(App.EXEName, vbUpperCase) = Trim(StrConv(Y_GLICSREC.MOTO_PROG_ID, vbUnicode)) Then
            Else
                If Trim(TEXT_NO) = Trim(StrConv(Y_GLICSREC.MOTO_TEXT_NO, vbUnicode)) Then
                Else
                    MOTO_TEXT_NO = TEXT_NO
                    Mid(TEXT_NO, 5, 1) = "A"
                    DUP_FLG = False
                End If
            
            End If
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' �������ɂ��f�[�^�m�F  2011.03.23



        If Not_SHUSI Then
            NAIGAI = "1"
        Else
            NAIGAI = Soko_T(i, j).NAIGAI
        End If

'        If sts = BtErrKeyNotFound Then
        If Not DUP_FLG Then



''            If Y_GLICS_PUT_PROC(StrConv(New_HS_IN_SIJREC.JGYOBU, vbUnicode), NAIGAI, INS_NOW) Then
''                Exit Function
''            End If
            
            
'''''''''''''''''''''2011.03.23 �����ǉ�
'            If Y_GLICS_PUT_PROC(JGYOBU_Code, NAIGAI, INS_NOW, _
'                                TEXT_NO, _
'                                JGYOBU_Code, _
'                                CYOK_KBN, _
'                                DEN_DT, _
'                                IO_KBN, _
'                                PM_KBN, _
'                                DEN_SYU, _
'                                DEN_NO, _
'                                CYU_KBN, _
'                                HIN_GAI, _
'                                HIN_NAI, _
'                                HIN_NAME, _
'                                YOTEI_QTY, _
'                                YOSAN_FROM, _
'                                YOSAN_TO, _
'                                HOST_SOKO, _
'                                HOST_TANA, _
'                                SYUK_CODE, _
'                                SYUK_NAME, _
'                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD) Then
                
                
                
            If Y_GLICS_PUT_PROC(JGYOBU_Code, NAIGAI, INS_NOW, _
                                TEXT_NO, _
                                JGYOBU_Code, _
                                CYOK_KBN, _
                                DEN_DT, _
                                IO_KBN, _
                                PM_KBN, _
                                DEN_SYU, _
                                DEN_NO, _
                                CYU_KBN, _
                                HIN_GAI, _
                                HIN_NAI, _
                                HIN_NAME, _
                                YOTEI_QTY, _
                                YOSAN_FROM, _
                                YOSAN_TO, _
                                HOST_SOKO, _
                                HOST_TANA, _
                                SYUK_CODE, _
                                SYUK_NAME, _
                                GENSANKOKU, GEN_GENSANKOKU, SHIIRE_WORK_CENTER, KANKYO_KBN, KANKYO_KBN_ST, KANKYO_KBN_SURYO, ID_NO2, AITESAKI_CODE, JYUCHU_YMD, SHITEI_NOUKI_YMD, MOTO_TEXT_NO) Then
                
                
                
                
'''''''''''''''''''''2011.03.23 �����ǉ�
                
                Exit Function
            End If

        End If



'-----------------------------------------  �ƍ��p���ח\��̏o�͏���    2007.06.15
    
        Skip_Flg = False
        If Not_SHUSI Then
            Skip_Flg = True
        End If
    
        If PM_KBN = "-" Then
            Skip_Flg = True
        End If
        
        If IO_KBN <> "1" Then
            
            If IO_KBN = "4" And Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" And Trim(HOST_SOKO) = "11B" Then
''''            2011.01.18
''''                If (JGYOBU_Code = SUIHAN Or JGYOBU_Code = DENKA) Then
                If Trim(RYOHEN_TANA) <> "" Then
''''            2011.01.18
                Else
                    Skip_Flg = True
                End If
            Else
                Skip_Flg = True
            End If
        Else
            Skip_Flg = True
        End If
    
    
        WORK_SOKO = KASO_NYUKA_SOKO
    
    
    
''''2011.01.18
''''        Select Case JGYOBU_Code
''''
''''            Case SOJIKI                         '�|���@
''''
''''
''''
''''            Case DENKA, SUIHAN, SENTAKU         '�d���A���сA����@�i�A�C�����j
''''
''''
''''                Select Case MyCenter
''''
''''                    Case "O"
''''
''''
''''
''''                        '2009.06.01 65�ԑq�ɏo�͒ǉ�
''''                        If (JGYOBU_Code = SUIHAN Or JGYOBU_Code = DENKA) Then
''''                            If IO_KBN = "4" Then
''''                                If Left(YOSAN_FROM, 4) = "0211" And Left(YOSAN_TO, 3) = "SDC" Then
''''
''''                                    If Trim(HOST_SOKO) = "11B" Then
''''                                        WORK_SOKO = "65"
''''                                    End If
''''                                End If
''''                            End If
''''                        End If
''''
''''
''''
''''
''''
''''                End Select
''''             Case AIRCON                     '�G�A�R��
''''
''''
''''
''''        End Select
        If Trim(RYOHEN_TANA) <> "" Then
            WORK_SOKO = RYOHEN_TANA
        End If
''''2011.01.18
            
            
            
            
        
    
    
        If Not Skip_Flg Then
                                        
                                        
            
                
                                        '���ח\��d���`�F�b�N
            Call UniCode_Conv(K0_Y_NYU.JGYOBU, JGYOBU_Code)
            Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, DEN_DT)
            Call UniCode_Conv(K0_Y_NYU.TEXT_NO, TEXT_NO)
    
            sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
            Select Case sts
                Case BtNoErr
                    Call LOG_OUT(LOG_F, "Y_NYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�s�d�w�s�h�c��" & TEXT_NO)
                    Skip_Flg = True
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ח\��", 0)
                    Exit Function
            End Select
        
            If Not Skip_Flg Then
                                                '�g�����U�N�V�����J�n
                sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
                    Exit Function
                End If
                                            '�i�ڃ}�X�^�`�F�b�N
                If Item_Check_Proc(In_Mode, JGYOBU_Code, Soko_T(i, j).NAIGAI, HIN_GAI, HIN_NAI, HIN_NAME) Then
                    GoTo Abort_Tran
                End If
                                            
                                            
                '2012.12.20
                If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "0" And StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> "1" Then
                    Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_F)
                End If
                '2012.12.20
                                            
                                            '���׃f�[�^�쐬
                Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)
                Call UniCode_Conv(Y_NYUREC.DT_SYU, "0")
                Call UniCode_Conv(Y_NYUREC.JGYOBU, JGYOBU_Code)
                Call UniCode_Conv(Y_NYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                Call UniCode_Conv(Y_NYUREC.TEXT_NO, TEXT_NO)
        
        
                Call UniCode_Conv(Y_NYUREC.JGYOBA, "")
                Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")
                Call UniCode_Conv(Y_NYUREC.ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NO, HIN_GAI)
                Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
                Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
                Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANKA, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO, "")
                Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")
                Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")
                Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")
                Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, DEN_DT)
                Call UniCode_Conv(Y_NYUREC.TANABAN1, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN2, "")
                Call UniCode_Conv(Y_NYUREC.TANABAN3, "")
                Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")
                Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU2, "")
                Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")
                Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")
                Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")
                Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")
                Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.BIKOU1, "")
                Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")
                Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.HIN_NAME, HIN_NAME)
                Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")
                Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")
                Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")
                Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")
                Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")
                Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")
                Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")
                Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD"))
        
        
                Last_Proc_F = True              '���������ް��폜�����@���s�L��
        
        
                '���������ް��X�V
                Call UniCode_Conv(K0_J_NYU.JGYOBU, JGYOBU_Code)
                Call UniCode_Conv(K0_J_NYU.NAIGAI, Soko_T(i, j).NAIGAI)
                Call UniCode_Conv(K0_J_NYU.HIN_GAI, HIN_GAI)
    
                WK_Y_QTY = CLng(YOTEI_QTY)
    
                Loop_Cnt = 0
    
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            If CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) > WK_Y_QTY Then
                                WK_Qty = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode)) - WK_Y_QTY
                                Call UniCode_Conv(J_NYUREC.JITU_QTY, Format(WK_Qty, "00000000"))
                        
                                Loop_Cnt = 0
                        
                                Do
                                
                                    sts = BTRV(BtOpUpdate, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'
'                                            Beep
'                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                Exit Function
                                            End If
                                            DoEvents
                                            Sleep (500)
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "���������ް�", 0)
                                            Exit Function
                                    End Select
                                
                                Loop
                                WK_E_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                            Else
                                
                                
                                Loop_Cnt = 0
                                Do
                                    sts = BTRV(BtOpDelete, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                            Beep
'                                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                            If ans = vbCancel Then
'                                                Exit Function
'                                            End If
                                        
                                            Loop_Cnt = Loop_Cnt + 1
                                            If Loop_Cnt > 5 Then
                                                Exit Function
                                            End If
                                            DoEvents
                                            Sleep (500)
                                        
                                                                                    
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpDelete, "���������ް�", 0)
                                            Exit Function
                                    End Select
                                Loop
                                WK_E_QTY = CLng(StrConv(J_NYUREC.JITU_QTY, vbUnicode))
                            End If
                    
                            Exit Do
                        Case BtErrKeyNotFound
                            WK_E_QTY = 0
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                            Beep
'                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<J_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                            If ans = vbCancel Then
'                                Exit Function
'                           End If
                        
                            Loop_Cnt = Loop_Cnt + 1
                            If Loop_Cnt > 5 Then
                                Exit Function
                            End If
                            DoEvents
                            Sleep (500)
                                            
                                            
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "���������ް�", 0)
                            Exit Function
                    End Select
                Loop
                                    '��s���א��i���׎��ѐ��j
                Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, Format(WK_E_QTY, "00000000"))
        
                                    '�\�Z�P�ʌ�
                Call UniCode_Conv(Y_NYUREC.YOSAN_FROM, YOSAN_FROM)
                                    '�\�Z�P�ʐ�
                Call UniCode_Conv(Y_NYUREC.YOSAN_TO, YOSAN_TO)
                                    '�W���I��
                Call UniCode_Conv(Y_NYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode))
                Call UniCode_Conv(Y_NYUREC.HIN_NAI, HIN_NAI)
                                    'H�q�� 2006.10.17
                Call UniCode_Conv(Y_NYUREC.H_SOKO, HOST_SOKO)

                                    '���׃��X�g�o�̓t���O   2007.06.12
                Call UniCode_Conv(Y_NYUREC.NYU_LIST_OUT, " ")
            
            '2011.01.18
                Call UniCode_Conv(K0_GENSAN.JGYOBU, StrConv(Y_NYUREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, StrConv(Y_NYUREC.HIN_NO, vbUnicode))
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")
                
                com = BtOpGetGreaterEqual
                
                GENSAN_CNT = 0
                
                Do
                    DoEvents
                
                    sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            If StrConv(GENSANREC.JGYOBU, vbUnicode) <> StrConv(Y_NYUREC.JGYOBU, vbUnicode) Or _
                                StrConv(GENSANREC.NAIGAI, vbUnicode) <> StrConv(Y_NYUREC.NAIGAI, vbUnicode) Or _
                                StrConv(GENSANREC.HIN_GAI, vbUnicode) <> StrConv(Y_NYUREC.HIN_NO, vbUnicode) Then
                                Exit Do
                            End If
                        
                        
                            GENSAN_CNT = GENSAN_CNT + 1
                            If GENSAN_CNT > 1 Then
                                GENSANKOKU = ""
                                Exit Do
                            End If
                        
                        
                            GENSANKOKU = StrConv(GENSANREC.GENSANKOKU, vbUnicode)
                        
                        Case BtErrEOF
                            Exit Do
                        
                        Case Else
                            Call File_Error(sts, com, "���Y���}�X�^", 0)
                            Exit Function
                    End Select
                
                    com = BtOpGetNext
                                
                Loop
                
                
                                
                
                
                Call UniCode_Conv(Y_NYUREC.GENSANKOKU, "")                          '���Y����
                If GENSAN_CNT = 1 Then
                    Call UniCode_Conv(Y_NYUREC.GENSANKOKU, GENSANKOKU)
                End If
                                                                                    '���ގd����ܰ�����
                Call UniCode_Conv(Y_NYUREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN, KANKYO_KBN)                  '����ދ敪
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_ST, KANKYO_KBN_ST)            '����ދ敪�K�p�J�n
                Call UniCode_Conv(Y_NYUREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)      '����ދ敪����
                Call UniCode_Conv(Y_NYUREC.ID_NO2, TEXT_NO)                         'ID_NO
                Call UniCode_Conv(Y_NYUREC.AITESAKI_CODE, AITESAKI_CODE)            '����溰��
                Call UniCode_Conv(Y_NYUREC.JYUCHU_YMD, JYUCHU_YMD)                  '�󒍔N����
                Call UniCode_Conv(Y_NYUREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)      '�w��[���N����
                Call UniCode_Conv(Y_NYUREC.LIST_OUT_END_F, "")                      '���Ɋ֘Aؽďo��F
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_KANRI_F, "")                    '���ɊǗ�ؽďo��F
                Call UniCode_Conv(Y_NYUREC.LIST_NYU_CHECK_F, "")                    '��������ؽďo��F
                Call UniCode_Conv(Y_NYUREC.NYUKO_TANABAN, WORK_SOKO & "010101")     '���ɒI��
                Call UniCode_Conv(Y_NYUREC.MAEGARI_SURYO, "")                       '�O�ؑ��E��
                
                Call UniCode_Conv(Y_NYUREC.INS_TANTO, "2015")       '�ǉ��@�S����
                Call UniCode_Conv(Y_NYUREC.Ins_DateTime, INS_NOW)   '�ǉ��@����
            
                Call UniCode_Conv(Y_NYUREC.UPD_TANTO, "")           '�X�V�@�S����
                Call UniCode_Conv(Y_NYUREC.UPD_DATETIME, "")        '�X�V�@����         2005.11.15
            '2011.01.18

                
                
                
                '2011.03.23 �������v���O����
                Call UniCode_Conv(Y_NYUREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
                '2011.03.23 ���e�L�X�g��
                If Trim(MOTO_TEXT_NO) = "" Then
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, "")
                Else
                    Call UniCode_Conv(Y_NYUREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
                End If
                
                
                Call UniCode_Conv(Y_NYUREC.FILLER, "")
                
                Loop_Cnt = 0
                
                Do
                    sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                            Beep
'                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                            If ans = vbCancel Then
'                                Exit Function
'                            End If
                        
                        
                            Loop_Cnt = Loop_Cnt + 1
                            If Loop_Cnt > 5 Then
                                Exit Function
                            End If
                            DoEvents
                            Sleep (500)
                        
                        
                        Case Else
                            Call File_Error(sts, BtOpInsert, "���ח\��", 0)
                            Exit Function
                    End Select
                Loop
            
'------------ 2005.12.30
'                Select Case JGYOBU_Code
'                    Case AIRCON, SENTAKU
'                        Call UniCode_Conv(K0_SOKO.Soko_No, WORK_SOKO)
'                        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
'                        Select Case sts
'                            Case BtNoErr
'                            Case Else
'                                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
'                                Exit Function
'                        End Select
'
'                        If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
'
'                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                            MI_QTY = 0
'                        Else
'
'                            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'                                MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                                SUMI_QTY = 0
'                            Else
'                                SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                                MI_QTY = 0
'                            End If
'                        End If
'
'------------ 2005.12.30
'
'                    Case Else
'
'
'                        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
'                            MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                            SUMI_QTY = 0
'                        Else
'                            SUMI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
'                            MI_QTY = 0
'                        End If
'                End Select
                
        
'                Wk_SOKO = KASO_NYUKA_SOKO
'                If Trim(StrConv(HS_IN_SIJREC.YOSAN_FROM, vbUnicode)) <> "91H" Then
'                    Wk_SOKO = KASO_SMODOSHI_SOKO
'
'                End If
        
                
                 MI_QTY = CLng(StrConv(Y_NYUREC.SURYO, vbUnicode))
                 SUMI_QTY = 0
                
                
                '���א��ō݌Ƀf�[�^�X�V�i�{�j
'                If Nyuko_Update_Proc(JGYOBU_Code, _
'                                    Soko_T(i, j).NAIGAI, _
'                                    HIN_GAI, _
'                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
'                                    (WORK_SOKO & "01" & "01" & "01"), _
'                                    YOIN_TU_NYUKA, _
'                                    SUMI_QTY, MI_QTY, _
'                                    WS_NO, WS_NO, , _
'                                    DEN_DT & " �`��:" & DEN_NO, , , , MENU_NO) Then
'                    Exit Function
'
'                End If
            
                
                If Nyuko_Update_Proc(JGYOBU_Code, _
                                    Soko_T(i, j).NAIGAI, _
                                    HIN_GAI, _
                                    StrConv(Y_NYUREC.SYUKA_YMD, vbUnicode), _
                                    (WORK_SOKO & "01" & "01" & "01"), _
                                    YOIN_TU_NYUKA, _
                                    SUMI_QTY, MI_QTY, _
                                    WS_NO, WS_NO, , _
                                    Trim(YOSAN_FROM) & " " & DEN_DT & " �`��:" & DEN_NO, , , , MENU_NO, , RYOHEN, GENSANKOKU, SHIIRE_WORK_CENTER, ID_NO2, YOSAN_FROM) Then
                    Exit Function
            
                End If
            
            
            
                '�O�؂萔�ō݌Ƀf�[�^�X�V�i�|�j
                If WK_E_QTY <> 0 Then
                '�݌Ƀf�[�^LOCK
                    If Zaiko_Lock_Proc((WORK_SOKO & "01" & "01" & "01"), _
                                        JGYOBU_Code, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        WS_NO, , , 5) Then
                        Exit Function
    
                    End If
        
                    If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                        MI_QTY = WK_E_QTY
                    Else
                        SUMI_QTY = WK_E_QTY
                    End If
            
            
                    If Syuko_Update_Proc(JGYOBU_Code, _
                                        Soko_T(i, j).NAIGAI, _
                                        HIN_GAI, _
                                        DEN_DT, _
                                        (WORK_SOKO & "01" & "01" & "01"), _
                                        YOIN_MAE_SOUSAI, _
                                        SUMI_QTY, MI_QTY, 0, _
                                        WS_NO, WS_NO, 5) Then
                        Exit Function
        
                    End If
            
            
            
            
            
            
                End If
                
                
                sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                If sts <> BtNoErr Then
                    GoTo Abort_Tran
                End If
                
                
                Out_Cnt = Out_Cnt + 1
                lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                DoEvents
    
            End If
        
        
        
        
        
        
        
        End If
        
        
        
    
    Loop

    New_Nyuka_Update_Proc = False
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If


    




End Function
    
Private Function New_Syuka_Update_Proc(JGYOBU As String) As Boolean
'----------------------------------------------------------------------------
'                   �u�o�ח\��f�[�^�v�X�V����
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
Dim Skip_Flg    As Boolean
Dim sts         As Integer
    
Dim ans         As Integer

Dim c               As String * 128

Dim Fast_Flg        As Boolean
Dim DUP_SYUKANo     As Integer
Dim FileName        As String
Dim Ret             As Integer
    
Dim INS_NOW         As String * 14
Dim wkCHOKU_KBN     As String * 1

Dim wkSS            As String
        
Dim wkMUKE_CODE     As String

Dim Loop_Cnt        As Integer          '2011.01.15


    
    New_Syuka_Update_Proc = True

    Fast_Flg = True


    DUP_SYUKANo = FreeFile
    FileName = DUP_SYUKA_DATA

    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)


    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")

    Do
        Get #FileNo, , New_HS_OUT_SIJREC
        If Left(StrConv(New_HS_OUT_SIJREC.JGYOBA, vbUnicode), 1) < " " Then
            Exit Do
        End If
    
    
        If StrConv(New_HS_OUT_SIJREC.CRLF, vbUnicode) <> vbCrLf Then
'''            Call NG_File_Make_Proc
            Err_FLg = True
            Exit Do
        End If
    
        In_Cnt = In_Cnt + 1
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents
    
        Skip_Flg = True
        For i = 0 To UBound(JGYOBU_T)               '���x�敪�̃`�F�b�N
            If JGYOBU = JGYOBU_T(i).CODE Then
                For j = 0 To UBound(Soko_T, 2)
                    If Trim(StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode)) = Trim(Soko_T(i, j).HS_SOKO) Then
                        Skip_Flg = False
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    
                                '�i�ڃ}�X�^�̃`�F�b�N
'                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
'                Call UniCode_Conv(K0_ITEM.NAIGAI, Soko_T(i, j).NAIGAI)
'                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(HS_OUT_SIJREC.HIN_NO, vbUnicode))
'
'                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
'                Select Case sts
'                    Case BtNoErr
'                    Case BtErrKeyNotFound
'                        Skip_Flg = True
'                        Call Log_Out(LOG_F, "�`�[ID=" & StrConv(HS_OUT_SIJREC.ID_NO, vbUnicode))
'                    Case Else
'                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
'                        Exit Function
'                End Select
                
                
                
        If Not Skip_Flg Then
                                                    '�g�����U�N�V�����J�n
        sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpBeginConcurrentTransaction, "", 0)
            Exit Function
        End If
        
        
                                    '�i�ڃ}�X�^�`�F�b�N
        If Item_Check_Proc(Out_Mode, JGYOBU, Soko_T(i, j).NAIGAI, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode)) Then
            GoTo Abort_Tran
        End If
                                                        
'2008.12.16        If Len(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Or _
'2008.12.16            IsNumeric(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) Then
'2008.12.16        Else
'2008.12.16            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'2008.12.16        End If
                                                        
                                                        
        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "" Then
            wkCHOKU_KBN = ""
        Else
            wkCHOKU_KBN = "1"
        End If
                                                                                                
                                                        
        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "" Then
            If Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A1" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A2" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A3" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A4" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A5" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A6" Or _
                Trim(StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = "A7" Then
                Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "3")
            End If
        
            If StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000440" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000441" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000442" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000443" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000444" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000445" Or _
                StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) = "22000446" Then
                Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "2")
            End If
        
        End If
                                                        
        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "" Then
            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "3")
        End If
                                                                
        '2006.07.20
        wkMUKE_CODE = StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)
'-----------    2005.12.30
'''                    If JGYOBU = AIRCON Then
'''
'''                        'MTS���ނ̓ǂݑւ�
'''                        If GetIni(App.EXEName, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode), App.EXEName, c) Then
'''                        Else
'''                            Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, Trim(c))
'''                            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                        End If
'''
'''
'''                    End If
            
        '�G�A�R���������ꍇ������ɒ�������  2004.12.01-->�S���ƕ�����
        
        If (StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode)) = StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode) Then
            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
        Else
            If Len(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) <> 0 Then
                Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
            End If
        End If


        If Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode)) = "" Then
        
            If GetIni(App.EXEName, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode), App.EXEName, c) Then
            Else
                Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, Trim(c))
                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
            End If
        End If


        If StrConv(New_HS_OUT_SIJREC.HAN_KBN, vbUnicode) = "2" Then
            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "E")
            '�f�Ղ́u39040�v�ɏW�� 2006.05.31
            Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, "39040")
            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
        
        
        End If
        
        
        
        
        
        
        '�����o�ד`�[�i�ް��敪��7�A����敪=29�j�̌�����ϊ��@2006.06.17
        If Trim(StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode)) = "7" And _
            Trim(StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode)) = "29" Then

            Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, _
                Right(StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode), 5))
            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")

        End If
        
        
        
        '�|���@�̏ꍇ�́A�������߂��B 2007.10.29
        If JGYOBU = SOJIKI Then
        
            If Trim(StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode)) = "7" And _
                Trim(StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode)) = "29" And _
                Trim(StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode)) = "00023210" And _
                Trim(wkMUKE_CODE) = "09002" Then
            
            
                Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, wkMUKE_CODE)
                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
            
            
            End If
        
        
        
        
        
        End If
            
        
        
        
        
        '�����敪��6�͂Q��
''                    If JGYOBU = SENTAKU And StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode) = "S2" Then
''
''                        If StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) = "6" Then
''                            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "2")
''
''
''                        End If
''                    End If

        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "E" Then
            wkCHOKU_KBN = ""
        End If


        If Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "1" Or _
            Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "2" Or _
            Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "3" Or _
            Trim(StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode)) = "E" Then
        Else

            Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "1")
        End If

'''                    Else
            
            '����@�̏ꍇ�A���l�P�𒼑���ɃZ�b�g 2006.03.25
'''                        If JGYOBU = SENTAKU And StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode) = "S2" Then
'''                            If StrComp(StrConv(New_HS_OUT_SIJREC.ODER_NO, vbUnicode), "FAX", vbTextCompare) Then
'''
'''                                wkSS = ""
'''
'''                                For k = 1 To Len(StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode))
'''                                    If IsNumeric(Mid(StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode), k, 1)) Then
'''                                        wkSS = wkSS & Mid(StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode), k, 1)
'''                                    Else
'''                                        Exit For
'''                                    End If
'''                                Next k
'''
'''                                Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, wkSS)
'''                            End If
'''
'''                        End If
'''
'''
'''
'''                        '���̎��ƕ��͌���̂܂�
'''                        If Len(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Or _
'''                            IsNumeric(Trim(StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))) Then
'''                        Else
'''                            Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                        End If
'''                    End If
'                   ���Ɉړ�    2004.12.01
'                    If Len(Trim(StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Or _
'                        IsNumeric(Trim(StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))) Then
'                    Else
'                        Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                    End If
                                                        
                                                        
                                                        
                                                        '������}�X�^�ǂݍ���
'                    If Len(Trim(StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))) = 0 Then
'                        Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))
'                        Call UniCode_Conv(K0_MTS.SS_CODE, "")
            
'                    Else
            
            Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
            Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
'                    End If
                 
                 
                 
                                     
                 
                 
                 
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                
                
'''�O���ƕ�Ͻ��N�[ 2006.05.31
'''                            If JGYOBU = AIRCON Then
                    '�G�A�R���������ꍇ������ɒ�����Ō�����Ͻ���V�K�쐬  2004.12.01
                
                    Call UniCode_Conv(MTSREC.NAIGAI, Soko_T(i, j).NAIGAI)
                    Call UniCode_Conv(MTSREC.DATA_KBN, "")
                    Call UniCode_Conv(MTSREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(MTSREC.SS_CODE, "")
                    Call UniCode_Conv(MTSREC.MUKE_NAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(MTSREC.SS_NAME, "")
                    Call UniCode_Conv(MTSREC.MUKE_DNAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "99")
                    Call UniCode_Conv(MTSREC.FILLER, "")
                    
                    Loop_Cnt = 0
                    
                    Do
                        sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                If ans = vbCancel Then
'                                    GoTo Abort_Tran
'                                End If
                                                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                                DoEvents
                                Sleep (500)
                                                            
                                                            
                            Case Else
                                Call File_Error(sts, BtOpInsert, "������Ǘ�Ͻ�" & "key=" & StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode) & "-" & StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode), 0)
                                GoTo Abort_Tran
                        End Select
                    Loop
                                            
                                            
                                            
                                            
                
'''                            Else
'''                               '���̎��ƕ��͌���̂܂�
'''                                If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                                Else
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
'''                                    Call UniCode_Conv(New_HS_OUT_SIJREC.SS_CODE, "")
'''                                End If
''                            End If
                
'                          ���Ɉړ�    2004.12.01
'                           If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
'                               Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
'                               Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                           Else
'                               Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
'                               Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                           End If
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^", 0)
                GoTo Abort_Tran
        End Select
                                                        
                                                        
'-----------    2005.12.30
                                                        
                                                        
                                                        
                                                        
                                                        
                                                        '������}�X�^�ǂݍ���
'                    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
'                    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(HS_OUT_SIJREC.SS_CODE, vbUnicode))
'
'                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
'                    Select Case sts
'                        Case BtNoErr
'                        Case BtErrKeyNotFound
'
'                            If Soko_T(i, j).NAIGAI = NAIGAI_NAI Then
'                                Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_NAI)
'                                Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                            Else
'                                Call UniCode_Conv(HS_OUT_SIJREC.MUKE_CODE, ETC_MTS_GAI)
'                                Call UniCode_Conv(HS_OUT_SIJREC.SS_CODE, "")
'                            End If
'
'                        Case Else
'                            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
'                            Exit Function
'                    End Select
                        
        
        
'''                    If StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "1" And _
'''                        StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "2" And _
'''                        StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "3" And _
'''                        StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) <> "E" Then
'''
'''                        Call UniCode_Conv(New_HS_OUT_SIJREC.CYU_KBN, "2")
'''
'''                    End If
                    
                    
                                '�o�ח\��d���`�F�b�N
            Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
            Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
    
    
            sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    
                    
                        
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
                    
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(New_HS_OUT_SIJREC.JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, StrConv(New_HS_OUT_SIJREC.KAIKEI_JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(New_HS_OUT_SIJREC.DEN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(New_HS_OUT_SIJREC.SURYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.HOJYO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(New_HS_OUT_SIJREC.TANKA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(New_HS_OUT_SIJREC.ODER_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(New_HS_OUT_SIJREC.ITEM_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, StrConv(New_HS_OUT_SIJREC.ODER_NO_R, vbUnicode))
                    '2011.10.31
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, Left(StrConv(New_HS_OUT_SIJREC.KOSO_KEITAI, vbUnicode), 10))
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    
                    
                    
                    If TANA_SPACE Then
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, StrConv(New_HS_OUT_SIJREC.TANABAN1, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, StrConv(New_HS_OUT_SIJREC.TANABAN2, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, StrConv(New_HS_OUT_SIJREC.TANABAN3, vbUnicode))
                    End If
                    
                    
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(New_HS_OUT_SIJREC.CYU_KBN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, StrConv(New_HS_OUT_SIJREC.ORIGIN1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, StrConv(New_HS_OUT_SIJREC.ORIGIN2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(New_HS_OUT_SIJREC.BIKOU2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, StrConv(New_HS_OUT_SIJREC.HAN_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(New_HS_OUT_SIJREC.CHOKU_KBN, vbUnicode))
                    
    
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, StrConv(New_HS_OUT_SIJREC.UNIT_ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, StrConv(New_HS_OUT_SIJREC.ZAIKO_HIKIATE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, StrConv(New_HS_OUT_SIJREC.GOKON_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, StrConv(New_HS_OUT_SIJREC.JYUCHU_ZAN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, StrConv(New_HS_OUT_SIJREC.KYOKYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHOHIN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.S_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.S_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(New_HS_OUT_SIJREC.CHOHA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, StrConv(New_HS_OUT_SIJREC.JYU_HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(New_HS_OUT_SIJREC.HIN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, StrConv(New_HS_OUT_SIJREC.HIN_CHANGE_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, StrConv(New_HS_OUT_SIJREC.MODULE_EXCHANGE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAIKO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, StrConv(New_HS_OUT_SIJREC.NOUKI_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, StrConv(New_HS_OUT_SIJREC.SERVICE_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(New_HS_OUT_SIJREC.KISHU_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, StrConv(New_HS_OUT_SIJREC.ENVIRONMENT_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, StrConv(New_HS_OUT_SIJREC.KEPIN_KAIJYO, vbUnicode))
                    
                    
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    
'2008.11.28                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    Call UniCode_Conv(Y_SYUREC.UPD_NOW, INS_NOW)            '2008.11.28
                
                    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
                    
                    sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��", 0)
                        GoTo Abort_Tran
                    End If
                                
                                
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
                                
                                
                                
                    
                    Call LOG_OUT(LOG_F, "Y_SYUKA.DAT DUP ���ƕ�=" & JGYOBU & "�`�[�h�c��" & StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    Skip_Flg = True
                
                
                    If Fast_Flg Then
                        Open (FileName) For Output As DUP_SYUKANo
                        Write #DUP_SYUKANo, , , "�o�׏d�����X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS")
                        Write #DUP_SYUKANo, "�o�ד�", "�`�[��", "�x���溰��", "�q��/�r�r����", "����", "�����敪", "�����敪����", "�i��", "����", "�`�[�h�c"
                        Fast_Flg = False
                    End If
                
                
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.DEN_NO, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.CYU_KBN_NAME, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.SURYO, vbUnicode),
                    Write #DUP_SYUKANo, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode)
                
                
                
                Case BtErrKeyNotFound
                
                    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                    Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                    
                    
                                        
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    
    '''2006.07.15                    If (JGYOBU = "D" Or JGYOBU = "4") And StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode) = "E" Then
    '''2006.07.15                        Call UniCode_Conv(Y_SYUREC.NAIGAI, "2")
    '''2006.07.15                    Else
                        Call UniCode_Conv(Y_SYUREC.NAIGAI, Soko_T(i, j).NAIGAI)
    '''2006.07.15                    End If
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    
                    
                    
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, StrConv(New_HS_OUT_SIJREC.JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, StrConv(New_HS_OUT_SIJREC.DATA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, StrConv(New_HS_OUT_SIJREC.TORI_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ID_NO, StrConv(New_HS_OUT_SIJREC.ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KAIKEI_JGYOBA, StrConv(New_HS_OUT_SIJREC.KAIKEI_JGYOBA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHISAN_JGYOBA, StrConv(New_HS_OUT_SIJREC.SHISAN_JGYOBA, vbUnicode))
                    
                    
                    
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, StrConv(New_HS_OUT_SIJREC.HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, StrConv(New_HS_OUT_SIJREC.DEN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SURYO, StrConv(New_HS_OUT_SIJREC.SURYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, StrConv(New_HS_OUT_SIJREC.MUKE_CODE, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, StrConv(New_HS_OUT_SIJREC.SYUKO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.HOJYO_SYUSI, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.TANKA, StrConv(New_HS_OUT_SIJREC.TANKA, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, StrConv(New_HS_OUT_SIJREC.ODER_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, StrConv(New_HS_OUT_SIJREC.ITEM_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, StrConv(New_HS_OUT_SIJREC.ODER_NO_R, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, StrConv(New_HS_OUT_SIJREC.KOSO_KEITAI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, StrConv(New_HS_OUT_SIJREC.SYUKA_YMD, vbUnicode))
                    
                    If TANA_SPACE Then
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    Else
                        Call UniCode_Conv(Y_SYUREC.TANABAN1, StrConv(New_HS_OUT_SIJREC.TANABAN1, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN2, StrConv(New_HS_OUT_SIJREC.TANABAN2, vbUnicode))
                        Call UniCode_Conv(Y_SYUREC.TANABAN3, StrConv(New_HS_OUT_SIJREC.TANABAN3, vbUnicode))
                    End If
                    
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, StrConv(New_HS_OUT_SIJREC.MUKE_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, StrConv(New_HS_OUT_SIJREC.CYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, StrConv(New_HS_OUT_SIJREC.CYU_KBN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, StrConv(New_HS_OUT_SIJREC.ORIGIN1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, StrConv(New_HS_OUT_SIJREC.ORIGIN2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, StrConv(New_HS_OUT_SIJREC.BIKOU2, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, StrConv(New_HS_OUT_SIJREC.HAN_KBN, vbUnicode))
                    
                    '2006.07.31 ���ړ��e�����̂܂ܾ��
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, StrConv(New_HS_OUT_SIJREC.CHOKU_KBN, vbUnicode))
    '''                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, wkCHOKU_KBN)
                    
    
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, StrConv(New_HS_OUT_SIJREC.UNIT_ID_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, StrConv(New_HS_OUT_SIJREC.ZAIKO_HIKIATE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, StrConv(New_HS_OUT_SIJREC.GOKON_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, StrConv(New_HS_OUT_SIJREC.JYUCHU_ZAN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, StrConv(New_HS_OUT_SIJREC.KYOKYU_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, StrConv(New_HS_OUT_SIJREC.SHOHIN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.S_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.S_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, StrConv(New_HS_OUT_SIJREC.BIKOU1, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, StrConv(New_HS_OUT_SIJREC.CHOHA_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, StrConv(New_HS_OUT_SIJREC.JYU_HIN_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(New_HS_OUT_SIJREC.HIN_NAME, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, StrConv(New_HS_OUT_SIJREC.HIN_CHANGE_KBN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, StrConv(New_HS_OUT_SIJREC.MODULE_EXCHANGE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAIKO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_SHISAN_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.ZAN_HOJYO_SYUSI, StrConv(New_HS_OUT_SIJREC.ZAN_HOJYO_SYUSI, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, StrConv(New_HS_OUT_SIJREC.NOUKI_YMD, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, StrConv(New_HS_OUT_SIJREC.SERVICE_KANRI_NO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, StrConv(New_HS_OUT_SIJREC.KISHU_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, StrConv(New_HS_OUT_SIJREC.ENVIRONMENT_KBN, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, StrConv(New_HS_OUT_SIJREC.SS_CODE, vbUnicode))
                    
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, StrConv(New_HS_OUT_SIJREC.KEPIN_KAIJYO, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.HIN_NAI, "")
                    
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode))
                    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "0000000")
                    
                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    
                    Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")       '2006.07.20
                    Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")              '2006.07.20
                    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, wkMUKE_CODE)   '2006.07.20
                    Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")               '2006.07.20
                    
                    
                    Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")              '2006.09.07
                    Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")            '2006.09.07
                    
                    
                    Call UniCode_Conv(Y_SYUREC.UPD_NOW, "")                 '2008.11.28
                    
                    
                    Call UniCode_Conv(Y_SYUREC.FILLER, "")
    
            
    
                    
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                If ans = vbCancel Then
'                                    GoTo Abort_Tran
'                                End If
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    GoTo Abort_Tran
                                End If
                                
                                DoEvents
                                Sleep (500)
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�o�ח\��", 0)
                                GoTo Abort_Tran
                        End Select
                    Loop
    
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
        
        
        
        
                    Out_Cnt = Out_Cnt + 1
                    lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
                    DoEvents
        
        
        
        
        
        
        
                    If SYUKA_LOG_ON Then
                        Call SYUKA_LOG_OUT_PROC("INS", "AFT")
                    End If
                
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�o�ח\��", 0)
                    GoTo Abort_Tran
            End Select
                    
                                
                    
                    
                    
                    
                    
                    
                    
                    
                    
        End If
    Loop
    
        
    Close #DUP_SYUKANo
        
        
        
    New_Syuka_Update_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "", 0)
    End If

End Function
                                            '�w�b�_�[����i�u�i�ԕύX���X�g�v�j
Private Sub P_Hin_Head(Lcnt As Integer, JGYO_Kbn As String)
Dim i As Integer
Dim sts As Integer
Dim RetBuf As String

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If JGYO_Kbn = JGYOBU_T(i).CODE Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(37);
    Printer.Print "�������@�i�ԕύX���X�g�@������";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print
                                        '���׃w�b�_���
    Printer.Print "------- �i�ԁi�O���j-------";
    Printer.Print Tab(30);
    Printer.Print "------- �i�ԁi�����j-------";
    Printer.Print
    
    Printer.Print Tab(MGN_L);
    Printer.Print "��M�f�[�^";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "�}�X�^";
    Printer.Print Tab(MGN_L + 30);
    Printer.Print "��M�f�[�^";
    Printer.Print Tab(MGN_L + 44);
    Printer.Print "�}�X�^";
    Printer.Print Tab(MGN_L + 58);
    Printer.Print "�`�[���t";
    Printer.Print Tab(MGN_L + 69);
    Printer.Print "���o�ɋ�";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "���o��";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "�q";
    Printer.Print Tab(MGN_L + 96);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 103);
    Printer.Print "�o�א�"
    Printer.Print

    Lcnt = 7 + MGN_U

End Sub
                                            '���׈���i�u�i�ԕύX���X�g�v�j
Private Sub P_Hin_Proc()

Dim Lcnt As Integer
Dim Ldata As String
Dim wk_IO As String
Dim Work As String
Dim Emsg As String
Dim Wqty As Long
Dim i As Integer
Dim sts As Integer
Dim B_Jgyobu As String

    Lcnt = 99

    For i = 0 To LBox_Hin.ListCount - 1
        
'        Ldata = LBox_Hin.List(i)
'
'                                        '�w�b�_�[�R���g���[��
        If Lcnt > LMAX Or _
           B_Jgyobu <> Left(Ldata, 1) Then
            Call P_Hin_Head(Lcnt, Left(Ldata, 1))
            B_Jgyobu = Left(Ldata, 1)
        End If
'
'                                        '���׈��
'        Ldata = Mid(Ldata, 11, Len(Ldata) - 11)                     '���ƕ��C÷�ć��C�����O�@���O'
'
'        Printer.Print Tab(MGN_L);
'        Printer.Print ChrCut(Ldata, 13);                            '��M�ް��i�ԁi�O���j
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "1" Or Right(Ldata, 1) = "2" Then      '�O���i�ԕύX�H
'            Printer.Print Tab(MGN_L + 15);
'            Printer.Print Work;                                     '�}�X�^�i�ԁi�O���j
'        End If
'
'        Printer.Print Tab(MGN_L + 30);
'        Printer.Print ChrCut(Ldata, 13);                            '��M�ް��i�ԁi�����j
'        Work = ChrCut(Ldata, 13)
'        If Right(Ldata, 1) = "0" Then                               '�����i�ԕύX�H
'            Printer.Print Tab(MGN_L + 44);
'            Printer.Print Work;                                     '�}�X�^�i�ԁi�����j
'        End If
'
'        Printer.Print Tab(MGN_L + 58);                              '�`�[���t
'        Printer.Print ChrCut(Ldata, 4) & "/" & ChrCut(Ldata, 2) & "/" & ChrCut(Ldata, 2);
'
'        Printer.Print Tab(MGN_L + 69);                              '���o�ɋ敪
'        wk_IO = ChrCut(Ldata, 1)
'        Select Case wk_IO
'            Case IO_KBN_URI
'                Printer.Print wk_IO & " " & (IO_KBN_0);
'            Case IO_KBN_NYU
'                Printer.Print wk_IO & " " & (IO_KBN_1);
'            Case IO_KBN_SYU
'                Printer.Print wk_IO & " " & (IO_KBN_2);
'            Case IO_KBN_ZAT
'                Printer.Print wk_IO & " " & (IO_KBN_3);
'            Case Else
'                Printer.Print wk_IO;
'        End Select
'
'        Printer.Print Tab(MGN_L + 78);
'        Printer.Print ChrCut(Ldata, 6);                             '�`�[��
'
'        Printer.Print Tab(MGN_L + 85);                              '���o�ɐ�
'        Wqty = CLng(ChrCut(Ldata, 6))
'
'
'        sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Format(Wqty, "00000000"), Work)
'
'        Printer.Print Work;
'
'        Printer.Print Tab(MGN_L + 93);
'        Printer.Print ChrCut(Ldata, 2);                             '�q�ɋ敪�iνāj
'
'        Printer.Print Tab(MGN_L + 96);                              '�����敪
'        Select Case Left(Ldata, 1)
'            Case CYU_KBN_TUK
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_1);
'            Case CYU_KBN_SPO
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_2);
'            Case CYU_KBN_HJU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_3);
'            Case CYU_KBN_BOU
'                Printer.Print ChrCut(Ldata, 1) & " " & (CYU_KBN_E);
'            Case Else
'                Printer.Print ChrCut(Ldata, 1);
'        End Select
'
'        Printer.Print Tab(MGN_L + 103);
'        Printer.Print ChrCut(Ldata, 5);                             '�x����^�o�א�7
'
'        Printer.Print Tab(MGN_L + 110);                             '�ύX���b�Z�[�W
'        Select Case Left(Ldata, 1)
'            Case "0"
'                Printer.Print "�����ύX Ͻ��i�ԓ���";
'            Case "1"
'                Printer.Print "�O���ύX Ͻ��i�ԓ���";
'            Case "2"
'                Printer.Print "�݌ɗL�I�O���ύX�s��";
'        End Select
        
        Printer.Print LBox_Hin.List(i)
        
        Call LOG_OUT(LOG_F, LBox_Hin.List(i))
        
        Printer.Print

        Printer.Print

        Lcnt = Lcnt + 2
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
        Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    End If

End Sub

Private Sub Form_Activate()

Dim Ret         As String


Dim i           As Integer
Dim FullPath    As String


    Call NG_File_Make_Proc

    Err_FLg = False

    '---------------------------------------------  ���ƕ������C�����[�v
    For i = 0 To UBound(JGYOBU_T)
        
        In_Cnt = 0
        Out_Cnt = 0

        lblJGYOBU(i).Caption = JGYOBU_T(i).NAME
        lblJGYOBU(i).ForeColor = JGYOBU_T(i).COLOR

        lblOUTCNT(i).Caption = Format(Out_Cnt, "#0")
        lblINCNT(i).Caption = Format(In_Cnt, "#0")
        DoEvents

'2007.06.22�@���Ɏ�荞�݂𕜊�

        FileNo = FreeFile
        FileName = New_HS_IN_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)

        On Error GoTo Error_Proc

        Open FileName For Input As #FileNo

        On Error GoTo 0


        If New_Nyuka_Update_Proc(JGYOBU_T(i).CODE) Then     '���ח\��f�[�^�X�V����

            Unload Me

        End If


        Close #FileNo

        '-----------------------------------------------
    
        FileNo = FreeFile
        FileName = New_HS_OUT_SIJ

        Ret = InStr(1, Trim(FileName), ".") - 1
        FileName = Left(Trim(FileName), Ret) & "_" & JGYOBU_T(i).CODE & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
        
        On Error GoTo Error_Proc
        
        Open FileName For Binary As #FileNo
    
        On Error GoTo 0
    
    
        If New_Syuka_Update_Proc(JGYOBU_T(i).CODE) Then  '�o�ח\��f�[�^�X�V����

            Unload Me
        End If
    
    
        Close #FileNo
    
    
    
    
    Next i


    If Not Err_FLg Then
        Call NG_File_Kill_Proc
    End If


    Unload Me

Error_Proc:

    Unload Me


End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim j           As Integer

Dim c           As String * 128
Dim sts         As Integer


Dim sBuffer     As String * 255
Dim com         As String
    
Dim Max_Soko    As Integer
    
    If App.PrevInstance Then
 '       Beep
 '       MsgBox "����v���O�������s���ł��B"
        End
    End If

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '�o�׃��O�t�@�C������荞��
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "�o�׃��O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                 '�o�׏d���f�[�^�o�̓t�@�C������荞��
    If GetIni("FILE", "DUP_SYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "�o�׏d���f�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    DUP_SYUKA_DATA = Trim(c)
                               
    If JGYOB_TB_Set(1) Then      '���ƕ��̊l��
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                '�q�ɍő吔����荞��
    If GetIni(App.EXEName, "MAX_SOKO", App.EXEName, c) Then
        Max_Soko = 1
    Else
        If Not IsNumeric(RTrim(c)) Then
            Max_Soko = 1
        Else
            Max_Soko = CInt(RTrim(c))
        End If
    End If
                                    
                                
                                
                                
                                '�݌Ɏ�荞�ݗp�e�[�u���쐬
    ReDim Soko_T(0 To UBound(JGYOBU_T), 0 To Max_Soko - 1)
                                '�q�ɏ���荞��
    For i = 0 To UBound(JGYOBU_T)
        j = 0
        Do
                                '�L���q�Ɋl��
            If GetIni(App.EXEName, "SOKO" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "�q�ɏ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
    
            If Trim(c) = "**" Then  '�q�Ɏw��I��
                Exit Do
            End If
    
    
'            ReDim Preserve JSOKO_T(i).JSOKO_T(0 To j)
            Soko_T(i, j).HS_SOKO = Trim(c)
                                '�����O���l��
            If GetIni(App.EXEName, "NAIG" & JGYOBU_T(i).CODE & Format(j + 1, "0"), App.EXEName, c) Then
                Beep
                MsgBox "�����O���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
                End
            End If
            
            Soko_T(i, j).NAIGAI = Trim(c)
            j = j + 1
        Loop
    
    Next i
                                
                                
                                
                                
    '�Ǖi�ԕi���ɒI��   2011.01.18
    If GetIni(App.EXEName, "RYOHEN_TANA", App.EXEName, c) Then
        RYOHEN_TANA = ""
    Else
        RYOHEN_TANA = RTrim(c)
    End If
                                
                                
                                '�i���ɂ�鏜�O 2011.07.04
    NOT_Hin_Name_F = False
    If GetIni(App.EXEName, "NOT_HIN_NAME", App.EXEName, c) Then
    Else
        NOT_Hin_Name = Split(Trim(c), ",", -1)
        NOT_Hin_Name_F = True
    End If
                                '�i���ɂ�鏜�O 2011.07.04
                                
                                
                                '���Ƀf�[�^�t�@�C�����̊l��
    If GetIni("FILE", "NEW_HS_SIJ_IN", "SYS", c) Then
        Beep
        MsgBox "���Ƀf�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    New_HS_IN_SIJ = Trim(c)
                                
                                
                                '�Vڲ��� �o�Ƀf�[�^�t�@�C�����̊l�� 2006.05.23
    If GetIni("FILE", "NEW_HS_SIJ_OUT", "SYS", c) Then
        Beep
        MsgBox "�V�@�o�Ƀf�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    New_HS_OUT_SIJ = Trim(c)
                                
                                
                                
                                '�u�ʏ���ׁv�v���̊l��
    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
        Beep
        MsgBox "�u�ʏ���ׁv�v���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_TU_NYUKA = Trim(c)
                                
                                '�u�O�ؑ��E�v�v���̊l��
    If GetIni("YOIN", "YOIN_MAE_SOUSAI", "SYS", c) Then
        Beep
        MsgBox "�u�O�ؑ��E�v�v���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_MAE_SOUSAI = Trim(c)
                                
                                '���z���בq�ɂ̊l��
    If GetIni("SYSTEM", "KASO_NYUKA", "SYS", c) Then
        Beep
        MsgBox "���z���בq�ɂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    KASO_NYUKA_SOKO = Trim(c)
                                '���z�x���߂��q�ɂ̊l��
    If GetIni("SYSTEM", "KASO_SMODOSHI ", "SYS", c) Then
        Beep
        MsgBox "���z���בq�ɂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    KASO_SMODOSHI_SOKO = Trim(c)
                                
                                
                                '���̑�������i�����j�̊l��
    If GetIni(App.EXEName, "ETC_MTS_NAI", App.EXEName, c) Then
        Beep
        MsgBox "���̑�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    ETC_MTS_NAI = Trim(c)
                                
                                '���̑�������i�C�O�j�̊l��
    If GetIni(App.EXEName, "ETC_MTS_GAI", App.EXEName, c) Then
        Beep
        MsgBox "���̑�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    ETC_MTS_GAI = Trim(c)
                                
'---------------------------------------------- '�ƭ����̊l��    2007.11.06
    If GetIni(App.EXEName, "MENU_NO", App.EXEName, c) Then
        MENU_NO = ""
    Else
        MENU_NO = RTrim(c)
    End If
                                
                                '����@��p
    If GetIni(App.EXEName, "CENTER", "SYS", c) Then
        MyCenter = "O"
    Else
        MyCenter = Trim(c)
    End If
'---------------------------------------------- '�Ǖi�ԕi�̗v�� 2009.07.10
    RYOHEN = YOIN_TU_NYUKA
    If GetIni(App.EXEName, "RYOHEN", App.EXEName, c) Then
    Else
        RYOHEN = RTrim(c)
    End If
                                
                                
                                '���̑�������̊l��
'    If GetIni(App.EXEName, "ETC_SS_NAI", "SYS", c) Then
'        Beep
'        MsgBox "���̑�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
'        End
'    End If
'    ETC_SS_NAI = Trim(c)
                                
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

'---------------------------------------------- '�I�Ԑݒ���̊l��    2009.03.07
    If GetIni(App.EXEName, "TANA_SPACE", App.EXEName, c) Then
        TANA_SPACE = False
    Else
        If Trim(c) = "1" Then
            TANA_SPACE = True
        Else
            TANA_SPACE = False
        End If
    End If

'---------------------------------------------- '���i����̫��    2012.12.20
    If GetIni(App.EXEName, "GOODS_F", App.EXEName, c) Then
        GOODS_F = "0"
    Else
        If Trim(c) = "1" Then
            GOODS_F = "1"
        Else
            GOODS_F = "0"
        End If
    End If
'---------------------------------------------- '���i����̫��    2012.12.20



                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m '2005.12.30
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�i�X�V�p���[�N�j�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m   2005.12.30
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '���Y���}�X�^�n�o�d�m   2010.07.08
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                'PN�}�X�^�n�o�d�m   2010.09.01
    If PN_M_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    If Country_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���������ް��n�o�d�m
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�ƍ��p���ח\��n�o�d�m 2007.06.15
    If Y_GLICS_Open(BtOpenNomal) Then
        Unload Me
    End If

    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If

'���ԃ}�X�^�n�o�d�m ################################################################## 2005/05/16 Add ��
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
'#################################################################################### 2005/05/16 Add ��
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                '���Y���}�X�^�n�o�d�m 2011.01.18
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1020151.FontName
        .Size = F1020151.FontSize
    End With
    Set Printer.Font = NormalFont

    Last_Proc_F = False         '���������ް��폜�����@���s�L���t���O�N���A


    '�d������l��       2005.12.30
    i = -1
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")
    com = BtOpGetGreater
    SHIMUKE_Flg = False
    
    Do
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> P_KBN04_CD Then
                    Exit Do
                End If
            
                i = i + 1
                ReDim Preserve SHIMUKE_T(0 To i)
            
            
                SHIMUKE_Flg = True
            
                SHIMUKE_T(i).SHIMUKE_CODE = StrConv(P_CODEREC.C_Code, vbUnicode)
                SHIMUKE_T(i).JGYOBU = StrConv(P_CODEREC.OPTION1, vbUnicode)
                SHIMUKE_T(i).NAIGAI = StrConv(P_CODEREC.OPTION2, vbUnicode)
                            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^", 0)
                Unload Me
        End Select
    
        com = BtOpGetNext
    Loop
        
    
    
    '�d������l��       2005.12.30


    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    DoEvents
    
'    If Last_Proc_F = True Then              '���������ް��폜�����@���s�L��H
'        Call Last_Proc
'    End If

                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�i�X�V�p���[�N�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '���ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ח\��")
        End If
    End If
                                            '�ƍ��p���ח\��b�k�n�r�d   2007.06.16
    sts = BTRV(BtOpClose, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�ƍ��p���ח\��")
        End If
    End If
                                            '�o�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��")
        End If
    End If
                                            '���������ް��b�k�n�r�d
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���������ް�")
        End If
    End If
                                            '�a���������������Z�b�g
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020151 = Nothing

    End
End Sub
Private Function Item_Check_Proc(Mode As Integer, JGYOBU As String, NAIGAI As String, HIN_GAI As String, _
                                                                                        Optional HIN_NAI As String = "             ", _
                                                                                        Optional HIN_NAME As String = "                         ", _
                                                                                        Optional GENSANKOKU As String = "                    ", _
                                                                                        Optional GEN_GENSANKOKU As String = "                    ", _
                                                                                        Optional SHIIRE_WORK_CENTER As String = "                    ", _
                                                                                        Optional KANKYO_KBN As String = "   ", _
                                                                                        Optional KANKYO_KBN_ST As String = "        ", _
                                                                                        Optional KANKYO_KBN_SURYO As String = "          ") As Integer
'----------------------------------------------------------------------------
'                   �u�i�ڃ}�X�^�v�`�F�b�N���X�V����
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer

Dim HIN_CHANGE  As Integer

    
    
Dim BEF_GAI     As String * 13
Dim BEF_NAI     As String * 13
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
        
Dim i           As Integer
    
    
Dim sBuffer     As String * 255     '2009.01.21
Dim wkTanto     As String           '2009.01.21
    
Dim PN_M_STS    As Integer          '2010.09.01
    
    
    
Dim Loop_Cnt    As Integer          '2011.01.19
    
    
    Item_Check_Proc = True

    HIN_CHANGE = 0
    

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Loop_Cnt = 0

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                


                    If Mode = In_Mode Then          '�Γ��i�ԕύX�̃`�F�b�N
    '                Else
    
                        If Len(Trim(HIN_NAI)) <> 0 Then
                            If Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) <> Trim(HIN_NAI) Then
                                HIN_CHANGE = NAI_CHANGE
                                BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                                '�����i�ԓ���ւ�
                                Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
                            
                            
                                '�S���ҍX�V�ǉ� 2009.11.11
                                    
                                                                                        '�X�V�S����
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '�X�V����
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            
                            
                            
                            
                            
                            
                            End If
                        End If
                    
                    
                    
                                        
                        '---------------    2010.07.08  ��
                        '���Y������ւ��`�F�b�N
                        If Len(Trim(GENSANKOKU)) <> 0 Or Len(Trim(GEN_GENSANKOKU)) <> 0 Or Len(Trim(SHIIRE_WORK_CENTER)) <> 0 Then
    '                        If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) <> Trim(GENSANKOKU) Then
                                '���Y������ւ�
                                
                            
                                If Trim(GENSANKOKU) <> "" Then
                                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                                Else
                                    Debug.Print
                                End If
                                
                                
    '                            If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
    '
    '                                Call UniCode_Conv(ITEMREC.GENSANKOKU, GENSANKOKU)
    '
    '
    '                            End If
                                
                                
                                Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                                Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                            
                            
                                '�S���ҍX�V�ǉ� 2009.11.11
                                    
                                                                                        '�X�V�S����
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '�X�V����
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                                
                                
                            
    '                        End If
                        End If
                        '---------------    2010.07.08  ��
                    
                    
                        '---------------    2010.07.27  ��
                        '���敪�`�F�b�N
                        If Len(Trim(KANKYO_KBN)) <> 0 Or Len(Trim(KANKYO_KBN_ST)) <> 0 Or Len(Trim(KANKYO_KBN_SURYO)) <> 0 Then
                            
                            
                            
                            If Val(KANKYO_KBN_SURYO) = 0 Then
                            Else
                            
                                
                                '���敪����ւ�
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                                    
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                                
                                
                                '�S���ҍX�V�ǉ� 2009.11.11
                                        
                                                                                        '�X�V�S����
                                Call UniCode_Conv(ITEMREC.UPD_TANTO, "2010")
                                                                                        '�X�V����
                                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                            End If
                                
                            
                        End If
                        '---------------    2010.07.08  ��
                    
                    
                    
                    End If
                
                
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                
                com = BtOpInsert
                
                PN_M_STS = PN_M_GET(JGYOBU, HIN_GAI, 0)
                Select Case PN_M_STS
                
                    Case False
                    
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")
                    
                    Case True
                        
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")

                
                End Select
                '2010.09.01
                
                
                
                
                Call Rclr_ITEMREC               '2012.02.11
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '���ƕ�
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '�����O
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '�i�ԁi�O���j
                                                            '�i��
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
    
                    
'2009.01.21                If Mode = In_Mode Then  '�V�K�i�Ԏ�*���Z�b�g2008.10.29
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
                    
'                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'                End If
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    
                
                
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))        '�i�ԁi�����j
                Else
                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)         '�i�ԁi�����j
                End If
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '���l �z�X�g�q��
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '���l �z�X�g�I��
'                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '���ރR�[�h
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '��[�_
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '�����Ϗo�א�
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          '�T���v����
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '�ŏI���ד��t
    
'                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '�r���t���O
'                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '�g�p�q�@�h�c
'                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '�g�p���v���O����
    
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '�ŏI�ƍ����t
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '�ŏI�ƍ����݌ɐ�
'                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '�������ƕ�
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '������l
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '������萔
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Jan�R�[�h
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '�i�ԓǂݑւ�
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '���i���L��
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '������
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          '��د���I��1
                
                
                                
                
                
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��
                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '�Ɩ��Ǘ��@ �d���敪
                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           �̔��敪
                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           ���x�P��
                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           �g�����i
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           �W���e�������P���@9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           �W���e�������ݒ��
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           �W���e�������P��  9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           �W���e�������ݒ��
                                            
                                            
                                                                        '           �d������
                For i = 0 To 2
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             '����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '�d���P��
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '�P���ݒ��
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ۯĐ�
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ذ�����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ذ�����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ذ�����
                
                Next i
                                            
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           �O���݌ɋ��z
                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           ���ދ敪
                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ���x���\�t
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��

'*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '���i����   �i��
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           ���l
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           ��ЃR�[�h
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           �@��(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           �@��(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           ��
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           �v���X�`�b�N
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           ���i(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           ���i(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           ���i(3)
                
                
                                                                '           ���i(1)
                If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
                End If
                                                                
                                                                
                                                                '           ���i(2)
                If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
                End If
                                                                
                                                                
                                                                '           ���i(3)
                If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
                End If
                '2010.09.01
                
                
                
                
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           �K�p�@������
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           ��������
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           �K�p�@����l
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           ��Ǝw��
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           ���l�R
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           ���ƕ��R�[�h
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           ���萔
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           �I��(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           �I��(2)
                
                
                
'*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '���P�^�S���҃R�[�h
                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '�݌ɊǗ��ΏۗL�� 1:�Ώ� 0:�ΏۊO
    
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
    
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           �O���݌ɐ���
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           �ŏI�o�א�
    
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS�݌�(S2) �܈�p
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS�݌�(P2) �܈�p
                    
                '2010.09.01
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '���`��
                Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))
                '2010.09.01
    
    
    

    
'2010.09.01
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               '�Ưĕ��i�敪
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '�����������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '�C�O�������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '�W���P��   2006.07.28


                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      '�Ưĕ��i�敪
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '�����������i�敪   2006.07.28
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '�C�O�������i�敪   2006.07.28
                Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '�W���P��   2006.07.28
'2010.09.01
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '�ŏI�d����R�[�h   2007.05.29
                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '�ŏI�d���P��       2007.05.29
    
                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'Ұ������           2007.06.06
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'Ұ������           2007.06.06
    
    
                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '�č���ϰ�          2007.11.08
    
                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '�ː�               2008.02.14
    
                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '�`��               2008.02.14
                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '�ގ�               2008.02.14
                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              '����ްف@����      2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 '����ްٻ��ށiW�j   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 '����ްٻ��ށiD�j   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 '����ްٻ��ށiH�j   2008.02.14
        
                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '�������^���Ȃ�   2008.02.14
            
        
                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '���i���@�H��       2008.02.14
        
                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '���i���@�H������   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '���i���@�H������   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '���i���@�P���ݒ�� 2008.02.14
        
    
                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '���i���@���ތ���   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '���i���@���ޔ���   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '���i���@�P���ݒ�� 2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '�A�����@�o���׸�   2008.02.14
    
                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '�g�p�e�[�v���     2008.02.14
                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '�g�p�e�[�v��       2008.02.14
    
                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '�I�ԃ}�[�N         2008.04.02
    
    
                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '�����P���@����     2008.04.15
    
                '2010.07.08 ��
                'Call UniCode_Conv(ITEMREC.GENSANKOKU, "")              '���Y��             2008.06.11
                Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")              '���Y��
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '���Y��
                Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))
                '2010.09.01
                
                
                If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Else
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")
                End If
                '2010.07.08 ��
    
                '2010.09.01
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
                    
                    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
                    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
                    Select Case sts
                        Case BtNoErr
                            Debug.Print
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(CountryREC.CountryName2, "")
                        Case Else
                    
                            Call File_Error(sts, BtOpGetEqual, "COUNTRY")
                            Exit Function
                    
                    End Select
                
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
                
                
                End If
                '2010.09.01
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '�O���P�� 9(8)V99   2008.06.12
                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC���H�P��9(8)   2008.06.12
                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU���H�P��9(8)     2008.06.12
    
    
                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '���Y���b�g         2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '�����[�g           2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '�W������           2008.07.07
    
    
                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '�P���ݒ�S����     2008.07.09
    
                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '�d������           2008.07.09

                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '�����敪           2008.07.16

                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            '���x���\�薇��     2008.07.19

                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '���ތ���     �@    2008.08.20�ǉ�
                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '��������           2008.08.20�ǉ�
         

'*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
                
                
                                
                
                
                
                '��2009.02.20
                For i = 0 To 9
                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")

                Next i


                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
                '��2009.02.20
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.STAT, "1")                    '��ԋ敪           2009.01.21
    

                sBuffer = Space(255)                                    '2009.01.21
                If GetComputerNameA(sBuffer, 255) <> 0 Then
                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                Else
                    wkTanto = "???"
                End If

                
                
                
                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '���iү���� 2009.08.28
                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '�����e 2009.08.28
                
                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
            
                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                
                
                
                
                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
                
                
                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
                    
                    If Val(KANKYO_KBN_SURYO) <> 0 Then
                    
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                
                    End If
            
                End If
                
                
                                                                        '�ǉ��S����
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
                Else
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
                End If
                '2010.09.01
                                                                        
                                                                        '�ǉ�����
                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                
                
                
                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
    
                
                
                Call UniCode_Conv(ITEMREC.BIKOU20, "")
                
                '2011.07.05
                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
                
                If NOT_Hin_Name_F Then
                    For i = 0 To UBound(NOT_Hin_Name)
                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
                            Exit For
                        End If
                    Next i
                End If
                '2011.07.05
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '�X�V�S����
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '�X�V����
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                                
                
                
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
           
            
            Case BtErrDEAD_LOCK
                Exit Function
                        
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^", 0)
                Exit Function
        End Select
    Loop
    
    Loop_Cnt = 0
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
            
                DoEvents
                Sleep (500)
            
            
            
            Case BtErrDEAD_LOCK
                Exit Function
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^", 0)
                Exit Function
        End Select
    Loop
        
    If SHIMUKE_Flg Then
        If com = BtOpInsert Then
            '�\���}�X�^�̒ǉ�       2005.12.30
            For i = 0 To UBound(SHIMUKE_T)
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIMUKE_T(i).JGYOBU And _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) = SHIMUKE_T(i).NAIGAI Then
                                                                            '�d�����溰��
                    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, SHIMUKE_T(i).SHIMUKE_CODE)
                                                                            '���ƕ�
                    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, SHIMUKE_T(i).JGYOBU)
                                                                            '�����O
                    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, SHIMUKE_T(i).NAIGAI)
                                                                            '�i��
                    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            '�ް��敪
                    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
                                                                            '�ǔ�
                    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
                                                                            '��{�N���X
                    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, "")
                                                                            '���l
                    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "")
                    
                    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
                                                                            '�X�V�S����
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                            '�X�V����
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
'                                If ans = vbCancel Then
'                                    Exit Function
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                            
                                DoEvents
                                Sleep (500)
                            
                            
                            Case BtErrDEAD_LOCK
                                Exit Function
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�\���}�X�^", 0)
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If
        
    If HIN_CHANGE <> 0 Then
        LBox_Hin.AddItem JGYOBU & NAIGAI & StrConv(ITEMREC.HIN_GAI, vbUnicode) & BEF_GAI & StrConv(ITEMREC.HIN_NAI, vbUnicode) & BEF_NAI & NAI_CHANGE
    End If

    Item_Check_Proc = False

End Function

Private Function old_Item_Check_Proc(Mode As Integer, JGYOBU As String, NAIGAI As String, HIN_GAI As String, _
                                                                                        Optional HIN_NAI As String = "             ", _
                                                                                        Optional HIN_NAME As String = "                         ", _
                                                                                        Optional GENSANKOKU As String = "                    ", _
                                                                                        Optional GEN_GENSANKOKU As String = "                    ", _
                                                                                        Optional SHIIRE_WORK_CENTER As String = "                    ", _
                                                                                        Optional KANKYO_KBN As String = "   ", _
                                                                                        Optional KANKYO_KBN_ST As String = "        ", _
                                                                                        Optional KANKYO_KBN_SURYO As String = "          ") As Integer
'----------------------------------------------------------------------------
'                   �u�i�ڃ}�X�^�v�`�F�b�N���X�V����
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer

Dim HIN_CHANGE  As Integer

    
    
Dim BEF_GAI     As String * 13
Dim BEF_NAI     As String * 13
    
Dim SUMI_QTY    As Long
Dim MI_QTY      As Long
        
Dim i           As Integer
    
Dim sBuffer     As String * 255     '2009.01.21
Dim wkTanto     As String           '2009.01.21
    
Dim Loop_Cnt        As Integer          '2011.01.15
    
Dim PN_M_STS    As Integer          '2010.09.01
    
    
    old_Item_Check_Proc = True

    HIN_CHANGE = 0
    
    
           
'    If Mode = Out_Mode Then
'        Item_Check_Proc = False
'        Exit Function
'    End If
    
'    If Len(Trim(StrConv(HS_IN_SIJREC.HIN_NAI, vbUnicode))) <> 0 Then
'
'        Call UniCode_Conv(K2_ITEM.JGYOBU, JGYOBU)
'        Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI)
'        Call UniCode_Conv(K2_ITEM.HIN_NAI, StrConv(HS_IN_SIJREC.HIN_NAI, vbUnicode))
'
'
'        Do                          '�ΊO�i�ԓ���ւ��̃��[�v
'
'            sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'            Select Case sts
'                Case BtNoErr
'                    If StrConv(ITEMREC.HIN_GAI, vbUnicode) <> Trim(HIN_GAI) Then
'                                    '�O���i�Ԃ̓���ւ��ׂ̈̍݌ɗL���`�F�b�N
'                        If Zaiko_Syukei_Proc(Sumi_Qty, MI_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
'                            Exit Function
'                        End If
'
'
'                        If (Sumi_Qty + MI_Qty) = 0 Then
'                            Do
'                                sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'                                Select Case sts
'                                    Case BtNoErr
'                                        Exit Do
'
'
'                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                        Beep
'                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                                        If ans = vbCancel Then
'                                            Exit Function
'                                        End If
'                                    Case Else
'                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
'                                        Exit Function
'                                End Select
'                            Loop
'
'
'                            HIN_CHANGE = GAI_CHANGE
'                            BEF_GAI = HIN_GAI
'
'                        Else
'                            '�i�ԓ���ւ��s��
'                            sts = BTRV(BtOpUnlock, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
'                            If sts Then
'                                Call File_Error(sts, BtOpUnlock, "�i�ڃ}�X�^")
'                                Exit Function
'                            End If
'                            HIN_CHANGE = NOT_GAI_CHANGE
'                            BEF_GAI = HIN_GAI
'
'                        End If
'                    End If
'
'                    Exit Do
'
'                Case BtErrKeyNotFound
'                    Exit Do
'                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                    Beep
'                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                    If ans = vbCancel Then
'                        Exit Function
'                    End If
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
'                    Exit Function
'            End Select
'
'        Loop
'
'    End If

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Loop_Cnt = 0

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr

                If Mode = In_Mode Then          '�Γ��i�ԕύX�̃`�F�b�N
'                Else

                    If Len(Trim(Trim(HIN_NAI))) <> 0 Then
                        If Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode)) <> Trim(HIN_NAI) Then
                            HIN_CHANGE = NAI_CHANGE
                            BEF_NAI = StrConv(ITEMREC.HIN_NAI, vbUnicode)
                            '�����i�ԓ���ւ�
                            Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
                        
                        
                            '�S���ҍX�V�ǉ� 2009.11.11
                                
                                                                                    '�X�V�S����
                            Call UniCode_Conv(ITEMREC.UPD_TANTO, "2015")
                                                                                    '�X�V����
                            Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                        
                        
                        End If
                    End If
                End If
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                
                
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �����ւ�    2012.07.05
                com = BtOpInsert
                
                PN_M_STS = PN_M_GET(JGYOBU, HIN_GAI, 0)
                Select Case PN_M_STS
                
                    Case False
                    
                    Case BtErrKeyNotFound
                    
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")
                    
                    Case True
                        
                        Call UniCode_Conv(PN_MREC.UnitKbn, "")
                        Call UniCode_Conv(PN_MREC.NaiKbn, "")
                        Call UniCode_Conv(PN_MREC.GaiKbn, "")
                        Call UniCode_Conv(PN_MREC.HyoTan, "")
                        Call UniCode_Conv(PN_MREC.Tanka2, "")
                        Call UniCode_Conv(PN_MREC.Tanka3, "")
                        Call UniCode_Conv(PN_MREC.Tanka4, "")
                        Call UniCode_Conv(PN_MREC.MadeIn, "")
                        Call UniCode_Conv(PN_MREC.MadeInCode, "")

                
                End Select
                '2010.09.01
                
                
                
                
                Call Rclr_ITEMREC               '2012.02.11
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '���ƕ�
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '�����O
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '�i�ԁi�O���j
                                                            '�i��
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
    
                    
'2009.01.21                If Mode = In_Mode Then  '�V�K�i�Ԏ�*���Z�b�g2008.10.29
                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
                    
'                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'                End If
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
    
                
                
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.HIN_NAI, StrConv(PN_MREC.SPn, vbUnicode))        '�i�ԁi�����j
                Else
                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)         '�i�ԁi�����j
                End If
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '���l �z�X�g�q��
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '���l �z�X�g�I��
'                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '���ރR�[�h
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '��[�_
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '�����Ϗo�א�
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          '�T���v����
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '�ŏI���ד��t
    
'                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '�r���t���O
'                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '�g�p�q�@�h�c
'                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '�g�p���v���O����
    
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '�ŏI�ƍ����t
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '�ŏI�ƍ����݌ɐ�
'                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '�������ƕ�
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '������l
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '������萔
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Jan�R�[�h
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '�i�ԓǂݑւ�
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '���i���L��
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '������
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          '��د���I��1
                
                
                                
                
                
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��
                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '�Ɩ��Ǘ��@ �d���敪
                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           �̔��敪
                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           ���x�P��
                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           �g�����i
                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           �W���e�������P���@9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           �W���e�������ݒ��
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           �W���e�������P��  9(8)V99
                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           �W���e�������ݒ��
                                            
                                            
                                                                        '           �d������
                For i = 0 To 2
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             '����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '�d���P��
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '�P���ݒ��
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ۯĐ�
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ذ�����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ذ�����
                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ذ�����
                
                Next i
                                            
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           �O���݌ɋ��z
                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           ���ދ敪
                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ���x���\�t
'*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��

'*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '���i����   �i��
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           ���l
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           ��ЃR�[�h
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           �@��(1)
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           �@��(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           ��
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           �v���X�`�b�N
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           ���i(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           ���i(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           ���i(3)
                
                
                                                                '           ���i(1)
                If IsNumeric(StrConv(PN_MREC.Tanka2, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, Format(Val(StrConv(PN_MREC.Tanka2, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN1, "0000000000")
                End If
                                                                
                                                                
                                                                '           ���i(2)
                If IsNumeric(StrConv(PN_MREC.Tanka3, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, Format(Val(StrConv(PN_MREC.Tanka3, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN2, "0000000000")
                End If
                                                                
                                                                
                                                                '           ���i(3)
                If IsNumeric(StrConv(PN_MREC.Tanka4, vbUnicode)) Then
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, Format(Val(StrConv(PN_MREC.Tanka4, vbUnicode)), "0000000000"))
                Else
                    Call UniCode_Conv(ITEMREC.L_URIKIN3, "0000000000")
                End If
                '2010.09.01
                
                
                
                
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           �K�p�@������
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           ��������
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           �K�p�@����l
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           ��Ǝw��
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           ���l�R
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           ���ƕ��R�[�h
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           ���萔
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           �I��(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           �I��(2)
                
                
                
'*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '���P�^�S���҃R�[�h
                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '�݌ɊǗ��ΏۗL�� 1:�Ώ� 0:�ΏۊO
    
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
    
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           �O���݌ɐ���
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           �ŏI�o�א�
    
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS�݌�(S2) �܈�p
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS�݌�(P2) �܈�p
                    
                '2010.09.01
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '���`��
                Call UniCode_Conv(ITEMREC.K_KEITAI, StrConv(PN_MREC.KKeitai, vbUnicode))
                '2010.09.01
    
    
    

    
'2010.09.01
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               '�Ưĕ��i�敪
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '�����������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '�C�O�������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '�W���P��   2006.07.28


                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, StrConv(PN_MREC.UnitKbn, vbUnicode))      '�Ưĕ��i�敪
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, StrConv(PN_MREC.NaiKbn, vbUnicode))        '�����������i�敪   2006.07.28
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, StrConv(PN_MREC.GaiKbn, vbUnicode))        '�C�O�������i�敪   2006.07.28
                Call UniCode_Conv(ITEMREC.HYO_TANKA, StrConv(PN_MREC.HyoTan, vbUnicode))        '�W���P��   2006.07.28
'2010.09.01
    
    
    
    
    
    
    
    
    
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '�ŏI�d����R�[�h   2007.05.29
                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '�ŏI�d���P��       2007.05.29
    
                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'Ұ������           2007.06.06
                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'Ұ������           2007.06.06
    
    
                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '�č���ϰ�          2007.11.08
    
                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '�ː�               2008.02.14
    
                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '�`��               2008.02.14
                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '�ގ�               2008.02.14
                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              '����ްف@����      2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 '����ްٻ��ށiW�j   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 '����ްٻ��ށiD�j   2008.02.14
                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 '����ްٻ��ށiH�j   2008.02.14
        
                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '�������^���Ȃ�   2008.02.14
            
        
                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '���i���@�H��       2008.02.14
        
                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '���i���@�H������   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '���i���@�H������   2008.02.14
                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '���i���@�P���ݒ�� 2008.02.14
        
    
                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '���i���@���ތ���   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '���i���@���ޔ���   2008.02.14
                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '���i���@�P���ݒ�� 2008.02.14
    
    
                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '�A�����@�o���׸�   2008.02.14
    
                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '�g�p�e�[�v���     2008.02.14
                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '�g�p�e�[�v��       2008.02.14
    
                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '�I�ԃ}�[�N         2008.04.02
    
    
                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '�����P���@����     2008.04.15
    
                '2010.07.08 ��
                'Call UniCode_Conv(ITEMREC.GENSANKOKU, "")              '���Y��             2008.06.11
                Call UniCode_Conv(ITEMREC.xGENSANKOKU, "")              '���Y��
                
                
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '���Y��
                Call UniCode_Conv(ITEMREC.GENSANKOKU, StrConv(PN_MREC.MadeIn, vbUnicode))
                '2010.09.01
                
                
                If Trim(GENSANKOKU) <> "" Or Trim(GEN_GENSANKOKU) <> "" Or Trim(SHIIRE_WORK_CENTER) <> "" Then
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, GEN_GENSANKOKU)
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)
                Else
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_GEN_GENSANKOKU, "")
                    Call UniCode_Conv(ITEMREC.TORI_SHIIRE_WORK_CENTER, "")
                End If
                '2010.07.08 ��
    
                '2010.09.01
                If Trim(StrConv(ITEMREC.TORI_GENSANKOKU, vbUnicode)) = "" Then
                    
                    Call UniCode_Conv(K0_Country.CountryCode, StrConv(PN_MREC.MadeInCode, vbUnicode))
                    sts = BTRV(BtOpGetEqual, Country_POS, CountryREC, Len(CountryREC), K0_Country, Len(K0_Country), 0)
                    Select Case sts
                        Case BtNoErr
                            Debug.Print
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(CountryREC.CountryName2, "")
                        Case Else
                    
                            Call File_Error(sts, BtOpGetEqual, "COUNTRY")
                            Exit Function
                    
                    End Select
                
                    Call UniCode_Conv(ITEMREC.TORI_GENSANKOKU, StrConv(CountryREC.CountryName2, vbUnicode))
                
                
                End If
                '2010.09.01
    
    
    
    
    
                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '�O���P�� 9(8)V99   2008.06.12
                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC���H�P��9(8)   2008.06.12
                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU���H�P��9(8)     2008.06.12
    
    
                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '���Y���b�g         2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '�����[�g           2008.07.07
                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '�W������           2008.07.07
    
    
                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '�P���ݒ�S����     2008.07.09
    
                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '�d������           2008.07.09

                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '�����敪           2008.07.16

                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            '���x���\�薇��     2008.07.19

                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '���ތ���     �@    2008.08.20�ǉ�
                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '��������           2008.08.20�ǉ�
         

'*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
                
                
                                
                
                
                
                '��2009.02.20
                For i = 0 To 9
                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")

                Next i


                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
                '��2009.02.20
                
                
                
                
                
                Call UniCode_Conv(ITEMREC.STAT, "1")                    '��ԋ敪           2009.01.21
    

                sBuffer = Space(255)                                    '2009.01.21
                If GetComputerNameA(sBuffer, 255) <> 0 Then
                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
                Else
                    wkTanto = "???"
                End If

                
                
                
                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '���iү���� 2009.08.28
                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '�����e 2009.08.28
                
                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
            
                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
                
                
                
                
                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
                
                
                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
                    
                    If Val(KANKYO_KBN_SURYO) <> 0 Then
                    
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
                
                    End If
            
                End If
                
                
                                                                        '�ǉ��S����
                '2010.09.01
'                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
                If Mode = Out_Mode Then
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
                Else
                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
                End If
                '2010.09.01
                                                                        
                                                                        '�ǉ�����
                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))


                
                
                
                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
    
                
                
                Call UniCode_Conv(ITEMREC.BIKOU20, "")
                
                '2011.07.05
                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
                
                If NOT_Hin_Name_F Then
                    For i = 0 To UBound(NOT_Hin_Name)
                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
                            Exit For
                        End If
                    Next i
                End If
                '2011.07.05
                
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '�X�V�S����
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '�X�V����
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                                
                
                
                
                Exit Do




                
                
'                com = BtOpInsert
'
'                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)   '���ƕ�
'                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)   '�����O
'                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI) '�i�ԁi�O���j
'
'                If Mode = In_Mode Then                      '�i��
'                    Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)
'                Else
'                    Call UniCode_Conv(ITEMREC.HIN_NAME, StrConv(New_HS_OUT_SIJREC.HIN_NAME, vbUnicode))
'                End If
'
''                If Mode = Out_Mode Then                      '�W���I��
''                    If Len(Trim(HS_IN_SIJREC.TANABAN1)) = 0 Then
''                        Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
''                        Call UniCode_Conv(ITEMREC.ST_SOKO, "")
''                        Call UniCode_Conv(ITEMREC.ST_RETU, "")
''                        Call UniCode_Conv(ITEMREC.ST_REN, "")
''                        Call UniCode_Conv(ITEMREC.ST_DAN, "")
''                    Else
''                        Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))
''                        Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 1, 2)))
''                        Call UniCode_Conv(ITEMREC.ST_RETU, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 3, 2)))
''                        Call UniCode_Conv(ITEMREC.ST_REN, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 5, 2)))
''                        Call UniCode_Conv(ITEMREC.ST_DAN, Mid(StrConv(HS_IN_SIJREC.TANABAN1, 7, 2)))
''
''                    End If
''                Else
'                    Call UniCode_Conv(ITEMREC.ST_SET_DT, "")
''2009.01.21 "**" �ɕύX                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
''                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
''                    Call UniCode_Conv(ITEMREC.ST_REN, "")
''                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
'
'                    Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
'                    Call UniCode_Conv(ITEMREC.ST_RETU, "**")
'                    Call UniCode_Conv(ITEMREC.ST_REN, "**")
'                    Call UniCode_Conv(ITEMREC.ST_DAN, "**")
'
''                End If
'
'                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")
'                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
'                Call UniCode_Conv(ITEMREC.BEF_REN, "")
'                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
'
'                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")
'                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")
'
'                If Mode = In_Mode Then          '�Γ��i��
'                    Call UniCode_Conv(ITEMREC.HIN_NAI, HIN_NAI)
'                Else
'                    Call UniCode_Conv(ITEMREC.HIN_NAI, "")
'                End If
'
'                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '���l �z�X�g�q��
'                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '���l �z�X�g�I��
''                Call UniCode_Conv(ITEMREC.SIZAI_CD, "")             '���ރR�[�h
'                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '��[�_
'                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '�����Ϗo�א�
'                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          '�T���v����
'                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '�ŏI���ד��t
'
''                Call UniCode_Conv(ITEMREC.LOCK_F, "")               '�r���t���O
''                Call UniCode_Conv(ITEMREC.WEL_ID, "")               '�g�p�q�@�h�c
''                Call UniCode_Conv(ITEMREC.PRG_ID, "")               '�g�p���v���O����
'
'                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '�ŏI�ƍ����t
'                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '�ŏI�ƍ����݌ɐ�
''                Call UniCode_Conv(ITEMREC.MOTO_JIGYOBU, "")         '�������ƕ�
'                Call UniCode_Conv(ITEMREC.BIKOU, "")                '������l
'                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '������萔
'                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Jan�R�[�h
'                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '�i�ԓǂݑւ�
'                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '���i���L��
'                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '������
'                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          '��د���I��1
'                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          '��د���I��1
'                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          '��د���I��1
'
'
'
'
'
''*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��
'                Call UniCode_Conv(ITEMREC.G_SHIIRE_KBN, "")             '�Ɩ��Ǘ��@ �d���敪
'                Call UniCode_Conv(ITEMREC.G_HANBAI_KBN, "")             '           �̔��敪
'                Call UniCode_Conv(ITEMREC.G_SYUSHI, "")                 '           ���x�P��
'                Call UniCode_Conv(ITEMREC.G_KUMITATE, "")               '           �g�����i
'                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")              '           �W���e�������P���@9(8)V99
'                Call UniCode_Conv(ITEMREC.G_ST_URITAN_DT, "")           '           �W���e�������ݒ��
'                Call UniCode_Conv(ITEMREC.G_ST_SHITAN, "")              '           �W���e�������P��  9(8)V99
'                Call UniCode_Conv(ITEMREC.G_ST_SHITAN_DT, "")           '           �W���e�������ݒ��
'
'
'                                                                        '           �d������
'                For i = 0 To 2
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).CODE, "")             '����
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA, "")            '�d���P��
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).TANKA_DT, "")         '�P���ݒ��
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LOT, "")              'ۯĐ�
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LEAD_TIME, "")        'ذ�����
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_DT, "")    'ذ�����
'                    Call UniCode_Conv(ITEMREC.G_SHIIRE_TBL(i).LAST_ORDER_QTY, "")   'ذ�����
'
'                Next i
'
'                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_KIN, "")          '           �O���݌ɋ��z
'                Call UniCode_Conv(ITEMREC.G_SHIZAI_KBN, "")             '           ���ދ敪
'                Call UniCode_Conv(ITEMREC.G_LABEL_NON, P_G_LABEL_ON)    '           ���x���\�t
''*------------------------------------------ 2005.11.15 �ǉ�(�Ɩ��Ǘ�����) ��
'
''*------------------------------------------ 2005.11.15 �ǉ�(���i���ٍ���) ��
'                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '���i����   �i��
'                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           ���l
'                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           ��ЃR�[�h
'                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           �@��(1)
'                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
'                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           �@��(3)
'                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           ��
'                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           �v���X�`�b�N
'                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           ���i(1)
'                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           ���i(2)
'                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           ���i(3)
'                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           �K�p�@������
'                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           ��������
'                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           �K�p�@����l
'                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           ��Ǝw��
'                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           ���l�R
'                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           ���ƕ��R�[�h
'                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           ���萔
'                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           �I��(1)
'                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           �I��(2)
'
'
'
'
'
''*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
'
'                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '���P�^�S���҃R�[�h
'                Call UniCode_Conv(ITEMREC.ZAIKO_F, "")                  '�݌ɊǗ��ΏۗL�� 1:�Ώ� 0:�ΏۊO
'
'                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '           �@��(2)
'
'                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "")          '           �O���݌ɐ���
'                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "")         '           �ŏI�o�א�
'
'                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "")             'GLICS�݌�(S2) �܈�p
'                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "")             'GLICS�݌�(P2) �܈�p
'
'                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '���`��
'
'
'                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               '�Ưĕ��i�敪
'                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '�����������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '�C�O�������i�敪   2006.07.28
'                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '�W���P��   2006.07.28
'
'                Call UniCode_Conv(ITEMREC.LAST_CODE, "")                '�ŏI�d����R�[�h   2007.05.29
'                Call UniCode_Conv(ITEMREC.LAST_TANKA, "")               '�ŏI�d���P��       2007.05.29
'
'                Call UniCode_Conv(ITEMREC.MAKER_CODE, "")               'Ұ������           2007.06.06
'                Call UniCode_Conv(ITEMREC.MAKER_NAME, "")               'Ұ������           2007.06.06
'
'
'                Call UniCode_Conv(ITEMREC.L_MARK, "")                   '�č���ϰ�          2007.11.08
'
'                Call UniCode_Conv(ITEMREC.SAI_SU, "")                   '�ː�               2008.02.14
'
'                Call UniCode_Conv(ITEMREC.D_KEISHIKI, "")               '�`��               2008.02.14
'                Call UniCode_Conv(ITEMREC.D_MATERIAL, "")               '�ގ�               2008.02.14
'                Call UniCode_Conv(ITEMREC.D_THICKNESS, "")              '����ްف@����      2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.D_SIZE_W, "")                 '����ްٻ��ށiW�j   2008.02.14
'                Call UniCode_Conv(ITEMREC.D_SIZE_D, "")                 '����ްٻ��ށiD�j   2008.02.14
'                Call UniCode_Conv(ITEMREC.D_SIZE_H, "")                 '����ްٻ��ށiH�j   2008.02.14
'
'                Call UniCode_Conv(ITEMREC.D_PRINT, "")                  '�������^���Ȃ�   2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.S_KOUSU, "")                  '���i���@�H��       2008.02.14
'
'                Call UniCode_Conv(ITEMREC.S_KOUSU_GENKA, "")            '���i���@�H������   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_KOUSU_BAIKA, "")            '���i���@�H������   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_KOUSU_SET_DATE, "")         '���i���@�P���ݒ�� 2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.S_SHIZAI_GENKA, "")           '���i���@���ތ���   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_SHIZAI_BAIKA, "")           '���i���@���ޔ���   2008.02.14
'                Call UniCode_Conv(ITEMREC.S_SHIZAI_SET_DATE, "")        '���i���@�P���ݒ�� 2008.02.14
'
'
'                Call UniCode_Conv(ITEMREC.SE_USOU_F, "")                '�A�����@�o���׸�   2008.02.14
'
'                Call UniCode_Conv(ITEMREC.USE_TAPE_KIND, "")            '�g�p�e�[�v���     2008.02.14
'                Call UniCode_Conv(ITEMREC.USE_TAPE_LNG, "")             '�g�p�e�[�v��       2008.02.14
'
'                Call UniCode_Conv(ITEMREC.H_TANA_MAKE, "")              '�I�ԃ}�[�N         2008.04.02
'
'
'                Call UniCode_Conv(ITEMREC.SE_TANKA_MEMO, "")            '�����P���@����     2008.04.15
'
'
'                Call UniCode_Conv(ITEMREC.GENSANKOKU, "")               '���Y��             2008.06.11
'
'
'
'                Call UniCode_Conv(ITEMREC.S_GAISO_TANKA, "")            '�O���P�� 9(8)V99   2008.06.12
'                Call UniCode_Conv(ITEMREC.S_PPSC_KAKO_KOSU, "")         'PPSC���H�P��9(8)   2008.06.12
'                Call UniCode_Conv(ITEMREC.S_BU_KAKO_KOSU, "")           'BU���H�P��9(8)     2008.06.12
'
'
'                Call UniCode_Conv(ITEMREC.SEI_LOT, "")                  '���Y���b�g         2008.07.07
'                Call UniCode_Conv(ITEMREC.SEI_RATE, "")                 '�����[�g           2008.07.07
'                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")              '�W������           2008.07.07
'
'
'                Call UniCode_Conv(ITEMREC.SEI_TANKA_TANTO, "")          '�P���ݒ�S����     2008.07.09
'
'                Call UniCode_Conv(ITEMREC.SHIMUKE_CODE, "")             '�d������           2008.07.09
'
'                Call UniCode_Conv(ITEMREC.SEI_KBN, "")                  '�����敪           2008.07.16
'
'                Call UniCode_Conv(ITEMREC.SEI_LABEL_QTY, "")            '���x���\�薇��     2008.07.19
'
'                Call UniCode_Conv(ITEMREC.SEI_SZI_CNT, "")              '���ތ���     �@    2008.08.20�ǉ�
'                Call UniCode_Conv(ITEMREC.SEI_DKN_CNT, "")              '��������           2008.08.20�ǉ�
'
'*------------------------------------------ 2008.08.26 �V�K�ǉ����ڈꎮ ��
'
'                '��2009.02.20
'                For i = 0 To 9
'                    Call UniCode_Conv(ITEMREC.BEF_KOUTEI(i).BEF_KOUTEI, "")
'                    Call UniCode_Conv(ITEMREC.MAIN_KOUTEI(i).MAIN_KOUTEI, "")
'                    Call UniCode_Conv(ITEMREC.AFT_KOUTEI(i).AFT_KOUTEI, "")
'
'                Next i
'
'
'                Call UniCode_Conv(ITEMREC.SE_IO_TANKA_No, "")
'                '��2009.02.20
'
'                Call UniCode_Conv(ITEMREC.STAT, "1")                    '��ԋ敪           2009.01.21
'
'
'
'
'''''''''''''''' 2011.07.05  '''''''''''''''''''''''''''''''''''''''''''''
'
'
'                Call UniCode_Conv(ITEMREC.INSP_MESSAGE, "")             '���iү���� 2009.08.28
'                Call UniCode_Conv(ITEMREC.S_SEIKYU_F, "")               '�����e 2009.08.28
'
'                Call UniCode_Conv(ITEMREC.BEF_S_KOUSU_BAIKA, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_SHIZAI_BAIKA, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_GAISO_TANKA, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_PPSC_KAKO_KOSU, "")
'                Call UniCode_Conv(ITEMREC.BEF_S_BU_KAKO_KOSU, "")
'
'                Call UniCode_Conv(ITEMREC.M_BIKOU, "")
'                Call UniCode_Conv(ITEMREC.SHIYOU_NO, "")
'                Call UniCode_Conv(ITEMREC.MITSUMORI_KBN, "")
'                Call UniCode_Conv(ITEMREC.TANKA_KIRIKAE_DT, "")
'                Call UniCode_Conv(ITEMREC.KIRIKAE_KBN, "")
'
'
'
'
'                Call UniCode_Conv(ITEMREC.KANKYO_KBN, "")
'                Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, "")
'                Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, "")
'
'
'                If Trim(KANKYO_KBN) <> "" Or Trim(KANKYO_KBN_ST) <> "" Or Trim(KANKYO_KBN_SURYO) <> "" Then
'
'                    If Val(KANKYO_KBN_SURYO) <> 0 Then
'
'                        Call UniCode_Conv(ITEMREC.KANKYO_KBN, KANKYO_KBN)
'                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_ST, KANKYO_KBN_ST)
'                        Call UniCode_Conv(ITEMREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)
'
'                    End If
'
'                End If
'
'
'                                                                        '�ǉ��S����
'                '2010.09.01
''                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
'                If Mode = Out_Mode Then
'                    Call UniCode_Conv(ITEMREC.INS_TANTO, "ysyuk")
'                Else
'                    Call UniCode_Conv(ITEMREC.INS_TANTO, "yglcs")
'                End If
'                '2010.09.01
'
'                                                                        '�ǉ�����
'                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
'
'
'
'
'
'                Call UniCode_Conv(ITEMREC.BEF_L_LABEL, "")
'                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
'                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, "")
'                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, "")
'
'
'
'                Call UniCode_Conv(ITEMREC.BIKOU20, "")
'
'                '2011.07.05
'                Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "")
'
'                If NOT_Hin_Name_F Then
'                    For i = 0 To UBound(NOT_Hin_Name)
'                        If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
'                            Call UniCode_Conv(ITEMREC.GOODS_OUT_F, "1")
'                            Exit For
'                        End If
'                    Next i
'                End If
'                '2011.07.05
'
'
'
'''''''''''''''' 2011.07.05  '''''''''''''''''''''''''''''''''''''''''''''
'
'
'
'
'
'
'
'
'
'''''''''''''''' 2011.07.05  ''''''''''''''''''''''''''''''''''''''''''''' delete
''                sBuffer = Space(255)                                    '2009.01.21
''                If GetComputerNameA(sBuffer, 255) <> 0 Then
''                    wkTanto = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
''                Else
''                    wkTanto = "???"
''                End If
''
''
''                                                                        '�ǉ��S����
''                Call UniCode_Conv(ITEMREC.INS_TANTO, wkTanto)
''                                                                        '�ǉ�����
''                Call UniCode_Conv(ITEMREC.Ins_DateTime, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
'''''''''''''''' 2011.07.05  ''''''''''''''''''''''''''''''''''''''''''''' delete
'
'
'
'
'
'
'
'
'
'
'
'                Call UniCode_Conv(ITEMREC.FILLER, "")
'                                                                        '�X�V�S����
'                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
'                                                                        '�X�V����
'                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
'
'
'
'
'
'                Exit Do
            
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �����ւ�    2012.07.05
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If


                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
                DoEvents
                Sleep (500)


            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop
    
    Loop_Cnt = 0
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
                DoEvents
                Sleep (500)
            
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^", 0)
                Exit Function
        End Select
    Loop
        
    If SHIMUKE_Flg Then
        If com = BtOpInsert Then
            '�\���}�X�^�̒ǉ�       2005.12.30
            For i = 0 To UBound(SHIMUKE_T)
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIMUKE_T(i).JGYOBU And _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) = SHIMUKE_T(i).NAIGAI Then
                                                                            '�d�����溰��
                    Call UniCode_Conv(P_COMPO_O_REC.SHIMUKE_CODE, SHIMUKE_T(i).SHIMUKE_CODE)
                                                                            '���ƕ�
                    Call UniCode_Conv(P_COMPO_O_REC.JGYOBU, SHIMUKE_T(i).JGYOBU)
                                                                            '�����O
                    Call UniCode_Conv(P_COMPO_O_REC.NAIGAI, SHIMUKE_T(i).NAIGAI)
                                                                            '�i��
                    Call UniCode_Conv(P_COMPO_O_REC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            '�ް��敪
                    Call UniCode_Conv(P_COMPO_O_REC.DATA_KBN, P_HEAD)
                                                                            '�ǔ�
                    Call UniCode_Conv(P_COMPO_O_REC.SEQNO, "000")
                                                                            '��{�N���X
                    Call UniCode_Conv(P_COMPO_O_REC.CLASS_CODE, "")
                                                                            '���l
                    Call UniCode_Conv(P_COMPO_O_REC.BIKOU, "")
                    
                    Call UniCode_Conv(P_COMPO_O_REC.FILLER, "")
                                                                            '�X�V�S����
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                            '�X�V����
                    Call UniCode_Conv(P_COMPO_O_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                    Loop_Cnt = 0
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                                Beep
'                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
'                                If ans = vbCancel Then
'                                    Exit Function
'                                End If
                            
                            
                                Loop_Cnt = Loop_Cnt + 1
                                If Loop_Cnt > 5 Then
                                    Exit Function
                                End If
                                DoEvents
                                Sleep (500)
                            
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�\���}�X�^", 0)
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If
        
    If HIN_CHANGE <> 0 Then
        LBox_Hin.AddItem JGYOBU & NAIGAI & StrConv(ITEMREC.HIN_GAI, vbUnicode) & BEF_GAI & StrConv(ITEMREC.HIN_NAI, vbUnicode) & BEF_NAI & NAI_CHANGE
    End If

    old_Item_Check_Proc = False

End Function

Sub NG_File_Make_Proc()
'----------------------------------------------------------------------------
'                   �ُ�I���t�@�C���o�͏���
'----------------------------------------------------------------------------
Dim stream  As Integer                       '�t�@�C���ԍ�
Dim Buf     As String                           '�ǂݍ��݃o�b�t�@
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "�ُ�I���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    stream = FreeFile
    Open NG_FILE For Append As stream
    prog = StrConv(App.EXEName, vbUpperCase)
    
    Buf = (Date$ & " " & Time$ & " " & com & " " & prog)
    Print #stream, Buf
    Close stream



End Sub

Sub NG_File_Kill_Proc()
'----------------------------------------------------------------------------
'                   �ُ�I���t�@�C���폜����
'----------------------------------------------------------------------------
Dim stream  As Integer                       '�t�@�C���ԍ�
Dim Buf     As String                           '�ǂݍ��݃o�b�t�@
Dim prog    As String
Dim sBuffer As String * 255
Dim com     As String

Dim NG_FILE As String
Dim c       As String * 128
    
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "NG_FILE", "SYS", c) Then
        Beep
        MsgBox "�ُ�I���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    NG_FILE = RTrim(c)
    
    On Error GoTo Err_Proc
    Kill (NG_FILE)

Err_Proc:

End Sub

'Private Function Y_GLICS_PUT_PROC(JGYOBU As String, NAIGAI As String, INS_NOW As String) As Integer
Private Function Y_GLICS_PUT_PROC(JGYOBU As String, NAIGAI As String, INS_NOW As String, _
                                                                        TEXT_NO As String, _
                                                                        JGYOBU_Code As String, _
                                                                        CYOK_KBN As String, _
                                                                        DEN_DT As String, _
                                                                        IO_KBN As String, _
                                                                        PM_KBN As String, _
                                                                        DEN_SYU As String, _
                                                                        DEN_NO As String, _
                                                                        CYU_KBN As String, _
                                                                        HIN_GAI As String, _
                                                                        HIN_NAI As String, _
                                                                        HIN_NAME As String, _
                                                                        YOTEI_QTY As String, _
                                                                        YOSAN_FROM As String, _
                                                                        YOSAN_TO As String, _
                                                                        HOST_SOKO As String, _
                                                                        HOST_TANA As String, _
                                                                        SYUK_CODE As String, _
                                                                        SYUK_NAME As String, _
                                                                        GENSANKOKU As String, GEN_GENSANKOKU As String, SHIIRE_WORK_CENTER As String, KANKYO_KBN As String, KANKYO_KBN_ST As String, KANKYO_KBN_SURYO As String, ID_NO2 As String, AITESAKI_CODE As String, JYUCHU_YMD As String, SHITEI_NOUKI_YMD As String, MOTO_TEXT_NO As String) As Integer
'----------------------------------------------------------------------------
'           �ƍ��p���ח\��t�@�C���o�͏���
'           2007.06.15
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
        
    
Dim Loop_Cnt        As Integer          '2011.01.15
    
    Y_GLICS_PUT_PROC = True
        
    Call UniCode_Conv(Y_GLICSREC.KAN_KBN, KAN_KBN_FIN)
    Call UniCode_Conv(Y_GLICSREC.DT_SYU, "0")
    Call UniCode_Conv(Y_GLICSREC.JGYOBU, JGYOBU)
    Call UniCode_Conv(Y_GLICSREC.NAIGAI, NAIGAI)
    Call UniCode_Conv(Y_GLICSREC.TEXT_NO, TEXT_NO)


    Call UniCode_Conv(Y_GLICSREC.JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.DATA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.TORI_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KAIKEI_JGYOBA, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_JGYOBA, "")
    
    Call UniCode_Conv(Y_GLICSREC.HIN_NO, HIN_GAI)
    Call UniCode_Conv(Y_GLICSREC.DEN_NO, DEN_NO)
    
    
    '2008.01.10 �}�C�i�X�̑Ή�
    If YOTEI_QTY >= 0 Then
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "0000000"))
    Else
        Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(YOTEI_QTY), "000000"))
    End If
    
    Call UniCode_Conv(Y_GLICSREC.MUKE_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKO_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANKA, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ITEM_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ODER_NO_R, "")
    Call UniCode_Conv(Y_GLICSREC.KOSO_KEITAI, "")
    Call UniCode_Conv(Y_GLICSREC.SYUKA_YMD, DEN_DT)
    Call UniCode_Conv(Y_GLICSREC.TANABAN1, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN2, "")
    Call UniCode_Conv(Y_GLICSREC.TANABAN3, "")
    Call UniCode_Conv(Y_GLICSREC.MUKE_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN, CYU_KBN)
    Call UniCode_Conv(Y_GLICSREC.CYU_KBN_NAME, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN1, "")
    Call UniCode_Conv(Y_GLICSREC.ORIGIN2, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU2, "")
    Call UniCode_Conv(Y_GLICSREC.HAN_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.CHOKU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.UNIT_ID_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_HIKIATE, "")
    Call UniCode_Conv(Y_GLICSREC.GOKON_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_ZAN, "")
    Call UniCode_Conv(Y_GLICSREC.KYOKYU_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SHOHIN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.S_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.BIKOU1, "")
    Call UniCode_Conv(Y_GLICSREC.CHOHA_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.JYU_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAME, HIN_NAME)
    Call UniCode_Conv(Y_GLICSREC.HIN_CHANGE_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.MODULE_EXCHANGE, "")
    Call UniCode_Conv(Y_GLICSREC.ZAIKO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_SHISAN_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.ZAN_HOJYO_SYUSI, "")
    Call UniCode_Conv(Y_GLICSREC.NOUKI_YMD, "")
    Call UniCode_Conv(Y_GLICSREC.SERVICE_KANRI_NO, "")
    Call UniCode_Conv(Y_GLICSREC.KI_HIN_NO, "")
    Call UniCode_Conv(Y_GLICSREC.ENVIRONMENT_KBN, "")
    Call UniCode_Conv(Y_GLICSREC.SS_CODE, "")
    Call UniCode_Conv(Y_GLICSREC.KEPIN_KAIJYO, "")
    
    
    Call UniCode_Conv(Y_GLICSREC.KAN_DT, Format(Now, "YYYYMMDD"))


                        '��s���א��i���׎��ѐ��j
    Call UniCode_Conv(Y_GLICSREC.BEF_NYU_QTY, "00000000")

                        '�\�Z�P�ʌ�
    Call UniCode_Conv(Y_GLICSREC.YOSAN_FROM, YOSAN_FROM)
                        '�\�Z�P�ʐ�
    Call UniCode_Conv(Y_GLICSREC.YOSAN_TO, YOSAN_TO)
                        '�W���I��
    Call UniCode_Conv(Y_GLICSREC.HTANABAN, "")
    Call UniCode_Conv(Y_GLICSREC.HIN_NAI, HIN_NAI)
                        'H�q�� 2006.10.17
    Call UniCode_Conv(Y_GLICSREC.H_SOKO, HOST_SOKO)

                        '���׃��X�g�o�̓t���O   2007.06.12
    Call UniCode_Conv(Y_GLICSREC.NYU_LIST_OUT, " ")
                        '�����敪
    Call UniCode_Conv(Y_GLICSREC.CYOK_KBN, CYOK_KBN)
                        '���o�ɋ敪
    Call UniCode_Conv(Y_GLICSREC.IO_KBN, IO_KBN)
                        '�ԍ��敪
    Call UniCode_Conv(Y_GLICSREC.PM_KBN, PM_KBN)
                        '�`�[���
    Call UniCode_Conv(Y_GLICSREC.DEN_SYU, DEN_SYU)
                        '�x����^�o�א�
    Call UniCode_Conv(Y_GLICSREC.SYUK_CODE, SYUK_CODE)
                        '�x����^�o�א於
    Call UniCode_Conv(Y_GLICSREC.SYUK_NAME, SYUK_NAME)
                        '�}���N����
    Call UniCode_Conv(Y_GLICSREC.INS_NOW, INS_NOW)
    
    
    '----------------   2010.07.08 ��
    Call UniCode_Conv(Y_GLICSREC.GENSANKOKU, GENSANKOKU)                    '���Y����
    Call UniCode_Conv(Y_GLICSREC.GEN_GENSANKOKU, GEN_GENSANKOKU)            '�����\�����Y����
    Call UniCode_Conv(Y_GLICSREC.SHIIRE_WORK_CENTER, SHIIRE_WORK_CENTER)    '���ގd����ܰ�����
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN, KANKYO_KBN)                    '����ދ敪
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_ST, KANKYO_KBN_ST)              '����ދ敪�K�p�J�n
    Call UniCode_Conv(Y_GLICSREC.KANKYO_KBN_SURYO, KANKYO_KBN_SURYO)        '����ދ敪����
    Call UniCode_Conv(Y_GLICSREC.ID_NO2, ID_NO2)                            'ID_NO
    Call UniCode_Conv(Y_GLICSREC.AITESAKI_CODE, AITESAKI_CODE)              '����溰��
    Call UniCode_Conv(Y_GLICSREC.JYUCHU_YMD, JYUCHU_YMD)                    '�󒍔N����
    Call UniCode_Conv(Y_GLICSREC.SHITEI_NOUKI_YMD, SHITEI_NOUKI_YMD)        '�w��[���N����
    Call UniCode_Conv(Y_GLICSREC.LIST_OUT_END_F, "")                        '����ؽďo��F
    Call UniCode_Conv(Y_GLICSREC.NYUKO_TANABAN, "")                         '���ɒI��
    Call UniCode_Conv(Y_GLICSREC.MAEGARI_SURYO, "")                         '�O�ؑ��E��
    '----------------   2010.07.08 ��
    
    
    
    '2011.03.23 �������v���O����
    Call UniCode_Conv(Y_GLICSREC.MOTO_PROG_ID, StrConv(App.EXEName, vbUpperCase))
    '2011.03.23 ���e�L�X�g��
    If Trim(MOTO_TEXT_NO) = "" Then
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, "")
    Else
        Call UniCode_Conv(Y_GLICSREC.MOTO_TEXT_NO, MOTO_TEXT_NO)
    End If
    
    Call UniCode_Conv(Y_GLICSREC.FILLER, "")

    Loop_Cnt = 0

    Do
        sts = BTRV(BtOpInsert, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_GLICSKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
            
            
                Loop_Cnt = Loop_Cnt + 1
                If Loop_Cnt > 5 Then
                    Exit Function
                End If
                DoEvents
                Sleep (500)
            
            Case Else
                Call File_Error(sts, BtOpInsert, "���ח\��", 0)
                Exit Function
        End Select
    Loop

    Y_GLICS_PUT_PROC = False

End Function

'----------------------------------------------------------------------------
'           �ƍ��p���ח\��t�@�C���o�͏���
'           2007.06.15
'----------------------------------------------------------------------------
'Dim sts     As Integer
'Dim ans     As Integer
'
'    Y_GLICS_PUT_PROC = True
'
'    Call UniCode_Conv(Y_GLICSREC.KAN_KBN, KAN_KBN_FIN)
'    Call UniCode_Conv(Y_GLICSREC.DT_SYU, "0")
'    Call UniCode_Conv(Y_GLICSREC.JGYOBU, JGYOBU)
'    Call UniCode_Conv(Y_GLICSREC.NAIGAI, NAIGAI)
'    Call UniCode_Conv(Y_GLICSREC.TEXT_NO, StrConv(New_HS_IN_SIJREC.TEXT_NO, vbUnicode))
'
'
'    Call UniCode_Conv(Y_GLICSREC.JGYOBA, "")
'    Call UniCode_Conv(Y_GLICSREC.DATA_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.TORI_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.ID_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.KAIKEI_JGYOBA, "")
'    Call UniCode_Conv(Y_GLICSREC.SHISAN_JGYOBA, "")
'
'    Call UniCode_Conv(Y_GLICSREC.HIN_NO, StrConv(New_HS_IN_SIJREC.HIN_GAI, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.DEN_NO, StrConv(New_HS_IN_SIJREC.DEN_NO, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.SURYO, Format(CLng(StrConv(New_HS_IN_SIJREC.YOTEI_QTY, vbUnicode)), "0000000"))
'    Call UniCode_Conv(Y_GLICSREC.MUKE_CODE, "")
'    Call UniCode_Conv(Y_GLICSREC.SYUKO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.SHISAN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.HOJYO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.SYUKO_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.TANKA, "")
'    Call UniCode_Conv(Y_GLICSREC.ODER_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ITEM_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ODER_NO_R, "")
'    Call UniCode_Conv(Y_GLICSREC.KOSO_KEITAI, "")
'    Call UniCode_Conv(Y_GLICSREC.SYUKA_YMD, StrConv(New_HS_IN_SIJREC.DEN_DT, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.TANABAN1, "")
'    Call UniCode_Conv(Y_GLICSREC.TANABAN2, "")
'    Call UniCode_Conv(Y_GLICSREC.TANABAN3, "")
'    Call UniCode_Conv(Y_GLICSREC.MUKE_NAME, "")
'    Call UniCode_Conv(Y_GLICSREC.CYU_KBN, StrConv(New_HS_IN_SIJREC.CYU_KBN, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.CYU_KBN_NAME, "")
'    Call UniCode_Conv(Y_GLICSREC.ORIGIN1, "")
'    Call UniCode_Conv(Y_GLICSREC.ORIGIN2, "")
'    Call UniCode_Conv(Y_GLICSREC.BIKOU2, "")
'    Call UniCode_Conv(Y_GLICSREC.HAN_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.CHOKU_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.UNIT_ID_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAIKO_HIKIATE, "")
'    Call UniCode_Conv(Y_GLICSREC.GOKON_KANRI_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.JYUCHU_ZAN, "")
'    Call UniCode_Conv(Y_GLICSREC.KYOKYU_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.SHOHIN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.S_SHISAN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.S_HOJYO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.BIKOU1, "")
'    Call UniCode_Conv(Y_GLICSREC.CHOHA_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.JYU_HIN_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.HIN_NAME, StrConv(New_HS_IN_SIJREC.HIN_NAME, vbUnicode))
'    Call UniCode_Conv(Y_GLICSREC.HIN_CHANGE_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.MODULE_EXCHANGE, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAIKO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAN_SHISAN_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.ZAN_HOJYO_SYUSI, "")
'    Call UniCode_Conv(Y_GLICSREC.NOUKI_YMD, "")
'    Call UniCode_Conv(Y_GLICSREC.SERVICE_KANRI_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.KI_HIN_NO, "")
'    Call UniCode_Conv(Y_GLICSREC.ENVIRONMENT_KBN, "")
'    Call UniCode_Conv(Y_GLICSREC.SS_CODE, "")
'    Call UniCode_Conv(Y_GLICSREC.KEPIN_KAIJYO, "")
'
'
'    Call UniCode_Conv(Y_GLICSREC.KAN_DT, Format(Now, "YYYYMMDD"))
'
'
'                        '��s���א��i���׎��ѐ��j
'    Call UniCode_Conv(Y_GLICSREC.BEF_NYU_QTY, "00000000")
'
'                        '�\�Z�P�ʌ�
'    Call UniCode_Conv(Y_GLICSREC.YOSAN_FROM, StrConv(New_HS_IN_SIJREC.YOSAN_FROM, vbUnicode))
'                        '�\�Z�P�ʐ�
'    Call UniCode_Conv(Y_GLICSREC.YOSAN_TO, StrConv(New_HS_IN_SIJREC.YOSAN_TO, vbUnicode))
'                        '�W���I��
'    Call UniCode_Conv(Y_GLICSREC.HTANABAN, "")
'    Call UniCode_Conv(Y_GLICSREC.HIN_NAI, StrConv(New_HS_IN_SIJREC.HIN_NAI, vbUnicode))
'                        'H�q�� 2006.10.17
'    Call UniCode_Conv(Y_GLICSREC.H_SOKO, StrConv(New_HS_IN_SIJREC.HOST_SOKO, vbUnicode))

'                        '���׃��X�g�o�̓t���O   2007.06.12
'    Call UniCode_Conv(Y_GLICSREC.NYU_LIST_OUT, " ")
'                        '�����敪
'    Call UniCode_Conv(Y_GLICSREC.CYOK_KBN, StrConv(New_HS_IN_SIJREC.CYOK_KBN, vbUnicode))
'                        '���o�ɋ敪
'    Call UniCode_Conv(Y_GLICSREC.IO_KBN, StrConv(New_HS_IN_SIJREC.IO_KBN, vbUnicode))
'                        '�ԍ��敪
'    Call UniCode_Conv(Y_GLICSREC.PM_KBN, StrConv(New_HS_IN_SIJREC.PM_KBN, vbUnicode))
'                        '�`�[���
'    Call UniCode_Conv(Y_GLICSREC.DEN_SYU, StrConv(New_HS_IN_SIJREC.DEN_SYU, vbUnicode))
'                        '�x����^�o�א�
'    Call UniCode_Conv(Y_GLICSREC.SYUK_CODE, StrConv(New_HS_IN_SIJREC.SYUK_CODE, vbUnicode))
'                        '�x����^�o�א於
'    Call UniCode_Conv(Y_GLICSREC.SYUK_NAME, StrConv(New_HS_IN_SIJREC.SYUK_NAME, vbUnicode))
'                        '�}���N����
'    Call UniCode_Conv(Y_GLICSREC.INS_NOW, INS_NOW)
'
'
'    Call UniCode_Conv(Y_GLICSREC.FILLER, "")
'
'    Do
'        sts = BTRV(BtOpInsert, Y_GLICS_POS, Y_GLICSREC, Len(Y_GLICSREC), K0_Y_GLICS, Len(K0_Y_GLICS), 0)
'        Select Case sts
'            Case BtNoErr
'                Exit Do
'            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
'                Beep
'                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_GLICSKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                If ans = vbCancel Then
'                    Exit Function
'                End If
'            Case Else
'                Call File_Error(sts, BtOpInsert, "���ח\��")
'                Exit Function
'        End Select
'    Loop
'
'    Y_GLICS_PUT_PROC = False
'
'
'End Function


