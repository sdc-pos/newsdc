VERSION 5.00
Begin VB.Form F1020201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���o�ח\��f�[�^�捞�� 2009.10.05 10:30"
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
Attribute VB_Name = "F1020201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WS_NO       As String * 2           'ܰ��ð��ݔԍ�


Private Type SHIMUKE_TBL
    SHIMUKE_CODE            As String * 2   '�d������
    JGYOBU                  As String * 1   '���ƕ�
    NAIGAI                  As String * 1   '�����O
End Type

Private SHIMUKE_T()         As SHIMUKE_TBL

Private SHIMUKE_Flg         As Boolean


Private Const In_Mode% = 1                  '���׏���
Private Const Out_Mode% = 2                 '�o�׏���


'''Private INS_DATE            As String * 8   '���s���t
'''Private INS_BIN             As Integer      '��

                                            
                                        
    
Private Function Syuka_Update_Proc(JGYOBU As String, ix As Integer) As Boolean
'----------------------------------------------------------------------------
'                   �u�o�ח\��f�[�^�v�X�V����
'----------------------------------------------------------------------------
Dim In_Cnt          As Integer              '�f�[�^�ǂݍ��݌���
Dim Out_Cnt         As Integer              '�f�[�^�o�͌���



Dim INS_NOW         As String


Dim sts             As Integer
Dim Ret             As String

Dim DUP_SYUKANo     As Long

Dim HS_SMEISAINo    As Long
Dim HS_SMEISAI_OP   As Boolean

Dim HS_PICNo        As Long
Dim HS_PIC_OP       As Boolean

Dim fileName        As String

Dim c               As String * 128

Dim i               As Integer

Dim Input_Buffer    As String
Dim Pos             As Integer
        
Dim Skip_Flg        As Boolean
Dim Fast_Flg        As Boolean

Dim Input_Wk        As Variant

Dim LOCATION        As String
Dim HIN_NAME        As String

Dim SYUKA_NO        As String
Dim SYUKA_YMD       As String
Dim OKURISAKI       As String
Dim URIDEN          As String
Dim DEN_NO          As String
Dim HINBAN          As String
Dim SURYO           As String
Dim CYU_NO          As String
Dim TOKUI_CODE      As String
Dim TOKUI_NAME      As String
Dim BIKOU           As String
Dim UNSOU           As String
Dim INS_BIN         As String               '2007.01.16

Dim SV_DEN_NO       As String * 7
Dim SV_OKURISAKI    As String
Dim SV_TOKUI_CD     As String * 8

Dim SV_URIDEN       As String * 1           '2007.01.08



Dim DEN_SEQ         As Integer

Dim ID_SET_FLG      As Boolean
Dim SV_ID_NO        As String * 7
Dim ID_SEQ          As Integer



Dim ans             As Integer


    Syuka_Update_Proc = False

    '�o�ז��׃t�@�C������荞�� & �n�o�d�m
    If GetIni("FILE", "HS_SMEISAI", "SYS", c) Then
        Beep
        MsgBox "�o�ז��׃t�@�C���E�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Exit Function
    End If
    fileName = Trim(c)

    HS_SMEISAI_OP = False

    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Exit_Proc
    
    HS_SMEISAINo = FreeFile
    Open fileName For Input As #HS_SMEISAINo

    On Error GoTo Exit_Proc            '�����I��
    HS_SMEISAI_OP = True
    
    '�s�b�L���O�t�@�C������荞�� & �n�o�d�m
    If GetIni("FILE", "HS_PIC", "SYS", c) Then
        Beep
        MsgBox "�s�b�L���O���X�g�E�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Exit Function
    End If
    fileName = Trim(c)

    HS_PIC_OP = False
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Exit_Proc
    
    HS_PICNo = FreeFile
    Open fileName For Input As #HS_PICNo

    On Error GoTo Exit_Proc            '�����I��
    HS_PIC_OP = True
    
    Syuka_Update_Proc = True
    
    
    '�o�׏d���t�@�C������荞�� & �n�o�d�m
    
    If GetIni("FILE", "SYUDUP  ", "SYS", c) Then
        Beep
        MsgBox "�o�׏d���t�@�C���E�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Exit Function
    End If
    fileName = Trim(c)
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & Format(Now, "YYYYMMDDHHMMSS") & "_" & JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)




    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")

    '-----------------------------------------------------------------  �s�b�L���O���X�g���i�ڃ}�X�^�쐬�^�X�V
    Do While Not EOF(HS_PICNo)
        
        DoEvents
        
        Line Input #HS_PICNo, Input_Buffer
        
        Input_Wk = Split(Input_Buffer, vbTab, -1)
            
        LOCATION = ""
        HINBAN = ""
        HIN_NAME = ""
    
    
        If UBound(Input_Wk) > 6 Then
            LOCATION = StrConv(Input_Wk(1), vbNarrow)
            HINBAN = Input_Wk(3)
            HIN_NAME = Input_Wk(7)
        End If
    
        If Trim(HINBAN) = "" Or _
            Trim(HIN_NAME) = "" Then
        Else
                        '�i�ڃ}�X�^�`�F�b�N
            If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HINBAN, HIN_NAME, LOCATION) Then
                Exit Function
            End If
        
        
        End If
    Loop




    In_Cnt = 0
    Out_Cnt = 0

    



    SV_DEN_NO = ""
    SV_OKURISAKI = ""


    SV_ID_NO = ""



    Fast_Flg = True


    Do While Not EOF(HS_SMEISAINo)
        
        In_Cnt = In_Cnt + 1
        lblINCNT(ix).Caption = Format(In_Cnt, "#0")
        DoEvents
        
        
        
        Line Input #HS_SMEISAINo, Input_Buffer




        Input_Wk = Split(Input_Buffer, vbTab, -1)

        SYUKA_NO = ""
        SYUKA_YMD = ""
        OKURISAKI = ""
        URIDEN = ""
        DEN_NO = ""
        HINBAN = ""
        SURYO = ""
        CYU_NO = ""
        TOKUI_CODE = ""
        TOKUI_NAME = ""
        BIKOU = ""
        UNSOU = ""
        INS_BIN = ""
        
        
        '�o�ׇ�
        If UBound(Input_Wk) > 0 Then
            SYUKA_NO = Input_Wk(1)
        End If
        
        
If SYUKA_NO = "20" Then
    Debug.Print
End If
        
        If Not IsNumeric(SYUKA_NO) Then
        Else
            '�o�ד�
            If UBound(Input_Wk) > 1 Then
                
                If Mid(Format(Now, "YYYYMMDD"), 5, 2) = "12" Then
                    If Mid(CStr(Input_Wk(2)), 1, 2) = "01" Then
                        SYUKA_YMD = Format(CLng(Mid(Format(Now, "YYYYMMDD"), 1, 4) + 1), "0000") & "/" & Input_Wk(2)
                    Else
                        SYUKA_YMD = Mid(Format(Now, "YYYYMMDD"), 1, 4) & "/" & Input_Wk(2)
                    End If
                Else
                    SYUKA_YMD = Mid(Format(Now, "YYYYMMDD"), 1, 4) & "/" & Input_Wk(2)
                End If
            End If
        
            '����於
            If UBound(Input_Wk) > 3 Then
                
                ID_SET_FLG = False
                If Trim(Input_Wk(4)) <> "" Then
                    
                    If SV_OKURISAKI <> Trim(Input_Wk(4)) Then
                    
                        SV_OKURISAKI = Input_Wk(4)
                        ID_SET_FLG = True
                    
                    
                        SV_TOKUI_CD = "********"
                        SV_URIDEN = "*"             '2007.01.08
                    
                    
                    End If
                End If
                
                If Len(Input_Wk(7)) > 7 Then
                Else
                    If UBound(Input_Wk) > 11 Then
                    
                        If SV_TOKUI_CD = "********" Then
                            
                            SV_TOKUI_CD = Input_Wk(13)
                        
                            SV_URIDEN = Input_Wk(5) '2007.01.08
                        
                        End If
                    
                        If Trim(SV_TOKUI_CD) <> Input_Wk(13) Then
                            ID_SET_FLG = True
                            SV_TOKUI_CD = Input_Wk(13)
                        
                            SV_URIDEN = Input_Wk(5) '2007.01.08
                        
                        
                        End If
                    
                        If Trim(SV_URIDEN) <> Input_Wk(5) Then      '2007.01.08
                            If Left(SV_ID_NO, 2) = "06" Then        '2007.01.08
                            Else                                    '2007.01.08
                                ID_SET_FLG = True                   '2007.01.08
                                SV_URIDEN = Input_Wk(5)             '2007.01.08
                            End If                                  '2007.01.08
                        End If                                      '2007.01.08
                    
                    End If
                End If
                
                OKURISAKI = SV_OKURISAKI
            
            End If
            
            '���`
            If UBound(Input_Wk) > 4 Then
                URIDEN = Input_Wk(5)
            End If
            '�`�[�ԍ�
            If UBound(Input_Wk) > 6 Then
                
                If Len(Input_Wk(7)) > 7 Then
                
                    DEN_NO = Left(Input_Wk(7), 7)
                Else
                    DEN_NO = Input_Wk(7)
                End If
            
                If ID_SET_FLG Then
                    SV_ID_NO = DEN_NO
                    ID_SEQ = 0
                End If
            
            End If
            '�i��
            If UBound(Input_Wk) > 8 Then
                HINBAN = Input_Wk(9)
            End If
            '����
            If UBound(Input_Wk) > 9 Then
                SURYO = Input_Wk(10)
            End If
            '������
            If UBound(Input_Wk) > 11 Then
                CYU_NO = Input_Wk(12)
            End If
            '���Ӑ溰��
            If UBound(Input_Wk) > 12 Then
                TOKUI_CODE = Input_Wk(13)
            End If
            '���Ӑ於
            If UBound(Input_Wk) > 13 Then
                TOKUI_NAME = Input_Wk(14)
            End If
            '���l
            If UBound(Input_Wk) > 15 Then
                BIKOU = Input_Wk(16)
            End If
            '�^�����
            If UBound(Input_Wk) > 17 Then
                UNSOU = Input_Wk(18)
            End If
            '�� '2007.01.16
            If UBound(Input_Wk) > 18 Then
                INS_BIN = Input_Wk(19)
            End If
            
            
            
            '�װ����
            Skip_Flg = False
            
            If Trim(SYUKA_YMD) = "" Or _
                Trim(DEN_NO) = "" Or _
                Trim(HINBAN) = "" Or _
                Trim(SURYO) = "" Then
'''                Trim(TOKUI_CODE) = "" Then
                
                Skip_Flg = True
        
            Else
        
                If Not IsDate(SYUKA_YMD) Then
                    Skip_Flg = True
                Else
                    SYUKA_YMD = (Format(SYUKA_YMD, "YYYYMMDD"))
                End If
        
                If Not IsNumeric(SURYO) Then
                    Skip_Flg = True
                Else
                    If CLng(SURYO) = 0 Then
                        Skip_Flg = True
                    End If
                End If
        
        
            End If
        
            If Not Skip_Flg Then
                
                
                If Trim(SV_DEN_NO) = "" Then
                    SV_DEN_NO = DEN_NO
                    DEN_SEQ = 0
                End If
        
                If SV_DEN_NO <> DEN_NO Then
                    SV_DEN_NO = DEN_NO
                    DEN_SEQ = 0
                End If
                
                DEN_SEQ = DEN_SEQ + 1
                ID_SEQ = ID_SEQ + 1
        
        
                '�o�ח\��d������
                Call UniCode_Conv(K0_Y_SYU.JGYOBU, JGYOBU)
                Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, SV_ID_NO & Format(ID_SEQ, "00"))
        
                sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Skip_Flg = True
                
                
                        If Fast_Flg Then
                            DUP_SYUKANo = FreeFile
                            Open fileName For Append As #DUP_SYUKANo

                            Write #DUP_SYUKANo, , , "�o�׏d�����X�g", , "�쐬���F", Format(Now, "YYYY/MM/DD HH:MM:SS")
                            Write #DUP_SYUKANo, "��", "�o�ד�", "����於", "���`", "�`�[�ԍ�", "�i��", "����", "������", "���Ӑ�CD", "���Ӑ於", "���l", "�^�����"
                            Fast_Flg = False
                        
                        End If
                
                
                        Write #DUP_SYUKANo, SYUKA_NO,
                        Write #DUP_SYUKANo, SYUKA_YMD,
                        Write #DUP_SYUKANo, OKURISAKI,
                        Write #DUP_SYUKANo, URIDEN,
                        Write #DUP_SYUKANo, DEN_NO,
                        Write #DUP_SYUKANo, HINBAN,
                        Write #DUP_SYUKANo, SURYO,
                        Write #DUP_SYUKANo, CYU_NO,
                        Write #DUP_SYUKANo, TOKUI_CODE,
                        Write #DUP_SYUKANo, TOKUI_NAME,
                        Write #DUP_SYUKANo, BIKOU,
                        Write #DUP_SYUKANo, UNSOU

                
                
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�o�ח\��")
                        Exit Function
                End Select
        
        
                If Not Skip_Flg Then
                
                
                    '��ݻ޸��݊J�n
                    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
                        Exit Function
                    End If
                    '---------------------------------------------------------- ���Ӑ������
                    Call UniCode_Conv(K0_MTS.MUKE_CODE, TOKUI_CODE)
                    Call UniCode_Conv(K0_MTS.SS_CODE, "")
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            '���o�^�͎����쐬
                            Call UniCode_Conv(MTSREC.NAIGAI, NAIGAI_NAI)
                            Call UniCode_Conv(MTSREC.DATA_KBN, "")
                            Call UniCode_Conv(MTSREC.MUKE_CODE, TOKUI_CODE)
                            Call UniCode_Conv(MTSREC.SS_CODE, "")
                            Call UniCode_Conv(MTSREC.MUKE_NAME, TOKUI_NAME)
                            Call UniCode_Conv(MTSREC.SS_NAME, "")
                            Call UniCode_Conv(MTSREC.MUKE_DNAME, TOKUI_NAME)
                            Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "")
                            Call UniCode_Conv(MTSREC.FILLER, "")
                            Do
                                sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                        Beep
                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                        If ans = vbCancel Then
                                            GoTo Abort_Tran
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpInsert, "������Ǘ�Ͻ�" & "key=" & TOKUI_CODE)
                                        GoTo Abort_Tran
                                End Select
                            Loop
                                                        
                                                        
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                            GoTo Abort_Tran
                    End Select
                
                    '---------------------------------------------------------- �i��Ͻ�������
                    If Item_Check_Proc(Out_Mode, JGYOBU, NAIGAI_NAI, HINBAN) Then
                        GoTo Abort_Tran
                    End If
                
                
                    '---------------------------------------------------------- �o�ח\��쐬
                
                
                    '�g�p�[��ID
                    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                    '�g�p����۸���ID
                    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                    '�����敪
                    Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_UN)
                    '�ް����
                    Call UniCode_Conv(Y_SYUREC.DT_SYU, "0")
                    '���ƕ�
                    Call UniCode_Conv(Y_SYUREC.JGYOBU, JGYOBU)
                    '�����敪(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_CYU_KBN, CYU_KBN_SPO)
                    'ID-NO
                    Call UniCode_Conv(Y_SYUREC.KEY_ID_NO, SV_ID_NO & Format(ID_SEQ, "00"))
                    '�����O
                    Call UniCode_Conv(Y_SYUREC.NAIGAI, NAIGAI_NAI)
                    '�i��(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_HIN_NO, HINBAN)
                    '���Ӑ�(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_MUKE_CODE, TOKUI_CODE)
                    '������(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_SS_CODE, "")
                    '�o�ד�(KEY)
                    Call UniCode_Conv(Y_SYUREC.KEY_SYUKA_YMD, SYUKA_YMD)
                    '���Ə꺰��
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                    '�ް��敪
                    Call UniCode_Conv(Y_SYUREC.DATA_KBN, "")
                    '����敪
                    Call UniCode_Conv(Y_SYUREC.TORI_KBN, "")
                    'ID-NO
                    Call UniCode_Conv(Y_SYUREC.ID_NO, SV_ID_NO & Format(ID_SEQ, "00"))
                    '��v�p���Ə꺰��
                    Call UniCode_Conv(Y_SYUREC.JGYOBA, "")
                    '�i��
                    Call UniCode_Conv(Y_SYUREC.HIN_NO, HINBAN)
                    '�`�[�ԍ�
                    Call UniCode_Conv(Y_SYUREC.DEN_NO, SV_DEN_NO)
                    '�o�א���
                    Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(SURYO), "0000000"))
                    '���Ӑ�
                    Call UniCode_Conv(Y_SYUREC.MUKE_CODE, TOKUI_CODE)
                    '�o�Ɏ��x
                    Call UniCode_Conv(Y_SYUREC.SYUKO_SYUSI, "")
                    '���Y�Ǘ��p�݌Ɏ��x����
                    Call UniCode_Conv(Y_SYUREC.SHISAN_SYUSI, "")
                    '�⏕�݌Ɏ��x����
                    Call UniCode_Conv(Y_SYUREC.HOJYO_SYUSI, "")
                    '�o�ד��t
                    Call UniCode_Conv(Y_SYUREC.SYUKO_YMD, SYUKA_YMD)
                    '���ےP��
                    Call UniCode_Conv(Y_SYUREC.TANKA, "")
                    '���ް�ԍ�
                    Call UniCode_Conv(Y_SYUREC.ODER_NO, "")
                    '���єԍ�
                    Call UniCode_Conv(Y_SYUREC.ITEM_NO, "")
                    '�����Ǘ��ԍ�����
                    Call UniCode_Conv(Y_SYUREC.ODER_NO_R, "")
                    '���`�Ժ���
                    Call UniCode_Conv(Y_SYUREC.KOSO_KEITAI, "")
                    '�o�ח\���
                    Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, SYUKA_YMD)
                    '۹����1
                    Call UniCode_Conv(Y_SYUREC.TANABAN1, "")
                    '۹����2
                    Call UniCode_Conv(Y_SYUREC.TANABAN2, "")
                    '۹����3
                    Call UniCode_Conv(Y_SYUREC.TANABAN3, "")
                    '���Ӑ於��
                    Call UniCode_Conv(Y_SYUREC.MUKE_NAME, TOKUI_NAME)
                    '�����敪
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN, CYU_KBN_SPO)
                    '�����敪����
                    Call UniCode_Conv(Y_SYUREC.CYU_KBN_NAME, CYU_KBN_2)
                    '���Y��1
                    Call UniCode_Conv(Y_SYUREC.ORIGIN1, "")
                    '���Y��2
                    Call UniCode_Conv(Y_SYUREC.ORIGIN2, "")
                    '���l2
                    Call UniCode_Conv(Y_SYUREC.BIKOU2, "")
                    '�̔��敪
                    Call UniCode_Conv(Y_SYUREC.HAN_KBN, "")
                    '�����w���敪
                    Call UniCode_Conv(Y_SYUREC.CHOKU_KBN, "")
                    '�ƯďC���Ǘ��ԍ�
                    Call UniCode_Conv(Y_SYUREC.UNIT_ID_NO, "")
                    '�݌Ɉ�������
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_HIKIATE, "")
                    '�����Ǘ��ԍ�
                    Call UniCode_Conv(Y_SYUREC.GOKON_KANRI_NO, "")
                    '�󒍎c����
                    Call UniCode_Conv(Y_SYUREC.JYUCHU_ZAN, "")
                    '�����敪
                    Call UniCode_Conv(Y_SYUREC.KYOKYU_KBN, "")
                    '���i���[�i�݌Ɏ��x����
                    Call UniCode_Conv(Y_SYUREC.SHOHIN_SYUSI, "")
                    '���i���[�i���Y�Ǘ����x����
                    Call UniCode_Conv(Y_SYUREC.S_SHISAN_SYUSI, "")
                    '���i���[�i�⏕���x����
                    Call UniCode_Conv(Y_SYUREC.S_HOJYO_SYUSI, "")
                    '���l1
                    Call UniCode_Conv(Y_SYUREC.BIKOU1, "")
                    '���[�敪
                    Call UniCode_Conv(Y_SYUREC.CHOHA_KBN, "")
                    '��t�i�ڔԍ�
                    Call UniCode_Conv(Y_SYUREC.JYU_HIN_NO, "")
                    '�i��
                    Call UniCode_Conv(Y_SYUREC.HIN_NAME, StrConv(ITEMREC.HIN_NAME, vbUnicode))
                    '�i�ڔԍ��ύX�敪
                    Call UniCode_Conv(Y_SYUREC.HIN_CHANGE_KBN, "")
                    'Ӽޭ�ٌ����敪
                    Call UniCode_Conv(Y_SYUREC.MODULE_EXCHANGE, "")
                    '�c�݌ɂ܂Ƃߍ݌Ɏ��x����
                    Call UniCode_Conv(Y_SYUREC.ZAIKO_SYUSI, "")
                    '�c�݌ɂ܂Ƃߎ��Y�Ǘ����x����
                    Call UniCode_Conv(Y_SYUREC.ZAN_SHISAN_SYUSI, "")
                    '�w��[��
                    Call UniCode_Conv(Y_SYUREC.NOUKI_YMD, "")
                    '���޽��ЊǗ��ԍ�
                    Call UniCode_Conv(Y_SYUREC.SERVICE_KANRI_NO, "")
                    '�@��i�ں���
                    Call UniCode_Conv(Y_SYUREC.KISHU_CODE, "")
                    '����敔�i�敪
                    Call UniCode_Conv(Y_SYUREC.ENVIRONMENT_KBN, "")
                    '��������溰��
                    Call UniCode_Conv(Y_SYUREC.SS_CODE, "")
                    '���i�����敪
                    Call UniCode_Conv(Y_SYUREC.KEPIN_KAIJYO, "")
                    '�i�ԁi�����j
                    Call UniCode_Conv(Y_SYUREC.HIN_NAI, "")
                    '�W���I��
                    Call UniCode_Conv(Y_SYUREC.HTANABAN, StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode))
                    '�o�ɕ\������t
                    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, "")
                    '�������t
                    Call UniCode_Conv(Y_SYUREC.KAN_YMD, "")
                    '���i���t
                    Call UniCode_Conv(Y_SYUREC.KENPIN_YMD, "")
                    '������敪
                    Call UniCode_Conv(Y_SYUREC.TOK_KBN, "")
                    '�o�Ɏ��ѐ���
                    Call UniCode_Conv(Y_SYUREC.JITU_SURYO, "00000000")
                    '�捞�ݓ���
                    Call UniCode_Conv(Y_SYUREC.INS_NOW, INS_NOW)
                    '���i�S���Һ���
                    Call UniCode_Conv(Y_SYUREC.KENPIN_TANTO_CODE, "")
                    '���i����
                    Call UniCode_Conv(Y_SYUREC.KENPIN_HMS, "")
                    '����ݸ�p������
                    Call UniCode_Conv(Y_SYUREC.LK_MUKE_CODE, TOKUI_CODE)
                    '����ݸ�p�A��
                    Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, "")
                    '��ʌ��i�׸�
                    Call UniCode_Conv(Y_SYUREC.G_KENPIN_F, "")
                    '���i������
                    Call UniCode_Conv(Y_SYUREC.KENPIN_SURYO, "")
                    
                    
                    '
                    Call UniCode_Conv(Y_SYUREC.FILLER, "")


    
                    Do
                        sts = BTRV(BtOpInsert, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                     GoTo Abort_Tran
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�o�ח\��")
                                GoTo Abort_Tran
                        End Select
                    Loop
    
    
    
                    '---------------------------------------------------------- �o�ח\��(νĲҰ��)�쐬
                    'ID-NO
                    Call UniCode_Conv(Y_SYU_HREC.ID_NO, StrConv(Y_SYUREC.ID_NO, vbUnicode))
                    '��
                    Call UniCode_Conv(Y_SYU_HREC.SYUKA_NO, SYUKA_NO)
                    '�o�ד��t
                    Call UniCode_Conv(Y_SYU_HREC.SYUKA_YMD, SYUKA_YMD)
                    '����於
                    Call UniCode_Conv(Y_SYU_HREC.OKURISAKI, OKURISAKI)
                    '���`
                    If Trim(URIDEN) = "" Then
                        Call UniCode_Conv(Y_SYU_HREC.URIDEN, "0")
                    Else
                        Call UniCode_Conv(Y_SYU_HREC.URIDEN, "1")
                    End If
                    '�`�[�ԍ�
                    Call UniCode_Conv(Y_SYU_HREC.DEN_NO, SV_DEN_NO)
                    '�ǔ�
                    Call UniCode_Conv(Y_SYU_HREC.SEQ_NO, Format(DEN_SEQ, "0"))
                    '�i��
                    Call UniCode_Conv(Y_SYU_HREC.HIN_NO, HINBAN)
                    '����
                    Call UniCode_Conv(Y_SYU_HREC.SURYO, Format(CLng(SURYO), "0000000"))
                    '������
                    Call UniCode_Conv(Y_SYU_HREC.ODER_NO, CYU_NO)
                    '���Ӑ�
                    Call UniCode_Conv(Y_SYU_HREC.MUKE_CODE, TOKUI_CODE)
                    '���Ӑ於��
                    Call UniCode_Conv(Y_SYU_HREC.MUKE_NAME, TOKUI_NAME)
                    '���l
                    Call UniCode_Conv(Y_SYU_HREC.BIKOU, BIKOU)
                    '�^����Ж�
                    Call UniCode_Conv(Y_SYU_HREC.UNSOU_KAISHA, UNSOU)
                    '�捞�ݓ���
                    Call UniCode_Conv(Y_SYU_HREC.INS_NOW, INS_NOW)
                    '�o�����و������
                    Call UniCode_Conv(Y_SYU_HREC.PRINT_NOW, "")
                    '�ް�������
                    Call UniCode_Conv(Y_SYU_HREC.DATA_CNT, Format(Out_Cnt, "00000"))
                    '�����
                    Call UniCode_Conv(Y_SYU_HREC.OKURI_NO, "")
                    '���i����
                    Call UniCode_Conv(Y_SYU_HREC.KENPIN_NOW, "")
                    '���i�S����
                    Call UniCode_Conv(Y_SYU_HREC.KENPIN_TANTO_CODE, "")
                    '����
                    Call UniCode_Conv(Y_SYU_HREC.KUTI_SU, "0000")   '2007.02.01
                    Call UniCode_Conv(Y_SYU_HREC.xKUTI_SU, "00")    '2007.02.01
                    
                    '��������
                    Call UniCode_Conv(Y_SYU_HREC.KYOSEI_END, "")
                    '��ݾ�F
                    Call UniCode_Conv(Y_SYU_HREC.CANCEL_F, "")
                    '���l
                    Call UniCode_Conv(Y_SYU_HREC.INPUT_BIKOU, "")
                    '�� 2007.01.16
                    If IsNumeric(INS_BIN) Then
                        Call UniCode_Conv(Y_SYU_HREC.INS_BIN, Format(CInt(INS_BIN), "00"))
                    Else
                        Call UniCode_Conv(Y_SYU_HREC.INS_BIN, "")
                    End If
                    
                    Call UniCode_Conv(Y_SYU_HREC.FILLER, "")
                    
                    Do
                        sts = BTRV(BtOpInsert, Y_SYU_H_POS, Y_SYU_HREC, Len(Y_SYU_HREC), K0_Y_SYU_H, Len(K0_Y_SYU_H), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_H.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                     GoTo Abort_Tran
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�o�ח\��(νĲҰ��)")
                                GoTo Abort_Tran
                        End Select
                    Loop
                    
                    
                    
                    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    If sts <> BtNoErr Then
                        GoTo Abort_Tran
                    End If
                
                
                
                
                    Out_Cnt = Out_Cnt + 1
                    lblOUTCNT(ix).Caption = Format(Out_Cnt, "#0")
                
                
                
                
                
                
                End If
        
        
            End If
        
        
        End If




    Loop















    Close #HS_SMEISAINo
    Close #HS_PICNo
    
    Syuka_Update_Proc = False

    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
    Exit Function
    
Exit_Proc:
    
    If HS_SMEISAI_OP Then
        Close #HS_SMEISAINo
    End If
    
    If HS_PIC_OP Then
        Close #HS_PICNo
    End If
    
    
End Function

Private Sub Form_Activate()

Dim Ret         As String


Dim i           As Integer
Dim FullPath    As String


    '---------------------------------------------  ���ƕ������C�����[�v
    For i = 0 To UBound(JGYOBU_T)
        

        lblJGYOBU(i).Caption = JGYOBU_T(i).NAME
        lblJGYOBU(i).ForeColor = JGYOBU_T(i).COLOR

        lblOUTCNT(i).Caption = "0"
        lblINCNT(i).Caption = "0"
        DoEvents

    
    
        If Syuka_Update_Proc(JGYOBU_T(i).CODE, i) Then '�o�ח\��f�[�^�X�V����

            Unload Me
        End If
    
    
    
    
    
    
    Next i


    Unload Me

Error_Proc:

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
                               
    If JGYOB_TB_Set(1) Then     '���ƕ��̊l��
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)

                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m '2005.12.30
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m   2005.12.30
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��(νĲҰ��)�n�o�d�m
    If Y_SYU_H_Open(BtOpenNomal) Then
        Unload Me
    End If


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
                Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                Unload Me
        End Select
    
        com = BtOpGetNext
    Loop
        
                                '�ւ̊l��       '2007.01.16
'''    If GetIni(App.EXEName, "INS_DATE", App.EXEName, c) Then
'''        INS_DATE = Format(Now, "YYYYMMDD")
'''        INS_BIN = 1
'''    Else
'''        If Trim(c) <> Format(Now, "YYYYMMDD") Then
'''            INS_DATE = Format(Now, "YYYYMMDD")
'''            INS_BIN = 1
'''        Else
'''            INS_DATE = Trim(c)
'''
'''            If GetIni(App.EXEName, "INS_BIN", App.EXEName, c) Then
'''                INS_BIN = 1
'''            Else
'''                If IsNumeric(Trim(c)) Then
'''                    INS_BIN = CInt(Trim(c)) + 1
'''                Else
'''                    INS_BIN = 1
'''                End If
'''            End If
'''        End If
'''    End If
'''
'''                                '�h�m�h �{�����t�o��
'''    If WriteIni(App.EXEName, "INS_DATE", App.EXEName, INS_DATE) Then
'''    End If
'''                                '�h�m�h �֏o��
'''    If WriteIni(App.EXEName, "INS_BIN", App.EXEName, Format(INS_BIN, "0")) Then
'''    End If
    
    
    
    
    
    


    Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

    DoEvents
    
'    If Last_Proc_F = True Then              '���������ް��폜�����@���s�L��H
'        Call Last_Proc
'    End If

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
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
                                            '�o�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��")
        End If
    End If
                                            '�a���������������Z�b�g
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1020201 = Nothing

    End
End Sub

Private Function Item_Check_Proc(Mode As Integer, _
                                    JGYOBU As String, _
                                    NAIGAI As String, _
                                    HIN_GAI As String, _
                                    Optional HIN_NAME As String = "", _
                                    Optional LOCATION As String = "") As Integer
'----------------------------------------------------------------------------
'                   �u�i�ڃ}�X�^�v�`�F�b�N���X�V����
'----------------------------------------------------------------------------
Dim com         As Integer
Dim sts         As Integer
Dim ans         As Integer
        
Dim i           As Integer
    
    
    Item_Check_Proc = True

           

    Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)

    Do

        sts = BTRV(BtOpGetEqual + BtSNoWait, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                
                com = BtOpUpdate
                                
                If Trim(HIN_NAME) <> "" Then
                    Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)   '�i��
                End If
                Exit Do
            Case BtErrKeyNotFound
                
                com = BtOpInsert
                
                Call UniCode_Conv(ITEMREC.JGYOBU, JGYOBU)           '���ƕ�
                Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI)           '�����O
                Call UniCode_Conv(ITEMREC.HIN_GAI, HIN_GAI)         '�i�ԁi�O���j
    
                Call UniCode_Conv(ITEMREC.HIN_NAME, HIN_NAME)       '�i��
    
                Call UniCode_Conv(ITEMREC.ST_SET_DT, "")            '�W���I�Ԑݒ��
                
                
                                                                    '�W���I��
                If Len(Trim(LOCATION)) > 6 Then
                    Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(LOCATION, 1, 2))
                    Call UniCode_Conv(ITEMREC.ST_RETU, Mid(LOCATION, 3, 2))
                    Call UniCode_Conv(ITEMREC.ST_REN, Mid(LOCATION, 5, 2))
                    Call UniCode_Conv(ITEMREC.ST_DAN, "01")
                
                Else
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                End If
    
    
                Call UniCode_Conv(ITEMREC.BEF_SOKO, "")             '�O����ɑq��
                Call UniCode_Conv(ITEMREC.BEF_RETU, "")
                Call UniCode_Conv(ITEMREC.BEF_REN, "")
                Call UniCode_Conv(ITEMREC.BEF_DAN, "")
    
                Call UniCode_Conv(ITEMREC.LAST_NYU_DT, "")          '�ŏI���ɓ�
                Call UniCode_Conv(ITEMREC.LAST_SYU_DT, "")          '�ŏI�o�ɓ�
    
                Call UniCode_Conv(ITEMREC.HIN_NAI, "")              '�i�ԁi�����j
    
                Call UniCode_Conv(ITEMREC.BIKOU_SOKO, "")           '���l �z�X�g�q��
                Call UniCode_Conv(ITEMREC.BIKOU_TANA, "")           '���l �z�X�g�I��
                
                Call UniCode_Conv(ITEMREC.HOJYU_P, "00000000")      '��[�_
                Call UniCode_Conv(ITEMREC.AVE_SYUKA, "00000000")    '�����Ϗo�א�
                Call UniCode_Conv(ITEMREC.SAMPLE_QTY, "1")          '�T���v����
                
                Call UniCode_Conv(ITEMREC.LAST_INP_DT, "")          '�ŏI���ד��t
                Call UniCode_Conv(ITEMREC.LAST_CHK_DT, "")          '�ŏI�ƍ����t
                
                Call UniCode_Conv(ITEMREC.LAST_CHK_QTY, "")         '�ŏI�ƍ����݌ɐ�
                
                Call UniCode_Conv(ITEMREC.BIKOU, "")                '������l
                Call UniCode_Conv(ITEMREC.IRI_QTY, "")              '������萔
                
                Call UniCode_Conv(ITEMREC.JAN_CODE, "")             'Jan�R�[�h
                Call UniCode_Conv(ITEMREC.HIN_CHANGE, "")           '�i�ԓǂݑւ�
                
                Call UniCode_Conv(ITEMREC.GOODS_KBN, GOODS_ON)      '���i���L���i�L�j
                
                Call UniCode_Conv(ITEMREC.PACKING_NO, "")           '������
                
                Call UniCode_Conv(ITEMREC.RANK, "")                 '�����ݸ
                Call UniCode_Conv(ITEMREC.NEW_RANK, "")             '�V�ݸ
                
                
                Call UniCode_Conv(ITEMREC.GLICS1_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS2_TANA, "")          '��د���I��1
                Call UniCode_Conv(ITEMREC.GLICS3_TANA, "")          '��د���I��1
                
                
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
                
                Call UniCode_Conv(ITEMREC.L_HIN_NAME_E, "")             '���i����   �i��
                Call UniCode_Conv(ITEMREC.L_BIKOU, "")                  '           ���l
                Call UniCode_Conv(ITEMREC.L_KAISHA_CODE, "")            '           ��ЃR�[�h
                Call UniCode_Conv(ITEMREC.L_KISHU1, "")                 '           �@��(1)
                Call UniCode_Conv(ITEMREC.xL_KISHU2, "")                '           �@��(2)���g�p
                Call UniCode_Conv(ITEMREC.L_KISHU3, "")                 '           �@��(3)
                Call UniCode_Conv(ITEMREC.L_PAPER, "")                  '           ��
                Call UniCode_Conv(ITEMREC.L_PLASTIC, "")                '           �v���X�`�b�N
                Call UniCode_Conv(ITEMREC.L_URIKIN1, "")                '           ���i(1)
                Call UniCode_Conv(ITEMREC.L_URIKIN2, "")                '           ���i(2)
                Call UniCode_Conv(ITEMREC.L_URIKIN3, "")                '           ���i(3)
                Call UniCode_Conv(ITEMREC.L_LABEL, "")                  '           �K�p�@������
                Call UniCode_Conv(ITEMREC.L_MAISU, "")                  '           ��������
                Call UniCode_Conv(ITEMREC.L_KISHU_BIKOU, "")            '           �K�p�@����l
                Call UniCode_Conv(ITEMREC.L_SAGYO_SHIJI, "")            '           ��Ǝw��
                Call UniCode_Conv(ITEMREC.L_BIKOU3, "")                 '           ���l�R
                Call UniCode_Conv(ITEMREC.L_JGYOBU_CODE, "")            '           ���ƕ��R�[�h
                Call UniCode_Conv(ITEMREC.L_IRI_QTY, "")                '           ���萔
                Call UniCode_Conv(ITEMREC.L_TANA1, "")                  '           �I��(1)
                Call UniCode_Conv(ITEMREC.L_TANA2, "")                  '           �I��(2)
                
                Call UniCode_Conv(ITEMREC.S_TANTO, "")                  '���P�^�S���҃R�[�h
                Call UniCode_Conv(ITEMREC.ZAIKO_F, P_ZAIKO_F_ON)        '�݌ɊǗ��ΏۗL���@�i�Ώہj
                Call UniCode_Conv(ITEMREC.L_KISHU2, "")                 '�@��(2)
                
                Call UniCode_Conv(ITEMREC.G_ZEN_ZAIKO_QTY, "00000000")  '�O���݌ɐ�
                Call UniCode_Conv(ITEMREC.G_LAST_SYUKA_QTY, "00000000") '�ŏI�o�א�
                            
                Call UniCode_Conv(ITEMREC.G_S2_ZAI_QTY, "00000000")     'S2 �݌�
                Call UniCode_Conv(ITEMREC.G_P2_ZAI_QTY, "00000000")     'P2 �݌�
                            
                Call UniCode_Conv(ITEMREC.K_KEITAI, "")                 '���`��
                            
                Call UniCode_Conv(ITEMREC.UNIT_BUHIN, "")               '�Ưĕ��i�敪
                Call UniCode_Conv(ITEMREC.NAI_BUHIN, "")                '���������敪
                Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")                '�C�O�����敪
                Call UniCode_Conv(ITEMREC.HYO_TANKA, "")                '�W���P��
    
                            
                Call UniCode_Conv(ITEMREC.FILLER, "")
                                                                        '�X�V�S����
                Call UniCode_Conv(ITEMREC.UPD_TANTO, StrConv(App.EXEName, vbUpperCase))
                                                                        '�X�V����
                Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                Exit Function
        End Select
    Loop
    
    Do
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
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
                
                
                
                    Do
                        sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr, BtErrDuplicates
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                                Exit Function
                        End Select
                    Loop
                
                
                End If
            Next i
        
        End If
        
    End If

    Item_Check_Proc = False

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

