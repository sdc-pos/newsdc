VERSION 5.00
Begin VB.Form F1020301 
   Appearance      =   0  '�ׯ�
   BackColor       =   &H80000005&
   Caption         =   "���׃`�F�b�N"
   ClientHeight    =   6876
   ClientLeft      =   2136
   ClientTop       =   2496
   ClientWidth     =   11292
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
   ScaleHeight     =   6876
   ScaleWidth      =   11292
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ComboBox Combo 
      Height          =   336
      Index           =   0
      Left            =   5880
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   15
      Top             =   120
      Width           =   852
   End
   Begin VB.TextBox Text 
      Height          =   348
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   14
      Top             =   120
      Width           =   372
   End
   Begin VB.TextBox Text 
      Height          =   348
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   13
      Top             =   120
      Width           =   372
   End
   Begin VB.TextBox Text 
      Height          =   348
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   12
      Top             =   120
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������f"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   4608
      Index           =   0
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   960
      Width           =   7452
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
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
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "�����F���̓z�X�g���Ŏ����r���A����POS�Ŕr���������ł��B"
      Height          =   240
      Left            =   1800
      TabIndex        =   29
      Top             =   5640
      Width           =   6840
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   8
      Left            =   3960
      TabIndex        =   28
      Top             =   240
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   7
      Left            =   3480
      TabIndex        =   27
      Top             =   240
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�O���j"
      Height          =   240
      Index           =   2
      Left            =   1800
      TabIndex        =   26
      Top             =   720
      Width           =   1440
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   25
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ł��B"
      Height          =   240
      Left            =   9600
      TabIndex        =   24
      Top             =   3840
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   240
      Index           =   6
      Left            =   8640
      TabIndex        =   22
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���א�"
      Height          =   240
      Index           =   5
      Left            =   7680
      TabIndex        =   21
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[��"
      Height          =   240
      Index           =   4
      Left            =   6720
      TabIndex        =   20
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�@��"
      Height          =   240
      Index           =   3
      Left            =   3480
      TabIndex        =   19
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   252
      Index           =   1
      Left            =   5040
      TabIndex        =   18
      Top             =   240
      Width           =   732
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[���t"
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   17
      Top             =   240
      Width           =   960
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1020301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxDEN_DT_YY% = ZERO      '�`�[���t�@�N
Private Const ptxDEN_DT_MM% = 1         '�`�[���t�@��
Private Const ptxDEN_DT_DD% = 2         '�`�[���t�@��

Private Const pcmbNAIGAI% = ZERO        '�����O

Private Const plstNYUKA% = ZERO         '���ח\��\�����X�g�{�b�N�X

Private Const Text_Max% = 2

Private Const H_Chk_Mark$ = "��"        '�z�X�g���������}�[�N
Private Const R_Chk_Mark$ = "��"        '���[�J�����������}�[�N

Dim WS_NO       As String

Dim U_DEN_No    As String               '�X�V���@�`�[��
Dim U_DEN_Dt    As String               '�@�@�@�@�`�[���t
Dim U_HINGAI    As String               '�@�@�@�@�O���i��
Dim U_Y_QTY     As Long                 '�@�@�@�@�\�萔��
Dim U_TEXTNO    As String               '�@�@�@�@�e�L�X�g��

Dim OK_DAY      As Integer              '�w���\�����i�����O�j
Dim OK_DATE     As String * 8           '�w���\���t

Dim NYUKA_DATA  As String               '���׃f�[�^�t���p�X

Dim PRT_CAN     As Boolean
Dim NormalFont  As New StdFont
Dim Pdate       As String               '����J�n���t�iͯ�ް�p�j
Dim Ptime       As String               '����J�n�����iͯ�ް�p�j


Private Const LMAX% = 46                '�œ��ő�s��
Private Const MGN_L% = 20               '���׈���J�n���ʒu�i�P����j
Private Const MGN_U% = 1                '��]���i�s���F�P����j


Private Sub Print_Head(Lcnt As Integer)

Dim i As Integer
Dim RetBuf As String
Dim sts As Integer

    If Lcnt <> 99 Then
        Printer.NewPage
    End If

    Printer.Print

    Printer.Print Tab(3);
    For i = ZERO To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).Code Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(41);
    Printer.Print "������  ���׃`�F�b�N���X�g   ������";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '���ׁi�P�j���
    Printer.Print Tab(5);
    Printer.Print "�`�[���t�F";
    Printer.Print Text(ptxDEN_DT_YY).Text & "/";
    Printer.Print Text(ptxDEN_DT_MM).Text & "/";
    Printer.Print Text(ptxDEN_DT_DD).Text;
    Printer.Print " �m" & Combo(pcmbNAIGAI).Text & "�n"

    Printer.Print
                                        '���ׁi�Q�j���
    Printer.Print Tab(MGN_L + 1);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 15);
    Printer.Print "�i  ��";
    Printer.Print Tab(MGN_L + 43);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 51);
    Printer.Print "�\�Z�P��";
    Printer.Print Tab(MGN_L + 68);
    Printer.Print "���א�";
    Printer.Print Tab(MGN_L + 77);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 86);
    Printer.Print "���E��";
    Printer.Print Tab(MGN_L + 95);
    Printer.Print "�W���I��"
    Printer.Print

    Lcnt = 7 + MGN_U

End Sub
Private Function Print_Proc() As Integer

Dim Lcnt        As Integer
Dim sts         As Integer
Dim i           As Integer
Dim RetBuf      As String
Dim NAIGAI      As String * 1
Dim Zan_Qty     As Long

    Print_Proc = True

    Call Input_Lock

    Label1.Visible = True
    Command1.Visible = True
    Command1.Enabled = True

    PRT_CAN = False

    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then    '�����O�̔���
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If


    Lcnt = 99
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    Pdate = Date
    Ptime = Time

    For i = ZERO To List1(plstNYUKA).ListCount - 1
        DoEvents
                                            '������f�v��
        If PRT_CAN Then
            Printer.KillDoc
            Label1.Visible = False
            Command1.Visible = False
            Call Input_UnLock
            Print_Proc = True
            Exit Function
        End If

        'Call UniCode_Conv(K2_Y_NYU.JGYOBU, Last_JGYOBU)
        'Call UniCode_Conv(K2_Y_NYU.DEN_DT, Text(ptxDEN_DT_YY) & Text(ptxDEN_DT_MM) & Text(ptxDEN_DT_DD))
        'Call UniCode_Conv(K2_Y_NYU.HIN_GAI, RTrim(Left(List1(plstNYUKA).List(i), 13)))
        'Call UniCode_Conv(K2_Y_NYU.NAIGAI, NAIGAI)
        'Call UniCode_Conv(K2_Y_NYU.DEN_NO, Mid(List1(plstNYUKA).List(i), 42, 6))
        'sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K2_Y_NYU, Len(K2_Y_NYU), 2)
        Call UniCode_Conv(K4_Y_NYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K4_Y_NYU.TEXT_NO, Right(List1(plstNYUKA).List(i), UBound(Y_NYUREC.TEXT_NO) + 1))
        sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ח\��t�@�C��")
                Exit Function
        End Select

                                        '�w�b�_�[�R���g���[��
        If sts = BtNoErr Then
            If Lcnt > LMAX Then
                Call Print_Head(Lcnt)
            End If
                                        '���׈��
            If StrConv(Y_NYUREC.DT_SYU, vbUnicode) = "R" Then
                Printer.Print Tab(MGN_L);
                Printer.Print "*";
            End If
            Printer.Print Tab(MGN_L + 1);
            Printer.Print StrConv(Y_NYUREC.HIN_GAI, vbUnicode);

            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(StrConv(Y_NYUREC.HIN_GAI, vbUnicode)))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
            Select Case sts
                Case BtNoErr
                    Printer.Print Tab(MGN_L + 15);
                    Printer.Print StrConv(ITEMREC.HIN_NAME, vbUnicode);
                Case BtErrKeyNotFound
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select

            Printer.Print Tab(MGN_L + 43);
            Printer.Print StrConv(Y_NYUREC.DEN_NO, vbUnicode);

            Printer.Print Tab(MGN_L + 51);
            Printer.Print StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode) & " ";
            Printer.Print StrConv(Y_NYUREC.YOSAN_TO, vbUnicode);
            Zan_Qty = CLng(StrConv(Y_NYUREC.YOTEI_QTY, vbUnicode)) - CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode))
            sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Format(Zan_Qty, "000000"), RetBuf)
            If Zan_Qty <> ZERO Then
                If Mid(List1(plstNYUKA).List(i), 58, 1) <> H_Chk_Mark And _
                    Mid(List1(plstNYUKA).List(i), 58, 1) <> R_Chk_Mark Then
                    Printer.Print Tab(MGN_L + 67);
                    Printer.Print RetBuf;
                Else
                    Printer.Print Tab(MGN_L + 76);
                    Printer.Print RetBuf;
                End If
            End If

            If CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode)) <> ZERO Then
                sts = Numeric_Check(EDIT_ONLY, 7, ZERO, NEGA_DIS, ZSUP_ENA, COMA_ENA, Right(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode), 6), RetBuf)
                Printer.Print Tab(MGN_L + 85);
                Printer.Print RetBuf;
            End If
            
            Printer.Print Tab(MGN_L + 94);
            Printer.Print StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-";
            Printer.Print StrConv(ITEMREC.ST_RETU, vbUnicode) & "-";
            Printer.Print StrConv(ITEMREC.ST_REN, vbUnicode) & "-";
            Printer.Print StrConv(ITEMREC.ST_DAN, vbUnicode);
            
            Printer.Print
            Printer.Print

            Lcnt = Lcnt + 2
        End If
    Next i

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If


    Label1.Visible = False
    Command1.Visible = False
    Call Input_UnLock

    Print_Proc = False

End Function
Private Function OUTPUT_Proc() As Integer

Dim sts         As Integer
Dim i           As Integer
Dim RetBuf      As String
Dim NAIGAI      As String * 1
Dim Zan_Qty     As Long

Dim Ret         As Integer
    
Dim fileName    As String
Dim FileNo      As Integer


    OUTPUT_Proc = True

    Call Input_Lock

    FileNo = FreeFile
    
    fileName = NYUKA_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    Open (fileName) For Output As FileNo


    Write #FileNo, "�`�[���t�F" & Text(ptxDEN_DT_YY).Text & "/" & Text(ptxDEN_DT_MM).Text & "/" & Text(ptxDEN_DT_DD).Text & "��"
    Write #FileNo, , "�i�ԁi�O���j", "�i��", "�`�[��", "�\�Z ��", "�\�Z ��", "���א�", "������", "�O�؂葊�E", "�W���I��"

    

    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then    '�����O�̔���
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If

    For i = ZERO To List1(plstNYUKA).ListCount - 1
        DoEvents

        'Call UniCode_Conv(K2_Y_NYU.JGYOBU, Last_JGYOBU)
        'Call UniCode_Conv(K2_Y_NYU.DEN_DT, Text(ptxDEN_DT_YY) & Text(ptxDEN_DT_MM) & Text(ptxDEN_DT_DD))
        'Call UniCode_Conv(K2_Y_NYU.HIN_GAI, RTrim(Left(List1(plstNYUKA).List(i), 13)))
        'Call UniCode_Conv(K2_Y_NYU.NAIGAI, NAIGAI)
        'Call UniCode_Conv(K2_Y_NYU.DEN_NO, Mid(List1(plstNYUKA).List(i), 42, 6))
        'sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K2_Y_NYU, Len(K2_Y_NYU), 2)
        Call UniCode_Conv(K4_Y_NYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K4_Y_NYU.TEXT_NO, Right(List1(plstNYUKA).List(i), UBound(Y_NYUREC.TEXT_NO) + 1))
        sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ח\��t�@�C��")
                Exit Function
        End Select

                                        '�w�b�_�[�R���g���[��
        If sts = BtNoErr Then
                                        '���׈��
            If StrConv(Y_NYUREC.DT_SYU, vbUnicode) = "R" Then
                Write #FileNo, "*",
            Else
                Write #FileNo, ,
            End If
            
            Write #FileNo, StrConv(Y_NYUREC.HIN_GAI, vbUnicode),

            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, RTrim(StrConv(Y_NYUREC.HIN_GAI, vbUnicode)))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
            Select Case sts
                Case BtNoErr
                    Write #FileNo, StrConv(ITEMREC.HIN_NAME, vbUnicode),
                Case BtErrKeyNotFound
                    Write #FileNo,
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select

            Write #FileNo, StrConv(Y_NYUREC.DEN_NO, vbUnicode),
            Write #FileNo, StrConv(Y_NYUREC.YOSAN_FROM, vbUnicode),
            Write #FileNo, StrConv(Y_NYUREC.YOSAN_TO, vbUnicode),
            Zan_Qty = CLng(StrConv(Y_NYUREC.YOTEI_QTY, vbUnicode)) - CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode))
            If Zan_Qty <> ZERO Then
                If Mid(List1(plstNYUKA).List(i), 58, 1) <> H_Chk_Mark And _
                    Mid(List1(plstNYUKA).List(i), 58, 1) <> R_Chk_Mark Then
                    Write #FileNo, Format(Zan_Qty, "#0"), ,
                Else
                    Write #FileNo, , Format(Zan_Qty, "#0"),
                End If
            End If

            If CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode)) <> ZERO Then
                
                Write #FileNo, Format(CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode)), "#0"),
            Else
                Write #FileNo, ,
            End If
            
            Write #FileNo, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
        End If
    Next i

    Close #FileNo
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    
    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"

    OUTPUT_Proc = False

    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If

End Function

Private Function List_Dsp() As Integer

Dim com     As Integer
Dim sts     As Integer
Dim Edit    As String
Dim NAIGAI  As String * 1

Dim i       As Integer

    List_Dsp = True

    List1(plstNYUKA).Clear
                                                    
    For i = ptxDEN_DT_YY To ptxDEN_DT_DD
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B", vbOKOnly + vbExclamation
            Text(i).SetFocus
            List_Dsp = False
            Exit Function
        Else
            Edit = Format(CInt(Text(i).Text), "0000")
            Text(i).Text = Right(Edit, Text(i).MaxLength)
        End If
    Next i
                
                
    If (Text(ptxDEN_DT_YY).Text & Text(ptxDEN_DT_MM).Text & Text(ptxDEN_DT_DD).Text) < OK_DATE Then
        MsgBox "���͂������ڂ̓G���[�ł��B�i�L�����t�͈́j", vbOKOnly + vbExclamation
        Text(ptxDEN_DT_YY).SetFocus
        List_Dsp = False
        Exit Function
    End If
                                                    
                                                    
    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then    '�����O�̔���
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                                    
                                                    '���ח\��f�[�^�ǂݍ���
    Call UniCode_Conv(K2_Y_NYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K2_Y_NYU.DEN_DT, Text(ptxDEN_DT_YY) & Text(ptxDEN_DT_MM) & Text(ptxDEN_DT_DD))
    Call UniCode_Conv(K2_Y_NYU.HIN_GAI, "")
    Call UniCode_Conv(K2_Y_NYU.NAIGAI, "")
    Call UniCode_Conv(K2_Y_NYU.DEN_NO, "")
    com = BtOpGetGreater

    Do
        sts = BTRV(com, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K2_Y_NYU, Len(K2_Y_NYU), 2)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_NYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(Y_NYUREC.DEN_DT, vbUnicode) <> (Text(ptxDEN_DT_YY) & Text(ptxDEN_DT_MM) & Text(ptxDEN_DT_DD)) Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ח\��t�@�C��")
                Exit Function
        End Select
        If StrConv(Y_NYUREC.NAIGAI, vbUnicode) = NAIGAI Then
            Call List_Edit(Edit)
            List1(plstNYUKA).AddItem Edit

            com = BtOpGetNext
        End If
    Loop

    If List1(plstNYUKA).ListCount <= ZERO Then
        Beep
        MsgBox "�Ώۃf�[�^����"
        Text(ptxDEN_DT_YY).SetFocus
    End If
    
    List_Dsp = False

End Function
Private Sub List_Edit(Edit As String)
Dim RetBuf  As String
Dim Zan_Qty As Integer
Dim sts     As Integer
    Edit = StrConv(Y_NYUREC.HIN_GAI, vbUnicode) & " "

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_NYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_NYUREC.HIN_GAI, vbUnicode))
    Do
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Exit Do
        End Select
    Loop
    Edit = Edit & StrConv(ITEMREC.HIN_NAME, vbUnicode) & "  "
    
    Edit = Edit & StrConv(Y_NYUREC.DEN_NO, vbUnicode) & "  "

    Zan_Qty = CLng(StrConv(Y_NYUREC.YOTEI_QTY, vbUnicode)) - CLng(StrConv(Y_NYUREC.BEF_NYU_QTY, vbUnicode))
    sts = Numeric_Check(EDIT_ONLY, 6, ZERO, NEGA_DIS, ZSUP_ENA, COMA_DIS, Format(Zan_Qty, "000000"), RetBuf)
    Edit = Edit & RetBuf & "  "

    'If StrConv(Y_NYUREC.CYOK_KBN, vbUnicode) = "C" Then    '2001.05.28MT
    Select Case Trim(StrConv(Y_NYUREC.CYOK_KBN, vbUnicode))
        Case "C"
            Edit = Edit & H_Chk_Mark & Space(3) & "X"
        Case "D"
            Edit = Edit & R_Chk_Mark & Space(3) & "X"
        Case Else
            Edit = Edit & "�@" & Space(3) & "A"
    End Select

    Edit = Edit & StrConv(Y_NYUREC.TEXT_NO, vbUnicode)

End Sub
                                            '�݌ɏ󋵂̍X�V
Private Function Update_Proc(i As Long) As Integer

Dim Edit    As String
Dim NAIGAI  As String * 1
Dim com     As Integer
Dim sts     As Integer
Dim ans     As Integer

    Update_Proc = True
    
    Select Case Mid(List1(plstNYUKA).List(i), 58, 1)
        Case H_Chk_Mark
            Beep
            MsgBox "���̓`�[�́A�z�X�g�V�X�e���ɂĒ����r���ς݂ł��B"
            List1(plstNYUKA).ListIndex = i
            Update_Proc = False
            Exit Function
        Case R_Chk_Mark
            Beep
            ans = MsgBox("�����r���̖߂��������s���܂����H", vbYesNo + vbQuestion, "�m�F����")
        Case Else
            Beep
            ans = MsgBox("�����r���������s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    End Select
            
    If ans <> vbYes Then
        Update_Proc = False
        Exit Function
    End If
    Edit = Right(List1(plstNYUKA).List(i), UBound(Y_NYUREC.TEXT_NO) + 1)
    
    Call UniCode_Conv(K4_Y_NYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K4_Y_NYU.TEXT_NO, Edit)
    com = BtOpGetEqual
    Do
        sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K4_Y_NYU, Len(K4_Y_NYU), 4)
        Select Case sts
            Case BtNoErr
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            
            Case BtErrKeyNotFound
                Beep
                MsgBox "�Y������f�[�^�����[���ŏ����������Ă��܂��B"
                Update_Proc = SYS_CANCEL
                Exit Function
            
            Case Else
                Call File_Error(sts, com, "���ח\��")
                Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    
    'If Mid(List1(plstNYUKA).List(i), 58, 1) = Chk_Mark Then
'    If StrConv(Y_NYUREC.CYOK_KBN, vbUnicode) = "C" Then
'        Beep
'        MsgBox "���̓`�[�́A�z�X�g�V�X�e���ɂĒ����r���ς݂ł��B"
'        List1(plstNYUKA).ListIndex = i
'        Update_Proc = False
'        Exit Function
'    End If
                                                            '�`�[���t
    'U_DEN_Dt = Text(ptxDEN_DT_YY) & Text(ptxDEN_DT_MM) & Text(ptxDEN_DT_DD)
    'U_DEN_No = Mid(List1(plstNYUKA).List(i), 42, 6)         '�`�[��
    'U_HINGAI = RTrim(Left(List1(plstNYUKA).List(i), 13))    '�i�ԁi�O���j
    'U_Y_QTY = CLng(Mid(List1(plstNYUKA).List(i), 50, 6))    '�\�萔��
    'U_TEXTNO = Right(List1(plstNYUKA).List(i), 9)           '�e�L�X�g��
    U_DEN_Dt = Trim(StrConv(Y_NYUREC.DEN_DT, vbUnicode))
    U_DEN_No = Trim(StrConv(Y_NYUREC.DEN_NO, vbUnicode))
    U_HINGAI = Trim(StrConv(Y_NYUREC.HIN_GAI, vbUnicode))
    U_Y_QTY = CLng(Trim(StrConv(Y_NYUREC.YOTEI_QTY, vbUnicode)))
    U_TEXTNO = Trim(StrConv(Y_NYUREC.TEXT_NO, vbUnicode))
    
    If U_Y_QTY <= ZERO Then
        Beep
        MsgBox "���̓`�[�́A�O�؂莩�����E����Ă��܂��B"
        List1(plstNYUKA).ListIndex = i
        Update_Proc = False
        Exit Function
    End If

    Call Input_Lock
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
    
    

    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then    '�����O�̔���
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                                '�݌��ް��̃��b�N
    sts = Zaiko_Lock_Proc(KASO_NYUKA_Soko & "01" & "01" & "01", _
                            Last_JGYOBU, _
                            NAIGAI, _
                            U_HINGAI, _
                            WS_NO)
    
    Select Case sts
        Case False, True
        Case SYS_CANCEL
            Update_Proc = False
            GoTo Abort_Tran
        Case SYS_ERR
            Update_Proc = SYS_ERR
            GoTo Abort_Tran
    End Select
            
    
    sts = Y_Nyu_Update()                        '���ח\��X�V
    Select Case sts
        Case False
        Case True
            Update_Proc = False
            GoTo Abort_Tran
        Case SYS_CANCEL
            Update_Proc = False
            GoTo Abort_Tran
        Case SYS_ERR
            Update_Proc = SYS_ERR
            GoTo Abort_Tran
    End Select

                                                '�݌��ް��̊J��
    sts = Zaiko_UNLock_Proc(KASO_NYUKA_Soko & "01" & "01" & "01", _
                            Last_JGYOBU, _
                            NAIGAI, _
                            U_HINGAI, _
                            "")
    Select Case sts
        Case False, True
        Case SYS_CANCEL
            Update_Proc = False
            GoTo Abort_Tran
        Case SYS_ERR
            Update_Proc = SYS_ERR
            GoTo Abort_Tran
    End Select

                                        '�g�����U�N�V�����I��
End_Tran:
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    'Edit = Left(List1(plstNYUKA).List(i), 57) & Chk_Mark & Right(List1(plstNYUKA).List(i), 12)
    
    List1(plstNYUKA).RemoveItem i
    Call List_Edit(Edit)
    List1(plstNYUKA).AddItem Edit, i
    List1(plstNYUKA).ListIndex = i
    
    
    Call Input_UnLock
    
    List1(plstNYUKA).SetFocus
    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock

End Function

Private Function Y_Nyu_Update() As Integer

Dim sts     As Integer
Dim RetBuf  As String
Dim Qty     As Long
Dim prog    As String
Dim com     As Integer
Dim NAIGAI  As String * 1
Dim ans     As Integer

    Y_Nyu_Update = True

    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then    '�����O�̔���
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If

'���ח\��f�[�^�ǂݍ���
    Call UniCode_Conv(K0_Y_NYU.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_Y_NYU.DEN_DT, U_DEN_Dt)
    Call UniCode_Conv(K0_Y_NYU.DEN_NO, U_DEN_No)

    com = BtOpGetEqual + BtSNoWait
    Do
        sts = BTRV(com, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
        Select Case sts
            Case BtNoErr
                If Trim(StrConv(Y_NYUREC.TEXT_NO, vbUnicode)) = U_TEXTNO And _
                    StrConv(Y_NYUREC.NAIGAI, vbUnicode) = NAIGAI And _
                   RTrim(StrConv(Y_NYUREC.HIN_GAI, vbUnicode)) = U_HINGAI Then
                    If Trim(StrConv(Y_NYUREC.CYOK_KBN, vbUnicode)) = "" Then
                        Call UniCode_Conv(Y_NYUREC.CYOK_KBN, "D")
                    Else
                        Call UniCode_Conv(Y_NYUREC.CYOK_KBN, "")
                    End If
'���ח\��f�[�^�X�V
                    Do
                        sts = BTRV(BtOpUpdate, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Y_Nyu_Update = SYS_CANCEL
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "�o�ח\��")
                                Y_Nyu_Update = SYS_ERR
                                Exit Function
                        End Select
                    Loop
                    
                    Exit Do
                
                End If
                
                com = BtOpGetNext + BtSNoWait
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Y_Nyu_Update = SYS_CANCEL
                    Exit Function
                End If
            
            Case BtErrKeyNotFound
                Beep
                MsgBox "�Y������f�[�^�����[���ŏ����������Ă��܂��B"
                Y_Nyu_Update = SYS_CANCEL
                Exit Function
            
            Case Else
                Call File_Error(sts, com, "���ח\��")
                Y_Nyu_Update = SYS_ERR
                Exit Function
        End Select
    Loop

    If Trim(StrConv(Y_NYUREC.CYOK_KBN, vbUnicode)) <> "" Then
        '�݌��ް��o�ɏ���
        sts = Syuko_Update_Proc(Last_JGYOBU, _
                            NAIGAI, _
                            U_HINGAI, _
                            StrConv(Y_NYUREC.DEN_DT, vbUnicode), _
                            KASO_NYUKA_Soko & "01" & "01" & "01", _
                            YOIN_CHOKUSO, _
                            U_Y_QTY, _
                            WS_NO)
    Else
        '�݌��ް����ɏ���
        sts = Nyuko_Update_Proc(Last_JGYOBU, _
                            NAIGAI, _
                            U_HINGAI, _
                            StrConv(Y_NYUREC.DEN_DT, vbUnicode), _
                            KASO_NYUKA_Soko & "01" & "01" & "01", _
                            YOIN_CHOKU_MODOSI, _
                            U_Y_QTY, _
                            WS_NO)
    End If
    
    If sts Then
        Y_Nyu_Update = sts
        Exit Function
    End If
    
    Y_Nyu_Update = False

End Function


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �R���{�{�b�N�X���́i�j�����c�������j����
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
        Case pcmbNAIGAI
            If List_Dsp() Then
                Unload Me
            End If
            
            If List1(plstNYUKA).ListCount <> ZERO Then
                List1(plstNYUKA).SetFocus
                List1(plstNYUKA).ListIndex = ZERO
            End If
    End Select

End Sub
Private Sub Command_Click(Index As Integer)
Dim ans As Integer
Dim sts As Integer

    Select Case Index
        Case 7
        
            If List1(plstNYUKA).ListCount <> ZERO Then
                Beep
                ans = MsgBox("�u���ח\��v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            
                If ans = vbYes Then
                    If OUTPUT_Proc() Then
                        Unload Me
                    End If
                End If
                List1(plstNYUKA).SetFocus
                List1(plstNYUKA).ListIndex = ZERO
            Else
                Beep
                MsgBox "�Ώۃf�[�^���\������Ă��܂���"
                Text(ptxDEN_DT_YY).SetFocus
            End If
        
        
        Case 8
            If List1(plstNYUKA).ListCount <> ZERO Then
                If Print_Proc() Then
                    Unload Me
                End If
                List1(plstNYUKA).SetFocus
                List1(plstNYUKA).ListIndex = ZERO
            Else
                Beep
                MsgBox "�Ώۃf�[�^���\������Ă��܂���"
                Text(ptxDEN_DT_YY).SetFocus
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select

End Sub
Private Sub Command1_Click()
    PRT_CAN = True
End Sub
Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = ZERO
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
Dim Work As String

Dim sBuffer As String * 255
Dim com     As String

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
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

                                '���׃f�[�^�t�@�C������荞��
    If GetIni("FILE", "NYUKA_DATA", "SYS", c) Then
        Beep
        MsgBox "���׃f�[�^�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    NYUKA_DATA = Trim(c)
'�w��\������荞��
    If GetIni("SYSTEM", "NYUKA_OK_DAY", "SYS", c) Then
        OK_DAY = 1
    Else
        If IsNumeric(Trim(c)) Then
            OK_DAY = CInt(Trim(c))
        Else
            OK_DAY = 1
        End If
    End If
'�V�X�e���\��ϗv����荞��
    If SYSTEM_YOIN_Set() Then
        Beep
        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

'���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = ZERO To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1020301.Caption = "���׃`�F�b�N�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i

'���z�q�ɔԍ��ԍ���荞��
    If Kaso_Soko_No_Set() Then
        Beep
        MsgBox "���z�q�ɂ̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

'�[���ԍ���荞��
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> ZERO Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�i�X�V�p���[�N�j�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���Ɏ��тn�o�d�m
    If J_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��t�@�C���n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����ް��n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1020301.FontName
        .Size = F1020301.FontSize
    End With
    Set Printer.Font = NormalFont

'�`�[���t�����ݒ�
    Work = DateAdd("d", -1, Date)
    Text(ptxDEN_DT_YY).Text = Left(Work, 4)
    Text(ptxDEN_DT_MM).Text = Mid(Work, 6, 2)
    Text(ptxDEN_DT_DD).Text = Right(Work, 2)
    
    OK_DATE = DateAdd("d", OK_DAY, Date)
    
'���ޏ����ݒ�i�����O�j
    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
                                '��̫�ā�����
    Combo(pcmbNAIGAI).ListIndex = ZERO
    
    Text(ptxDEN_DT_YY).SetFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�i�X�V�p���[�N�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^�i�X�V�p�j")
        End If
    End If
                                            '���Ɏ��тb�k�n�r�d
    sts = BTRV(BtOpClose, J_NYU_POS, J_NYUREC, Len(J_NYUREC), K0_J_NYU, Len(K0_J_NYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���Ɏ���")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '���ח\��t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ח\��t�@�C��")
        End If
    End If
                                            '�݌��ް��b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌��ް�")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), ZERO)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), ZERO)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1020301 = Nothing
    End
End Sub

Private Sub List1_DblClick(Index As Integer)

    If Index = plstNYUKA Then
        If Update_Proc(List1(Index).ListIndex) = True Then
            Unload Me
        End If
    End If

End Sub

Private Sub List1_GotFocus(Index As Integer)

    If List1(Index).ListCount > ZERO Then
        If List1(Index).ListIndex <= ZERO Then
            List1(Index).ListIndex = ZERO
        End If
    End If

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
Dim sts As Integer
    
    If List1(Index).ListCount = ZERO Then
        Exit Sub
    End If
    
    If Index = plstNYUKA Then
        Select Case KeyCode
            Case vbKeyReturn
                If Update_Proc(List1(Index).ListIndex) = True Then
                    Unload Me
                End If
            Case vbKeyEscape
                List1(Index).Clear
                Text(ptxDEN_DT_YY).SetFocus
        End Select
    End If
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1020301.Caption = "���׃`�F�b�N�i" + RTrim(JGYOBU_T(Index).NAME) + "�j"
    Last_JGYOBU = JGYOBU_T(Index).Code
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

    List1(plstNYUKA).Clear
    Text(ptxDEN_DT_YY).SetFocus

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = ZERO
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf As String
Dim i As Integer


    If KeyCode <> vbKeyReturn Then Exit Sub
            
    If Index = ptxDEN_DT_DD Then
        For i = ptxDEN_DT_YY To ptxDEN_DT_DD
            If Not IsNumeric(Text(i).Text) Then
 '              Beep
                MsgBox "���͂������ڂ̓G���[�ł��B", vbOKOnly + vbExclamation
                Text(i).SetFocus
                Exit Sub
            Else
                RetBuf = Format(CInt(Text(i).Text), "0000")
                Text(i).Text = Right(RetBuf, Text(i).MaxLength)
            End If
        Next i
                
                
        If (Text(ptxDEN_DT_YY).Text & Text(ptxDEN_DT_MM).Text & Text(ptxDEN_DT_DD).Text) < OK_DATE Then
            MsgBox "���͂������ڂ̓G���[�ł��B�i�L�����t�͈́j", vbOKOnly + vbExclamation
            Text(ptxDEN_DT_YY).SetFocus
            Exit Sub
        End If
                
        Combo(pcmbNAIGAI).SetFocus
                
        Exit Sub
    End If

    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1020301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020301)


    F1020301.MousePointer = vbDefault

End Sub

