VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "TDBG8.OCX"
Begin VB.Form F1050351 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���o�Ɏ��яƉ�"
   ClientHeight    =   10710
   ClientLeft      =   795
   ClientTop       =   -90
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   14910
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   8640
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8655
      Left            =   0
      OleObjectBlob   =   "F1050351.frx":0000
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1080
      Width           =   14895
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   300
      Index           =   0
      Left            =   1080
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2760
      MaxLength       =   20
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�ĕ\��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I �D"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
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
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�t ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   30
      Top             =   10440
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   3960
      TabIndex        =   29
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   3480
      TabIndex        =   28
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   2640
      TabIndex        =   27
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   2160
      TabIndex        =   26
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   1680
      TabIndex        =   25
      Top             =   720
      Width           =   252
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���t"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   24
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   120
      TabIndex        =   23
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i��"
      Height          =   252
      Index           =   0
      Left            =   2040
      TabIndex        =   22
      Top             =   240
      Width           =   612
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1050351"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbNAIGAI% = 0           '�����O

Private Const ptxHin_Gai% = 0           '�i�ԁi�O���j
Private Const ptxHin_Name% = 1          '�i��
Private Const ptxST_DT_YY% = 2          '�J�n���t �N
Private Const ptxST_DT_MM% = 3          '�J�n���t ��
Private Const ptxST_DT_DD% = 4          '�J�n���t ��
Private Const ptxEN_DT_YY% = 5          '�I�����t �N
Private Const ptxEN_DT_MM% = 6          '�I�����t ��
Private Const ptxEN_DT_DD% = 7          '�I�����t ��

Private Const Text_Max% = 7

Dim IDO     As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��
'Private Const Max_Row& = 2000           '�ő�s��
Dim Max_Row     As Long                 '���X�g�{�b�N�X�ő�\������

Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 22             '�ő��

Private Const ColHin_Gai% = 0           '�� �i�ԁi�O���j
Private Const ColHin_Name% = 1          '�� �i��
Private Const ColRIRK% = 2              '�� ����
Private Const ColDEN_DT% = 3            '�� �`�[���t
Private Const ColBIN_No% = 4            '�� ��
Private Const ColDEN_No% = 5            '�� �`�[��
Private Const ColNyuko_Qty% = 6         '�� ���ɐ�
Private Const ColSyuko_Qty% = 7         '�� �o�ɐ�
Private Const ColZAITEI_Qty% = 8        '�� �݌ɒ�����
Private Const ColIDO_Qty% = 9           '�� �ړ���
Private Const ColHin_Zaiko_Qty% = 10    '�� �i�ڕʍ݌ɐ�
Private Const ColMUKE_DNAME% = 11       '�� ������
Private Const ColTANTO_NAME% = 12       '�� ID
Private Const ColMEMO% = 13             '�� ����
Private Const ColJITU_DT% = 14          '�� ���ѓ��t
Private Const ColJITU_TM% = 15          '�� ���ю���
Private Const ColFrom_Location% = 16    '�� From�I
Private Const ColTO_Location% = 17      '�� To�I
Private Const ColNYUKA_DT% = 18         '�� ���ד�
Private Const ColHin_Nai% = 19          '�� �i�ԁi�����j
Private Const ColSS_Name% = 20          '�� �����於
Private Const ColTOKU_MARK% = 21        '�� ������}�[�N
Private Const ColID_NO% = 22            '�� �`�[�h�c

Private Sort_Tbl(ColHin_Gai To ColID_NO) _
                As Integer                  '��Ă̐��� 0:���� 1:�~��

Dim Excel_Put_Flg       As Boolean      '�I�D�o�͗L��


Dim Excel_Template      As String       '�I�D ����ڰ�(�٥�߽)
Dim Excel_PutPath       As String       '�I�D �������ݐ��߽

Dim Excel_Put_Yoin_IN   As Variant      '�I�D ���ɑΏۗv���z��
Dim Excel_Put_Yoin_OUT  As Variant      '�I�D ���ɑΏۗv���z��


Dim Excel_Bin_Mei       As Variant      '�I�D �֖��̔z��
Dim ExcelApp            As Excel.Application
Dim Excelbook           As Excel.Workbook
Dim ExcelWorkSheet      As Excel.Worksheet

Private Function Err_Chk_Proc() As Integer

'���t�͈͓��̓G���[�`�F�b�N    2007.5.15 (�����ۼ��ެ��)

Dim sts         As Integer
Dim ans         As Integer
Dim i           As Integer


    Err_Chk_Proc = True


    For i = ptxST_DT_YY To ptxEN_DT_DD
        Select Case i
            Case ptxST_DT_YY, ptxEN_DT_YY
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_YY Then
                        Text(i).Text = "0000"
                    Else
                        Text(i).Text = "9999"
                    End If
                Else
                    If Not IsNumeric(Text(i).Text) Then
                    Else
                        Text(i).Text = Format(CInt(Text(i).Text), "0000")
                    End If
                End If
            Case Else
                If Len(Trim(Text(i).Text)) = 0 Then
                    If i = ptxST_DT_MM Or i = ptxST_DT_DD Then

                        Text(i).Text = "00"
                    Else
                        Text(i).Text = "99"
                 End If
            Else
                If Not IsNumeric(Text(i).Text) Then
                Else
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            End If
        End Select
    Next i

    If (Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text) > _
        (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxST_DT_YY).SetFocus
        Exit Function
    End If


    Err_Chk_Proc = False

End Function

Private Function List_Disp_Proc(Mode As Integer) As Integer
                                    '��ʕ\�����e�ݒ�
                                    'Mode = 0:����
                                    'mode = 1:�~��
Dim sts         As Integer
Dim com         As Integer
Dim Key_Mode    As Integer
Dim NAIGAI      As String * 1

Dim ans         As Integer
Dim i           As Integer
Dim Row         As Long

Dim Skip_flg    As Boolean  '2004.07.16

    List_Disp_Proc = True

                                    '�G���[�`�F�b�N
    sts = Item_Read_Proc()
    Select Case sts
        Case False
        Case True
            ans = MsgBox("�i�ڃ}�X�^�͓o�^����Ă܂���B �������p�����܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbNo Then
                Text(ptxHin_Gai).SetFocus
                List_Disp_Proc = False
                Exit Function
            End If
        Case Else
            Exit Function
    End Select

    If Err_Chk_Proc Then            '���t�͈͓��ʹװ����    2007.5.15 (�����ۼ��ެ��)
        List_Disp_Proc = False
        Exit Function
    End If


    Call Input_Lock

                                    '�e�[�u�����Z�b�g
    Set IDO = Nothing

    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1                '����
            NAIGAI = NAIGAI_NAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI2                '�C�O
            NAIGAI = NAIGAI_GAI
            If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
                Key_Mode = 0
            Else
                Key_Mode = 1
            End If
        Case NAIGAI0                '���O�w��Ȃ�
            Key_Mode = 0
    End Select


                                    '�݌Ɉړ���ǂݍ��݊J�n
    If Key_Mode = 0 Then
                                    '���n��œǂݍ���
        Call UniCode_Conv(K0_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "")           '����
        Else
            Call UniCode_Conv(K0_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K0_IDO.JITU_TM, "zzzzzzzz")   '�~��
        End If
                                    '��\������ �i�ԁ^�i��
        TDBGrid1.Columns(ColHin_Gai).Visible = True
        TDBGrid1.Columns(ColHin_Name).Visible = True

    Else
                                    '�i�ԁ����n��œǂݍ���
        Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
        Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)
        If Mode = 0 Then
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "")           '����
        Else
            Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text)
            Call UniCode_Conv(K1_IDO.JITU_TM, "zzzzzzzz")   '�~��
        End If
                                    '��\���Ȃ� �i�ԁ^�i��
        TDBGrid1.Columns(ColHin_Gai).Visible = False
        TDBGrid1.Columns(ColHin_Name).Visible = False
    End If


    Row = Min_Row - 1

    If Mode = 0 Then
        com = BtOpGetGreater        '����
    Else
        com = BtOpGetLess           '�~��
    End If
    Do
        If Key_Mode = 0 Then
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
        Else
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
        End If

        Skip_flg = False

        Select Case sts
            Case BtNoErr

                If Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "A" Or _
                    Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1) = "C" Then
                    Skip_flg = True
                End If

            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ɉړ���")
                List_Disp_Proc = SYS_ERR
        End Select

        If Not Skip_flg Then

                                    '���ƕ� KEY��ڰ�
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
                                    '���t�͈͊O
            If Mode = 0 Then
                                    '����
                If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
                    Exit Do
                End If
            Else
                                    '�~��
                If StrConv(IDOREC.JITU_DT, vbUnicode) < (Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text) Then
                    Exit Do
                End If
            End If

            If Key_Mode = 1 Then
                                    '�i�Ԏw�莞�A�i����ڰ�
                If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                    Exit Do
                End If
            End If


            If Key_Mode = 0 Then
                If StrConv(IDOREC.NAIGAI, vbUnicode) = NAIGAI Then
                    Row = Row + 1
                    If Row > Max_Row Then
                        Beep
                        MsgBox "�ő�\���s���𒴂��܂����B"
                        Exit Do
                    End If
                    Call Grid_Set_Proc(Row)
                End If
            Else
                Row = Row + 1
                If Row > Max_Row Then
                    Beep
                    MsgBox "�ő�\���s���𒴂��܂����B"
                    Exit Do
                End If

                Call Grid_Set_Proc(Row)
            End If

        End If

        If Mode = 0 Then
            com = BtOpGetNext   '����
        Else
            com = BtOpGetPrev   '�~��
        End If
        DoEvents
    Loop
                                'DB�e�[�u�������N
    Set TDBGrid1.Array = IDO
    TDBGrid1.ReBind

    TDBGrid1.Update
    TDBGrid1.MoveFirst


    Call Input_UnLock




    Text(ptxHin_Gai).SetFocus

    List_Disp_Proc = False

End Function

Private Function Tana_Fuda_Put() As Integer

'   ���ޗ��ʌ��i�I�D�@�쐬                  2007.5.15

Dim strExelFile     As String
Dim Rec_Cnt         As Long
Dim Page_Offset     As Long
Dim posG            As Long

Dim sts             As Integer
Dim com             As Integer
Dim Key_Mode        As Integer
Dim NAIGAI          As String * 1
Dim ans             As Integer
Dim i               As Integer
Dim Skip_flg        As Boolean

'On Error GoTo ERR_PRT


    Tana_Fuda_Put = True

    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1
            NAIGAI = NAIGAI_NAI
        Case NAIGAI2
            NAIGAI = NAIGAI_GAI
        
        
        
        Case NAIGAI0
            MsgBox "�����O�͏ȗ��ł��܂���B", vbExclamation
            Text(ptxHin_Gai).SetFocus
            Tana_Fuda_Put = False
            Exit Function
    
        
    
    End Select
                                    '�G���[�`�F�b�N
    If Len(Trim(Text(ptxHin_Gai).Text)) = 0 Then
        MsgBox "�i�Ԃ͏ȗ��ł��܂���B", vbExclamation
        Text(ptxHin_Gai).SetFocus
        Tana_Fuda_Put = False
        Exit Function
    End If

    sts = Item_Read_Proc()
    Select Case sts
        Case False
        Case True
            ans = MsgBox("�i�ڃ}�X�^�͓o�^����Ă��܂���B�������p�����܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbNo Then
                Text(ptxHin_Gai).SetFocus
                Tana_Fuda_Put = False
                Exit Function
            End If
        Case Else
            Exit Function
    End Select

    If Err_Chk_Proc Then            '���ʹװ����    2007.5.15 (�����ۼ��ެ��)
        Exit Function
    End If


    Call Input_Lock

                                    '�o��̧�ٖ��ҏW
    strExelFile = Excel_PutPath & Trim(Text(ptxHin_Gai).Text) & ".xls"

    'Excel���ع���ݵ�޼ު�Ď擾
    Set ExcelApp = CreateObject("Excel.Application")

    Set Excelbook = ExcelApp.Workbooks.Open(Excel_Template)         '����ڰ��ޯ����J��
    Set ExcelWorkSheet = Excelbook.Worksheets(1)                    '�P��Ėڂ�I��

    '�i��
    ExcelWorkSheet.Application.Cells(3, 2).Value = Trim(Text(ptxHin_Gai).Text)
    '���s��
    ExcelWorkSheet.Application.Cells(1, 8).Value = Format(Now, "yyyy/m/d")

                                    '�i�ԁ����n��œǂݍ���
    Call UniCode_Conv(K1_IDO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_IDO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_IDO.HIN_GAI, Text(ptxHin_Gai).Text)

    Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxST_DT_YY).Text & Text(ptxST_DT_MM).Text & Text(ptxST_DT_DD).Text)
    Call UniCode_Conv(K1_IDO.JITU_TM, "")

    Rec_Cnt = 0
    Page_Offset = 6
    posG = 6

    com = BtOpGetGreater
    Do
        sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)

        Skip_flg = False

        Select Case sts
            Case BtNoErr


            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ɉړ���")
                Tana_Fuda_Put = SYS_ERR
        End Select

        If Not Skip_flg Then
                                    '���ƕ� KEY��ڰ�
            If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                Exit Do
            End If
                                    '���t�͈͊O
            If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxEN_DT_YY).Text & Text(ptxEN_DT_MM).Text & Text(ptxEN_DT_DD).Text) Then
                Exit Do
            End If

                                    '�i�Ԏw�莞�A�i����ڰ�
            If StrConv(IDOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(Text(ptxHin_Gai).Text) Then
                Exit Do
            End If


            For i = 0 To UBound(Excel_Put_Yoin_IN)
                If Trim(StrConv(IDOREC.RIRK_ID, vbUnicode)) = Excel_Put_Yoin_IN(i) Then
                    Call TanaFuda_Set_Proc(1, posG, Page_Offset)
                    Rec_Cnt = Rec_Cnt + 1
                    Exit For
                End If
            Next i
        
        
            For i = 0 To UBound(Excel_Put_Yoin_OUT)
                If Trim(StrConv(IDOREC.RIRK_ID, vbUnicode)) = Excel_Put_Yoin_OUT(i) Then
                    Call TanaFuda_Set_Proc(2, posG, Page_Offset)
                    Rec_Cnt = Rec_Cnt + 1
                    Exit For
                End If
            Next i
        
        End If

        com = BtOpGetNext
        DoEvents
    Loop

    '���Y�y�[�W�̎c��s���N���A
    If posG <= Page_Offset + 35 Then
        Call UniCode_Conv(IDOREC.JITU_DT, "")
        Call UniCode_Conv(IDOREC.BIN_NO, "")
        Call UniCode_Conv(IDOREC.DEN_NO, "")
        Call UniCode_Conv(IDOREC.SUM_KBN, "")
        Call UniCode_Conv(IDOREC.TANTO_NAME, "")
        Call UniCode_Conv(IDOREC.RIRK_NAME, "")
        Call UniCode_Conv(IDOREC.RIRK_ID, "")           '2007.09.06
        Do
            If posG > Page_Offset + 35 Then
                Exit Do
            End If
            Call TanaFuda_Set_Proc(0, posG, Page_Offset)        '�P�s���ҏW
        Loop
    End If



    '�ҏW����ܰ���Ă̐擪���\�������l�ɁuA1�v��è�ނɂ���
    ExcelWorkSheet.Application.Range("A1").Activate

    ExcelApp.DisplayAlerts = False              'ϸێ��s�װ�͕\�����Ȃ�


    If Rec_Cnt > 0 Then
        On Error Resume Next
     '   Kill strExelFile
        ExcelWorkSheet.SaveAs strExelFile
'        On Error GoTo 0
    End If


    ExcelApp.Visible = False
    ExcelApp.Workbooks.Close                                        'ܰ��ޯ�����
    ExcelApp.Quit

    Set ExcelWorkSheet = Nothing                                    'ܰ���ĊJ��
    Set Excelbook = Nothing                                         'ܰ��ޯ��J��

    Set ExcelApp = Nothing                                         'ܰ��ޯ��J��


    Call Input_UnLock

    Text(ptxHin_Gai).SetFocus


    Tana_Fuda_Put = False

End Function

Private Sub TanaFuda_Set_Proc(InOut As Integer, posG As Long, Page_Offset As Long)


'InOut =0(DUMMY) =1(In) =2()


Dim c   As String * 128


    '�P�ŕ��ҏW�����ˎ��ŕ��̃t�H�[�}�b�g���R�s�[
    If posG > Page_Offset + 35 Then
        ExcelWorkSheet.Application.Range(Page_Offset & ":" & Page_Offset + 35).Copy
        ExcelWorkSheet.Application.Range(Page_Offset + 36 & ":" & Page_Offset + 71).Select
        ExcelWorkSheet.Paste

        Page_Offset = Page_Offset + 36
        posG = Page_Offset
    End If

                                            '���ѓ��t
    If Len(Trim(StrConv(IDOREC.JITU_DT, vbUnicode))) <> 0 Then
        ExcelWorkSheet.Application.Cells(posG, 1).Value = _
                                          Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" _
                                        & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" _
                                        & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2)
    Else
        ExcelWorkSheet.Application.Cells(posG, 1).Value = ""
    End If
                                            '��
    Select Case StrConv(IDOREC.BIN_NO, vbUnicode)
        Case "01"         '�P��
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(0)
        Case "02"         '�Q��
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(1)
        Case "03"         '�R��
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Excel_Bin_Mei(2)
        Case Else
            ExcelWorkSheet.Application.Cells(posG, 2).Value = Trim(StrConv(IDOREC.BIN_NO, vbUnicode))
    End Select
                                            '�`�[��
    If InOut = 1 Then
        ExcelWorkSheet.Application.Cells(posG, 3).Value = Trim(StrConv(IDOREC.DEN_NO, vbUnicode))
    Else
        ExcelWorkSheet.Application.Cells(posG, 3).Value = ""    '2007.09.06
    End If
                                            '���ѐ�
    ExcelWorkSheet.Application.Cells(posG, 4).Value = ""
    ExcelWorkSheet.Application.Cells(posG, 5).Value = ""
    ExcelWorkSheet.Application.Cells(posG, 6).Value = ""
    Select Case InOut
        Case 1         '���ɐ�
            ExcelWorkSheet.Application.Cells(posG, 4).Value = _
                Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
        
        
                                            '�i�ڕʍ݌ɐ�
            ExcelWorkSheet.Application.Cells(posG, 6).Value = _
                Val(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + Val(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))
        
        Case 2         '�o�ɐ�
            ExcelWorkSheet.Application.Cells(posG, 5).Value = _
                Val(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + Val(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))
    
                                            '�i�ڕʍ݌ɐ�
            ExcelWorkSheet.Application.Cells(posG, 6).Value = _
                Val(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + Val(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode))
    
    End Select
                                            '�S����
    ExcelWorkSheet.Application.Cells(posG, 7).Value = Trim(StrConv(IDOREC.TANTO_NAME, vbUnicode))
                                        
                                        '�����i�v�����́j
    If GetIni(App.EXEName, StrConv(IDOREC.RIRK_ID, vbUnicode), "SYS", c) Then
        ExcelWorkSheet.Application.Cells(posG, 8).Value = ""
    Else
        ExcelWorkSheet.Application.Cells(posG, 8).Value = Trim(c)
    End If
    
    

    posG = posG + 1

End Sub
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer
   
    For i = Mode To Text_Max
        Text(i).Text = ""
    Next i
    
End Sub
                                    '�i�ڃ}�X�^���e���ڂ�\������
Private Function Item_Read_Proc() As Integer

Dim sts     As Integer
Dim NAIGAI  As String * 1

    Item_Read_Proc = True
                                                '�����O�̔���
    Select Case Combo(pcmbNAIGAI).Text
        Case NAIGAI1
            NAIGAI = NAIGAI_NAI
        Case NAIGAI2
            NAIGAI = NAIGAI_GAI
        Case NAIGAI0
            Text(ptxHin_Gai).Text = ""
            Text(ptxHin_Name).Text = ""
            Item_Read_Proc = False
            Exit Function
    End Select
                                                
    If Len(Text(ptxHin_Gai).Text) = 0 Then
        Text(ptxHin_Name).Text = ""
        Item_Read_Proc = False
        Exit Function
    End If
                                                '�܂��O���i�Ԃœǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        
        Case BtErrKeyNotFound
                                                '�����i�Ԃōēx�ǂݍ���
            Call UniCode_Conv(K2_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K2_ITEM.NAIGAI, NAIGAI)
            Call UniCode_Conv(K2_ITEM.HIN_NAI, Text(ptxHin_Gai).Text)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
            Select Case sts
                Case BtNoErr
                    
                    Text(ptxHin_Gai).Text = StrConv(ITEMREC.HIN_GAI, vbUnicode)
                    Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    
        
                Case BtErrKeyNotFound
        
                    Exit Function
        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Item_Read_Proc = SYS_ERR
            End Select
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Item_Read_Proc = SYS_ERR
    End Select
            
    Item_Read_Proc = False

End Function

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbNAIGAI
            Call Clear_Field(0)
            
            If Combo(Index).Text = NAIGAI0 Then
                Text(ptxHin_Gai).Text = ""
                Text(ptxHin_Name).Text = ""
                Text(ptxST_DT_YY).SetFocus
            Else
                Text(ptxHin_Gai).SetFocus
            End If
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim sts As Integer
Dim i   As Integer


On Error Resume Next
    Select Case Index
        Case 0                           '�����\��
            If List_Disp_Proc(0) Then
                Unload Me
            End If

            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             '��̫�ď���
            Next i


'            IDO.QuickSort Min_Row, (IDO.UpperBound(1)), 4, XORDER_ASCEND, XTYPE_DATE, 5, XORDER_ASCEND, XTYPE_DATE
'            TDBGrid1.Refresh
'            Exit Sub
        Case 3                              '�t���\��
            If List_Disp_Proc(1) Then
                Unload Me
            End If
'            IDO.QuickSort Min_Row, (IDO.UpperBound(1)), 4, XORDER_DESCEND, XTYPE_DATE, 5, XORDER_DESCEND, XTYPE_DATE
'            TDBGrid1.Refresh
'            Exit Sub

            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             '��̫�ď���
            Next i
            Sort_Tbl(ColJITU_DT) = 1        '��̫�č~��
            Sort_Tbl(ColJITU_TM) = 1        '��̫�č~��


        Case 4                             '�I�D        2007.5.15
            If Tana_Fuda_Put() Then
                Unload Me
            End If


        Case 7                             '�ĕ\��
            If List_Disp_Proc(0) Then
                Unload Me
            End If

            For i = 0 To UBound(Sort_Tbl)
                Sort_Tbl(i) = 0             '��̫�ď���
            Next i


        Case 8                             '�ް��o��

            Call Select_Set_Proc

            F1050352.Show vbModal

        Case 11                            '�I��
            Unload Me
        Case Else
            Beep
    End Select

End Sub

Private Sub Form_DblClick()
    
Call Form_HCopy(Picture1, vbPRPSA4, vbPRORLandscape)
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
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i       As Integer
Dim c       As String
Dim sts     As Integer

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
                    '�ő�\�������̊l��
    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "�ő�\�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    Max_Row = CLng(RTrim(c))








    '�I�D�p��`����荞�� 2007.05.15      ��������

                                            
                                            
                    '�I�D�o�͗L��
    If GetIni(App.EXEName, "Excel_Put", "SYS", c) Then
        Excel_Put_Flg = False
    Else
        If Trim(c) = "1" Then
            Excel_Put_Flg = True
        Else
            Excel_Put_Flg = False
        End If
    End If
                                            
                                            
    If Excel_Put_Flg Then
                                                '����ڰ�(�٥�߽)
        If GetIni(App.EXEName, "F105035_EXCEL_TEMPLATE", "SYS", c) Then
            Beep
            MsgBox "����ڰ�(�٥�߽)�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        Excel_Template = Trim(c)
                                                '�������ݐ��߽
        If GetIni(App.EXEName, "F105035_EXCEL_OUTPUT", "SYS", c) Then
            Beep
            MsgBox "�������ݐ��߽�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        Excel_PutPath = Trim(c)
                                                '�Ώۓ��ɗv���z��
        If GetIni(App.EXEName, "YOIN_IN", "SYS", c) Then
            Beep
            MsgBox "�Ώۓ��ɗv���z��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        Excel_Put_Yoin_IN = Split(Trim(c), ",", -1)
                                                '�Ώۏo�ɗv���z��
        If GetIni(App.EXEName, "YOIN_OUT", "SYS", c) Then
            Beep
            MsgBox "�Ώۏo�ɗv���z��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        Excel_Put_Yoin_OUT = Split(Trim(c), ",", -1)
                                                '�֖��̔z��
        If GetIni("F105035", "BIN", "SYS", c) Then
            Beep
            MsgBox "�֖��̔z��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
            End
        End If
        Excel_Bin_Mei = Split(Trim(c), ",", -1)
    
    
    Else
    
        Command(4).Enabled = False
    End If
    '�I�D�p��`����荞�� 2007.05.15      �����܂�



                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1050351.Caption = "���o�Ɏ��яƉ�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                '�����O��荞��
    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
    Combo(pcmbNAIGAI).AddItem NAIGAI0
    Combo(pcmbNAIGAI).Text = NAIGAI1
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If


    Load F1050352
    Load SDC_FLD_F
                                '��ʏ����ݒ�
    Call Clear_Field(0)


    TDBGrid1.Columns(ColHin_Gai).Visible = False
    TDBGrid1.Columns(ColHin_Name).Visible = False

    TDBGrid1.Style.Locked = True

    Combo(pcmbNAIGAI).SetFocus

    End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1050351 = Nothing
    Set F1050352 = Nothing
    Set SDC_FLD_F = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1050351.Caption = "���o�Ɏ��яƉ�i" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
        IDO.QuickSort Min_Row, IDO.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        
        Set TDBGrid1.Array = IDO
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If



End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i   As Integer
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        Case ptxHin_Gai             '�i��
            
            If (Combo(pcmbNAIGAI).Text = NAIGAI0 Or _
                Len(Trim(Text(ptxHin_Gai).Text)) = 0) Then
            Else
                sts = Item_Read_Proc()
                Select Case sts
                    Case False
                    Case True
                        Text(ptxHin_Name).Text = ""
                    Case SYS_ERR
                        Unload Me
                End Select
            End If
                        
        Case ptxST_DT_YY, ptxEN_DT_YY
            If Len(Trim(Text(Index).Text)) = 0 Then
                If Index = ptxST_DT_YY Then
                    Text(Index).Text = "0000"
                Else
                    Text(Index).Text = "9999"
                End If
            Else
                If Not IsNumeric(Text(Index).Text) Then
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "0000")
                End If
            End If
    
        Case ptxST_DT_MM, ptxST_DT_DD, ptxEN_DT_MM, ptxEN_DT_DD
            If Len(Trim(Text(Index).Text)) = 0 Then
                If Index = ptxST_DT_MM Or Index = ptxST_DT_DD Then
                                
                    Text(Index).Text = "00"
                Else
                    Text(Index).Text = "99"
                 End If
            Else
                If Not IsNumeric(Text(Index).Text) Then
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
    
            If Index = ptxEN_DT_DD Then
                If List_Disp_Proc(0) Then
                    Unload Me
                End If
            End If
    
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1050351.MousePointer = vbHourglass

    Call Ctrl_Lock(F1050351)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1050351)


    F1050351.MousePointer = vbDefault

End Sub

Private Sub Grid_Set_Proc(Row As Long)


    IDO.ReDim Min_Row, Row, Min_Col, Max_Col
                                            '�i�ځi�O���j
    IDO(Row, ColHin_Gai) = StrConv(IDOREC.HIN_GAI, vbUnicode)       '�i�ځi�O���j
                                            '�i��
    IDO(Row, ColHin_Name) = StrConv(IDOREC.HIN_NAME, vbUnicode)     '�i�ږ���
                                            '�����i�v���j
    IDO(Row, ColRIRK) = StrConv(IDOREC.RIRK_NAME, vbUnicode)        '�v������
                                            '������}�[�N
    IDO(Row, ColTOKU_MARK) = StrConv(IDOREC.TOKU_MARK, vbUnicode)   '������}�[�N
                                            '���ѓ��t
    IDO(Row, ColJITU_DT) = Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" _
                            & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" _
                            & Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2)
                                            '���ю���
    IDO(Row, ColJITU_TM) = Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 1, 2) & ":" _
                            & Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 3, 2) & ":" _
                            & Mid(StrConv(IDOREC.JITU_TM, vbUnicode), 5, 2)
                                            '�`�[���t
    If Len(Trim(StrConv(IDOREC.DEN_DT, vbUnicode))) <> 0 Then
        IDO(Row, ColDEN_DT) = Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(IDOREC.DEN_DT, vbUnicode), 7, 2)
    End If
                                            '��         2007.5.16
    IDO(Row, ColBIN_No) = StrConv(IDOREC.BIN_NO, vbUnicode)
                                            '�`�[��
    IDO(Row, ColDEN_No) = StrConv(IDOREC.DEN_NO, vbUnicode)
                                            '�i�ڕʍ݌ɐ�
    IDO(Row, ColHin_Zaiko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_HIN_Zaiko_Qty, vbUnicode)) + CLng(StrConv(IDOREC.MI_HIN_Zaiko_Qty, vbUnicode)), "#,##0")
                                            '���ѐ�
    Select Case StrConv(IDOREC.SUM_KBN, vbUnicode)
        Case SUM_KBN_IN
                                            '���ɐ�
            IDO(Row, ColNyuko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
        Case SUM_KBN_OT
                                            '�o�ɐ�
            IDO(Row, ColSyuko_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                
        Case SUM_KBN_ZT
                If Mid(StrConv(IDOREC.RIRK_ID, vbUnicode), 1, 1) = ACT_ZAITEI_IN Then
                                            '�ݒ��i�{�j
                    IDO(Row, ColZAITEI_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                Else
                                            '�ݒ��i�|�j
                    IDO(Row, ColZAITEI_Qty) = Format((CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode))) * -1), "#,##0")
                End If
        
        Case SUM_KBN_MV
                                            '�ړ���
                IDO(Row, ColIDO_Qty) = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
    End Select
                                            'FROM�I
    If Len(Trim(StrConv(IDOREC.FROM_SOKO, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColFrom_Location) = StrConv(IDOREC.FROM_SOKO, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_REN, vbUnicode) & "-" _
                                    & StrConv(IDOREC.FROM_DAN, vbUnicode)
    End If
                                            'TO�I
    If Len(Trim(StrConv(IDOREC.TO_SOKO, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColTO_Location) = StrConv(IDOREC.TO_SOKO, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_RETU, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_REN, vbUnicode) & "-" _
                                    & StrConv(IDOREC.TO_DAN, vbUnicode)
    End If
                                            '���ד�
    If Len(Trim(StrConv(IDOREC.NYUKA_DT, vbUnicode))) = 0 Then
    Else
        IDO(Row, ColNYUKA_DT) = Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" _
                                & Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" _
                                & Mid(StrConv(IDOREC.NYUKA_DT, vbUnicode), 7, 2)
    End If
                                            '������
    IDO(Row, ColMUKE_DNAME) = StrConv(IDOREC.MUKE_DNAME, vbUnicode)
                                            '�S����
    IDO(Row, ColTANTO_NAME) = StrConv(IDOREC.TANTO_NAME, vbUnicode)
                                            '�i�ԁi�����j
    IDO(Row, ColHin_Nai) = StrConv(IDOREC.HIN_NAI, vbUnicode)
                                            '����
    IDO(Row, ColMEMO) = StrConv(IDOREC.MEMO, vbUnicode)
                                            '�`�[�h�c
    IDO(Row, ColID_NO) = StrConv(IDOREC.ID_NO, vbUnicode)
                            
'    TDBGrid1.Update
End Sub

Private Sub Select_Set_Proc()

    F1050352.Combo(pcmbNAIGAI).Text = Combo(pcmbNAIGAI).Text
    F1050352.Text(ptxHin_Gai).Text = Text(ptxHin_Gai).Text
    F1050352.Text(ptxHin_Name).Text = Text(ptxHin_Name).Text
    F1050352.Text(ptxST_DT_YY).Text = Text(ptxST_DT_YY).Text
    F1050352.Text(ptxST_DT_MM).Text = Text(ptxST_DT_MM).Text
    F1050352.Text(ptxST_DT_DD).Text = Text(ptxST_DT_DD).Text
    If Len(Trim(Text(ptxEN_DT_YY).Text)) = 0 Then
        F1050352.Text(ptxEN_DT_YY).Text = "9999"
    Else
        F1050352.Text(ptxEN_DT_YY).Text = Text(ptxEN_DT_YY).Text
    End If
        
    If Len(Trim(Text(ptxEN_DT_MM).Text)) = 0 Then
        F1050352.Text(ptxEN_DT_MM).Text = "99"
    Else
        F1050352.Text(ptxEN_DT_MM).Text = Text(ptxEN_DT_MM).Text
    End If
    
    If Len(Trim(Text(ptxEN_DT_DD).Text)) = 0 Then
        F1050352.Text(ptxEN_DT_DD).Text = "99"
    Else
        F1050352.Text(ptxEN_DT_DD).Text = Text(ptxEN_DT_DD).Text
    End If
End Sub
