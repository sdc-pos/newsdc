VERSION 5.00
Begin VB.Form PM000101 
   Caption         =   "�Ǘ��}�X�^�����e�i���X"
   ClientHeight    =   8250
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   13155
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
   ScaleHeight     =   8250
   ScaleWidth      =   13155
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   8295
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2160
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   18
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   18
      Top             =   6240
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   16
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   15
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   14
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   14
      Top             =   3840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   13
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   13
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   17
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   17
      Top             =   5760
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   8295
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   6135
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2160
      Width           =   240
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   6135
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   8055
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7215
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   6135
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1320
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
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
      TabIndex        =   30
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   8
      Left            =   7800
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   7
      Left            =   6480
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   5
      Left            =   4800
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   4
      Left            =   3960
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   2640
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X �V"
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
      TabIndex        =   19
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "�d�����z�@�ۂ�"
      Height          =   255
      Index           =   24
      Left            =   420
      TabIndex        =   55
      Top             =   6360
      Width           =   1710
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "������z�@�ۂ�"
      Height          =   255
      Index           =   22
      Left            =   420
      TabIndex        =   54
      Top             =   5880
      Width           =   1710
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "FAX�ԍ�"
      Height          =   255
      Index           =   21
      Left            =   1200
      TabIndex        =   53
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "�d�b�ԍ�"
      Height          =   255
      Index           =   20
      Left            =   1200
      TabIndex        =   52
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "�Z���^�[��"
      Height          =   255
      Index           =   19
      Left            =   840
      TabIndex        =   51
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "��Ж�"
      Height          =   255
      Index           =   18
      Left            =   1440
      TabIndex        =   50
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label LblS_TANTO 
      Alignment       =   1  '�E����
      Height          =   255
      Left            =   3120
      TabIndex        =   49
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "���F��"
      Height          =   255
      Index           =   17
      Left            =   1440
      TabIndex        =   48
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "���ۂ�:(0:�؎̂� 5:�l�̌ܓ� 9:�؏グ)"
      Height          =   255
      Index           =   16
      Left            =   6720
      TabIndex        =   47
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label 
      Caption         =   "    �ۂ�"
      Height          =   255
      Index           =   15
      Left            =   7215
      TabIndex        =   46
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   14
      Left            =   9015
      TabIndex        =   45
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "�V�@�ŗ�"
      Height          =   255
      Index           =   13
      Left            =   7215
      TabIndex        =   44
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "    �ۂ�"
      Height          =   255
      Index           =   12
      Left            =   5055
      TabIndex        =   43
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "%"
      Height          =   255
      Index           =   11
      Left            =   6855
      TabIndex        =   42
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "���@�ŗ�"
      Height          =   255
      Index           =   10
      Left            =   5055
      TabIndex        =   41
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "��"
      Height          =   255
      Index           =   9
      Left            =   8535
      TabIndex        =   40
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "��"
      Height          =   255
      Index           =   8
      Left            =   7695
      TabIndex        =   39
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "�N"
      Height          =   255
      Index           =   7
      Left            =   6855
      TabIndex        =   38
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "����ŕύX���t"
      Height          =   255
      Index           =   6
      Left            =   4335
      TabIndex        =   37
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "���㇂"
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   36
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      Caption         =   "������"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   35
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�w�}�[��"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   34
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "(����=31)"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   33
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "��������"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "ں��އ�"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   31
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PM000101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxRec_No% = 0                'ں��އ�(���͕s��)
Private Const ptxSHIME_DD% = 1              '�������ߏ���

Private Const ptxSASHIZU_NO% = 2            '�w�}����
Private Const ptxORDER_NO% = 3              '������
Private Const ptxURIAGE_NO% = 4             '���ޔ���ں��އ�

Private Const ptxZEI_CHANGE_YY% = 5         '����ŕύX���t �N
Private Const ptxZEI_CHANGE_MM% = 6         '����ŕύX���t ��
Private Const ptxZEI_CHANGE_DD% = 7         '����ŕύX���t ��

Private Const ptxNOW_ZEI_RITU% = 8          '�V�@����ŗ�
Private Const ptxNOW_MARUME% = 9            '�@�@�ۂ�

Private Const ptxNEW_ZEI_RITU% = 10         '�V�@����ŗ�
Private Const ptxNEW_MARUME% = 11           '�@�@�ۂ�

Private Const ptxSHONIN_CODE% = 12          '���F�S��
Private Const ptxKAISHA_NAME% = 13          '��Ж�
Private Const ptxCENTER_NAME% = 14          '�Z���^�[��
Private Const ptxTEL_NO% = 15               '�d�b�ԍ�
Private Const ptxFAX_NO% = 16               'FAX�ԍ�

Private Const ptxURI_MARUME% = 17           '������z�@�ۂ�
Private Const ptxSHI_MARUME% = 18           '�d�����z�@�ۂ�







Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000101.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000101)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000101)


    PM000101.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxSHIME_DD    '�������ߓ�
            If Not IsNumeric(Text1(ptxSHIME_DD).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxSHIME_DD).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxSHIME_DD).Text) < 1 Or CInt(Text1(ptxSHIME_DD).Text) > 31 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxSHIME_DD).SetFocus
                Exit Function
            End If
    
            Text1(ptxSHIME_DD).Text = Format(CInt(Text1(ptxSHIME_DD).Text), "00")
    
        Case ptxSASHIZU_NO  '�w�}����
            
            If Not IsNumeric(Text1(ptxSASHIZU_NO).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxSASHIZU_NO).SetFocus
                Exit Function
            End If
    
'            If CLng(Text1(ptxSASHIZU_NO).Text) < 0 Or CLng(Text1(ptxSASHIZU_NO).Text) > 99999 Then     '2007.11.28
            If CLng(Text1(ptxSASHIZU_NO).Text) < 0 Or CLng(Text1(ptxSASHIZU_NO).Text) > 99999999 Then   '2007.11.28
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxSASHIZU_NO).SetFocus
                Exit Function
            End If
            
'            Text1(ptxSASHIZU_NO).Text = Format(CLng(Text1(ptxSASHIZU_NO).Text), "00000")       '2007.11.28
            Text1(ptxSASHIZU_NO).Text = Format(CLng(Text1(ptxSASHIZU_NO).Text), "00000000")     '2007.11.28
    
        Case ptxORDER_NO    '������
            
            If Not IsNumeric(Text1(ptxORDER_NO).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxORDER_NO).SetFocus
                Exit Function
            End If
    
            If CLng(Text1(ptxORDER_NO).Text) < 0 Or CLng(Text1(ptxORDER_NO).Text) > 99999 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxORDER_NO).SetFocus
                Exit Function
            End If
            
            Text1(ptxORDER_NO).Text = Format(CLng(Text1(ptxORDER_NO).Text), "00000")
    
        Case ptxURIAGE_NO   '���ޔ���ں��އ�
            
            If Not IsNumeric(Text1(ptxURIAGE_NO).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxURIAGE_NO).SetFocus
                Exit Function
            End If
    
            If CLng(Text1(ptxURIAGE_NO).Text) < 0 Or CLng(Text1(ptxURIAGE_NO).Text) > 99999 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxURIAGE_NO).SetFocus
                Exit Function
            End If
            
            Text1(ptxURIAGE_NO).Text = Format(CLng(Text1(ptxURIAGE_NO).Text), "00000")
    
        Case ptxZEI_CHANGE_YY '����ŕύX���t�@�N
            If Not IsNumeric(Text1(ptxZEI_CHANGE_YY).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxZEI_CHANGE_YY).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxZEI_CHANGE_YY).Text) < 1 Or CInt(Text1(ptxZEI_CHANGE_YY).Text) > 9999 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxZEI_CHANGE_YY).SetFocus
                Exit Function
            End If
            
            Text1(ptxZEI_CHANGE_YY).Text = Format(CInt(Text1(ptxZEI_CHANGE_YY).Text), "0000")
    
        Case ptxZEI_CHANGE_MM '����ŕύX���t�@��
            If Not IsNumeric(Text1(ptxZEI_CHANGE_MM).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxZEI_CHANGE_MM).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxZEI_CHANGE_MM).Text) < 1 Or CInt(Text1(ptxZEI_CHANGE_MM).Text) > 12 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxZEI_CHANGE_MM).SetFocus
                Exit Function
            End If
            
            Text1(ptxZEI_CHANGE_MM).Text = Format(CInt(Text1(ptxZEI_CHANGE_MM).Text), "00")
    
        Case ptxZEI_CHANGE_DD   '����ŕύX���t�@��
            If Not IsNumeric(Text1(ptxZEI_CHANGE_DD).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxZEI_CHANGE_DD).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxZEI_CHANGE_DD).Text) < 1 Or CInt(Text1(ptxZEI_CHANGE_DD).Text) > 31 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxZEI_CHANGE_DD).SetFocus
                Exit Function
            End If
            
            Text1(ptxZEI_CHANGE_DD).Text = Format(CInt(Text1(ptxZEI_CHANGE_DD).Text), "00")
            '���tOK�H
            If Not IsDate(Text1(ptxZEI_CHANGE_YY).Text & "/" & Text1(ptxZEI_CHANGE_MM).Text & "/" & Text1(ptxZEI_CHANGE_DD).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxZEI_CHANGE_YY).SetFocus
                Exit Function
            End If
    
        Case ptxNOW_ZEI_RITU    '���@�ŗ�
            If Not IsNumeric(Text1(ptxNOW_ZEI_RITU).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
    
            If CDbl(Text1(ptxNOW_ZEI_RITU).Text) < 0 Or CDbl(Text1(ptxNOW_ZEI_RITU).Text) > 99.9 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
            
            Text1(ptxNOW_ZEI_RITU).Text = Format(CDbl(Text1(ptxNOW_ZEI_RITU).Text), "#0.0")
    
        Case ptxNOW_MARUME      '���@�ۂ�
            If Not IsNumeric(Text1(ptxNOW_MARUME).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNOW_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxNOW_MARUME).Text) <> 0 And _
                CInt(Text1(ptxNOW_MARUME).Text) <> 5 And _
                CInt(Text1(ptxNOW_MARUME).Text) <> 9 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNOW_MARUME).SetFocus
                Exit Function
            End If
            
        Case ptxNEW_ZEI_RITU    '�V�@�ŗ�
            If Not IsNumeric(Text1(ptxNEW_ZEI_RITU).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
    
            If CDbl(Text1(ptxNEW_ZEI_RITU).Text) < 0 Or CDbl(Text1(ptxNEW_ZEI_RITU).Text) > 99.9 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNOW_ZEI_RITU).SetFocus
                Exit Function
            End If
            
            Text1(ptxNEW_ZEI_RITU).Text = Format(CDbl(Text1(ptxNEW_ZEI_RITU).Text), "#0.0")
    
        Case ptxNEW_MARUME      '���@�ۂ�
            If Not IsNumeric(Text1(ptxNEW_MARUME).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNEW_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxNEW_MARUME).Text) <> 0 And _
                CInt(Text1(ptxNEW_MARUME).Text) <> 5 And _
                CInt(Text1(ptxNEW_MARUME).Text) <> 9 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxNEW_MARUME).SetFocus
                Exit Function
            End If
    
        Case ptxSHONIN_CODE     '���F�S���Һ���
            If Trim(Text1(ptxSHONIN_CODE).Text) = "" Then
                LblS_TANTO.Caption = ""
            Else
                Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).Text)
            
                sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                Select Case sts
                    Case BtNoErr
                        LblS_TANTO.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        LblS_TANTO.Caption = ""
                        MsgBox "���͂������ڂ̓G���[�ł��B"
                        Text1(ptxSHONIN_CODE).SetFocus
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                        Exit Function
                End Select
            
            
            End If
    
    
        Case ptxKAISHA_NAME     '��Ж�
        Case ptxCENTER_NAME     '��Ж�
        Case ptxTEL_NO          '�d�b�ԍ�
        Case ptxFAX_NO          'FAX�ԍ�
    
        Case ptxURI_MARUME      '������z�@�ۂ�
            If Not IsNumeric(Text1(ptxURI_MARUME).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxURI_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxURI_MARUME).Text) <> 0 And _
                CInt(Text1(ptxURI_MARUME).Text) <> 5 And _
                CInt(Text1(ptxURI_MARUME).Text) <> 9 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxURI_MARUME).SetFocus
                Exit Function
            End If
    
        Case ptxSHI_MARUME      '�d�����z�@�ۂ�
            If Not IsNumeric(Text1(ptxSHI_MARUME).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxSHI_MARUME).SetFocus
                Exit Function
            End If
    
            If CInt(Text1(ptxSHI_MARUME).Text) <> 0 And _
                CInt(Text1(ptxSHI_MARUME).Text) <> 5 And _
                CInt(Text1(ptxSHI_MARUME).Text) <> 9 Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxSHI_MARUME).SetFocus
                Exit Function
            End If
    
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    '�Ǘ�Ͻ��iKEY=0�j�ǂݍ���
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    
    
    
    Text1(ptxRec_No).Text = P_ST_KANRI_No   'ں��އ�
    
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
            'ں��ޓ��e�̕\��
                                            '�������ߓ�
            Text1(ptxSHIME_DD).Text = StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
                                            '�w�}����
            Text1(ptxSASHIZU_NO).Text = StrConv(P_KANRIREC.SASHIZU_NO, vbUnicode)
                                            '������
            Text1(ptxORDER_NO).Text = StrConv(P_KANRIREC.ORDER_NO, vbUnicode)
                                            '���ޔ���ں��އ�
            Text1(ptxURIAGE_NO).Text = StrConv(P_KANRIREC.URIAGE_NO, vbUnicode)
                                            '����ŕύX���t �N
            Text1(ptxZEI_CHANGE_YY).Text = Left(StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode), 4)
                                            '����ŕύX���t ��
            Text1(ptxZEI_CHANGE_MM).Text = Mid(StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode), 5, 2)
                                            '����ŕύX���t ��
            Text1(ptxZEI_CHANGE_DD).Text = Right(StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode), 2)
                                            '���@����ŗ�
            Text1(ptxNOW_ZEI_RITU).Text = Format(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)), "#0.0")
                                            '���@�ۂ�
            Text1(ptxNOW_MARUME).Text = StrConv(P_KANRIREC.NOW_MARUME, vbUnicode)
                                            '�V�@����ŗ�
            Text1(ptxNEW_ZEI_RITU).Text = Format(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)), "#0.0")
                                            '�V�@�ۂ�
            Text1(ptxNEW_MARUME).Text = StrConv(P_KANRIREC.NEW_MARUME, vbUnicode)
                                            '���F�S����
            Text1(ptxSHONIN_CODE).Text = Trim(StrConv(P_KANRIREC.SHONIN_CODE, vbUnicode))
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxSHONIN_CODE).Text)
        
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    LblS_TANTO.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    LblS_TANTO.Caption = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Exit Function
            End Select
                                            '��Ж�
            Text1(ptxKAISHA_NAME).Text = Trim(StrConv(P_KANRIREC.KAISHA_NAME, vbUnicode))
                                            '�Z���^�[��
            Text1(ptxCENTER_NAME).Text = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))
                                            '�d�b�ԍ�
            Text1(ptxTEL_NO).Text = Trim(StrConv(P_KANRIREC.TEL_NO, vbUnicode))
                                            'FAX�ԍ�
            Text1(ptxFAX_NO).Text = Trim(StrConv(P_KANRIREC.FAX_NO, vbUnicode))
                                            '������z�@�ۂ�
            Text1(ptxURI_MARUME).Text = StrConv(P_KANRIREC.URI_MARUME, vbUnicode)
                                            '�d�����z�@�ۂ�
            Text1(ptxSHI_MARUME).Text = StrConv(P_KANRIREC.SHI_MARUME, vbUnicode)
        
        
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^")
            Exit Function
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �Ǘ��}�X�^�o��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer


    Update_Proc = True
    '�Ǘ�Ͻ��iKEY=0�j�ǂݍ���
    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = True
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�Ǘ��}�X�^")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------���R�[�h���e�ҏW
    
    Call UniCode_Conv(P_KANRIREC.REC_NO, P_ST_KANRI_No)                                 'ں��އ�
    Call UniCode_Conv(P_KANRIREC.SHIME_DD, Text1(ptxSHIME_DD).Text)                     '�������ߓ�
    
    Call UniCode_Conv(P_KANRIREC.xSASHIZU_NO, "")                                       '�w�}����2007.11.28
    
    Call UniCode_Conv(P_KANRIREC.SASHIZU_NO, Text1(ptxSASHIZU_NO).Text)                 '�w�}����
    Call UniCode_Conv(P_KANRIREC.ORDER_NO, Text1(ptxORDER_NO).Text)                     '������
    Call UniCode_Conv(P_KANRIREC.URIAGE_NO, Text1(ptxURIAGE_NO).Text)                   '���ޔ���ں��އ�
    Call UniCode_Conv(P_KANRIREC.ZEI_CHANGE_YMD, Text1(ptxZEI_CHANGE_YY).Text & _
                                            Text1(ptxZEI_CHANGE_MM).Text & _
                                            Text1(ptxZEI_CHANGE_DD).Text)               '����ŕύX���t
    Call UniCode_Conv(P_KANRIREC.NOW_ZEI_RITU, Format(CDbl(Text1(ptxNOW_ZEI_RITU).Text), "00.0"))   '���@�ŗ�
    Call UniCode_Conv(P_KANRIREC.NOW_MARUME, Text1(ptxNOW_MARUME).Text)                             '���@�܂��
    Call UniCode_Conv(P_KANRIREC.NEW_ZEI_RITU, Format(CDbl(Text1(ptxNEW_ZEI_RITU).Text), "00.0"))   '�V�@�ŗ�
    Call UniCode_Conv(P_KANRIREC.NEW_MARUME, Text1(ptxNEW_MARUME).Text)                             '�V�@�܂��
    Call UniCode_Conv(P_KANRIREC.SHONIN_CODE, Text1(ptxSHONIN_CODE).Text)               '���F�S����
    Call UniCode_Conv(P_KANRIREC.KAISHA_NAME, Text1(ptxKAISHA_NAME).Text)               '��Ж�
    Call UniCode_Conv(P_KANRIREC.CENTER_NAME, Text1(ptxCENTER_NAME).Text)               '�Z���^�[��
    Call UniCode_Conv(P_KANRIREC.TEL_NO, Text1(ptxTEL_NO).Text)                         '�d�b�ԍ�
    Call UniCode_Conv(P_KANRIREC.FAX_NO, Text1(ptxFAX_NO).Text)                         'FAX�ԍ�
    
    Call UniCode_Conv(P_KANRIREC.URI_MARUME, Text1(ptxURI_MARUME).Text)                 '������z�@�ۂ�
    Call UniCode_Conv(P_KANRIREC.SHI_MARUME, Text1(ptxSHI_MARUME).Text)                 '�d�����z�@�ۂ�
        
    
    Call UniCode_Conv(P_KANRIREC.FILLER, "")                                            'Filler
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_KANRI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "�Ǘ��}�X�^")
                Exit Function
        End Select
    Loop

    Update_Proc = False


End Function


Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd        '�X�V
            
            
            For i = ptxRec_No To ptxSHONIN_CODE
            
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            
            Next i
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            Text1(ptxSASHIZU_NO).SetFocus
        Case P_CMD_DEL                      '�폜
        Case P_CMD_DSP                      '����/�\��
        
        Case 5                              '�����֘A���ڂ� 2008.02.13
        
            PM000102.Show vbModal
        
        
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        
        Case P_CMD_End                      '�I��
            Unload Me
    End Select

End Sub

Private Sub Form_DblClick()
'    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

                                '���O�t�@�C������荞��
'    If GetIni("FILE", "LOGF", "SYS", c) Then
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
    PM000101.Caption = PM000101.Caption & LAST_UPDATE_DAY
                                
                                
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                
                                '��ʏ����ݒ�
    If Item_Disp_Proc() Then
        Unload Me
    End If
                                
    Text1(ptxSHIME_DD).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�Ǘ��}�X�^")
        End If
    End If
    
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000101 = Nothing

    End
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�
End Sub

