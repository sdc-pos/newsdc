VERSION 5.00
Begin VB.Form PM000102 
   Caption         =   "�Ǘ��}�X�^�����e�i���X"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   12045
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
   ScaleHeight     =   6300
   ScaleWidth      =   12045
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   8820
      MaxLength       =   8
      TabIndex        =   15
      Text            =   "99999999"
      Top             =   1320
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   8820
      MaxLength       =   8
      TabIndex        =   14
      Text            =   "99999999"
      Top             =   840
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   8820
      MaxLength       =   8
      TabIndex        =   13
      Text            =   "99999999"
      Top             =   360
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   6300
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "999"
      Top             =   2280
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   6300
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "999"
      Top             =   1920
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   6300
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "999"
      Top             =   1560
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   6300
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "9999.99"
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   6300
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "9999.99"
      Top             =   840
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   6300
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "999.99"
      Top             =   480
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   6
      Text            =   "9999.99"
      Top             =   2640
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   5
      Text            =   "9999.99"
      Top             =   2280
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "9999.99"
      Top             =   1800
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   3
      Text            =   "9999.99"
      Top             =   1440
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   2
      Text            =   "9999.99"
      Top             =   960
      Width           =   900
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   2310
      MaxLength       =   7
      TabIndex        =   1
      Text            =   "9999.99"
      Top             =   600
      Width           =   900
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
      Index           =   11
      Left            =   10320
      TabIndex        =   27
      Top             =   5880
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����"
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   16
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "���㇂"
      Height          =   240
      Index           =   15
      Left            =   7980
      TabIndex        =   43
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   240
      Index           =   14
      Left            =   7770
      TabIndex        =   42
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "���Ϗ���"
      Height          =   240
      Index           =   13
      Left            =   7770
      TabIndex        =   41
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@���x���\�t����"
      Height          =   240
      Index           =   12
      Left            =   3780
      TabIndex        =   40
      Top             =   2400
      Width           =   2400
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@�����i�_������"
      Height          =   240
      Index           =   11
      Left            =   3780
      TabIndex        =   39
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@�����ޓ_������"
      Height          =   240
      Index           =   10
      Left            =   3780
      TabIndex        =   38
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@�@�]�T��"
      Height          =   240
      Index           =   9
      Left            =   3780
      TabIndex        =   37
      Top             =   1320
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@�@�����[�g"
      Height          =   240
      Index           =   8
      Left            =   3780
      TabIndex        =   36
      Top             =   960
      Width           =   1920
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�H���@�@���b�g��"
      Height          =   240
      Index           =   7
      Left            =   3780
      TabIndex        =   35
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@�]�T��"
      Height          =   240
      Index           =   6
      Left            =   630
      TabIndex        =   34
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@�]�T��"
      Height          =   240
      Index           =   5
      Left            =   630
      TabIndex        =   33
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�o�ׁ@�����[�g"
      Height          =   240
      Index           =   4
      Left            =   630
      TabIndex        =   32
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�o�Ɂ@�����[�g"
      Height          =   240
      Index           =   3
      Left            =   630
      TabIndex        =   31
      Top             =   1560
      Width           =   1680
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "�@�@�@�]�T��"
      Height          =   240
      Index           =   1
      Left            =   630
      TabIndex        =   30
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "���Ɂ@�����[�g"
      Height          =   240
      Index           =   0
      Left            =   630
      TabIndex        =   29
      Top             =   720
      Width           =   1680
   End
   Begin VB.Label Label 
      Caption         =   "ں��އ�"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PM000102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxRec_No% = 0                'ں��އ�(���͕s��)


Private Const ptxNYUKO_S_RATE% = 1          '���Ɂ@�����[�g
Private Const ptxNYUKO_R_RATE% = 2          '���Ɂ@�]�T��

Private Const ptxSYUKO_S_RATE% = 3          '�o�Ɂ@�����[�g
Private Const ptxSYUKO_R_RATE% = 4          '�o�Ɂ@�]�T��

Private Const ptxSYUKA_S_RATE% = 5          '�o�ׁ@�����[�g
Private Const ptxSYUKA_R_RATE% = 6          '�o�ׁ@�]�T��

Private Const ptxKOUTEI_LOT% = 7            '�H���@�O��H���W�����b�g
Private Const ptxKOUTEI_S_RATE% = 8         '�H���@�����[�g
Private Const ptxKOUTEI_R_RATE% = 9         '�H���@�]�T��
Private Const ptxKOUTEI_SHIZAI% = 10        '�H���@�����ފm�F�_��
Private Const ptxKOUTEI_BUHIN% = 11         '�H���@�������i�m�F�_��
Private Const ptxKOUTEI_LABEL% = 12         '�H���@���x���\�t����

Private Const ptxMITSUMORI_NO% = 13         '���Ϗ���
Private Const ptxSEIKYU_NO% = 14            '��������
Private Const ptxMIN_URIAGE_NO% = 15        '�~�j�}���@���㇂







Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000102.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000102)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000102)


    PM000102.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxNYUKO_S_RATE            '���Ɂ@�����[�g

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If


        Case ptxNYUKO_R_RATE            '���Ɂ@�]�T��
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKO_S_RATE            '�o�Ɂ@�����[�g

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKO_R_RATE                '�o�Ɂ@�]�T��

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKA_S_RATE                '�o�ׁ@�����[�g

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxSYUKA_R_RATE                '�o�ׁ@�]�T��
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_LOT                  '�H���@�O��H���W�����b�g
            
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If


        Case ptxKOUTEI_S_RATE               '�H���@�����[�g

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_R_RATE               '�H���@�]�T��

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_SHIZAI               '�H���@�����ފm�F�_��

            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxKOUTEI_BUHIN                '�H���@�������i�m�F�_��
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If


        Case ptxKOUTEI_LABEL                '�H���@���x���\�t����
            If Trim(Text1(Mode).Text) = "" Then
            Else
                If Not IsNumeric(Text1(Mode).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If

        Case ptxMITSUMORI_NO                '���Ϗ���
            If Not IsNumeric(Text1(Mode).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            End If
        Case ptxSEIKYU_NO                   '��������
            If Not IsNumeric(Text1(Mode).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
                Exit Function
            End If
        Case ptxMIN_URIAGE_NO               '�~�j�}�����㇂
            If Not IsNumeric(Text1(Mode).Text) Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(Mode).SetFocus
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
            
            If IsNumeric(StrConv(P_KANRIREC.NYUKO_S_RATE, vbUnicode)) Then
                Text1(ptxNYUKO_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.NYUKO_S_RATE, vbUnicode)), "#0.00")    '���Ɂ@�����[�g
            Else
                Text1(ptxNYUKO_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.NYUKO_R_RATE, vbUnicode)) Then
                Text1(ptxNYUKO_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.NYUKO_R_RATE, vbUnicode)), "#0.00")    '���Ɂ@�]�T��
            Else
                Text1(ptxNYUKO_R_RATE).Text = ""
            End If
            
            If IsNumeric(StrConv(P_KANRIREC.SYUKO_S_RATE, vbUnicode)) Then
                Text1(ptxSYUKO_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKO_S_RATE, vbUnicode)), "#0.00")    '�o�Ɂ@�����[�g
            Else
                Text1(ptxSYUKO_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.SYUKO_R_RATE, vbUnicode)) Then
                Text1(ptxSYUKO_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKO_R_RATE, vbUnicode)), "#0.00")    '�o�Ɂ@�]�T��
            Else
                Text1(ptxSYUKO_R_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.SYUKA_S_RATE, vbUnicode)) Then
                Text1(ptxSYUKA_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKA_S_RATE, vbUnicode)), "#0.00")    '�o�ׁ@�����[�g
            Else
                Text1(ptxSYUKA_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.SYUKA_R_RATE, vbUnicode)) Then
                Text1(ptxSYUKA_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.SYUKA_R_RATE, vbUnicode)), "#0.00")    '�o�ׁ@�����[�g
            Else
                Text1(ptxSYUKA_R_RATE).Text = ""
            End If
            
            
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)) Then
                Text1(ptxKOUTEI_LOT).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_LOT, vbUnicode)), "#0.00")        '�H���@�O��H���W�����b�g
            Else
                Text1(ptxKOUTEI_LOT).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)) Then
                Text1(ptxKOUTEI_S_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_S_RATE, vbUnicode)), "#0.00")  '�H���@�����[�g
            Else
                Text1(ptxKOUTEI_S_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)) Then
                Text1(ptxKOUTEI_R_RATE).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_R_RATE, vbUnicode)), "#0.00")  '�H���@�]�T��
            Else
                Text1(ptxKOUTEI_R_RATE).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_SHIZAI, vbUnicode)) Then
                Text1(ptxKOUTEI_SHIZAI).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_SHIZAI, vbUnicode)), "#")      '�H���@�����ފm�F�_��
            Else
                Text1(ptxKOUTEI_SHIZAI).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_BUHIN, vbUnicode)) Then
                Text1(ptxKOUTEI_BUHIN).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_BUHIN, vbUnicode)), "#")       '�H���@�������i�m�F�_��
            Else
                Text1(ptxKOUTEI_BUHIN).Text = ""
            End If
            If IsNumeric(StrConv(P_KANRIREC.KOUTEI_LABEL, vbUnicode)) Then
                Text1(ptxKOUTEI_LABEL).Text = Format(CDbl(StrConv(P_KANRIREC.KOUTEI_LABEL, vbUnicode)), "#")       '�H���@���x���\�t������
            Else
                Text1(ptxKOUTEI_LABEL).Text = ""
            End If
  
            If IsNumeric(StrConv(P_KANRIREC.MITSUMORI_NO, vbUnicode)) Then
                Text1(ptxMITSUMORI_NO).Text = Format(CDbl(StrConv(P_KANRIREC.MITSUMORI_NO, vbUnicode)), "00000000")     '���Ϗ���
            Else
                Text1(ptxMITSUMORI_NO).Text = "00000001"
            End If
            
            If IsNumeric(StrConv(P_KANRIREC.SEIKYU_NO, vbUnicode)) Then
                Text1(ptxSEIKYU_NO).Text = Format(CDbl(StrConv(P_KANRIREC.SEIKYU_NO, vbUnicode)), "00000000")           '��������
            Else
                Text1(ptxSEIKYU_NO).Text = "00000001"
            End If
            If IsNumeric(StrConv(P_KANRIREC.MIN_URIAGE_NO, vbUnicode)) Then
                Text1(ptxMIN_URIAGE_NO).Text = Format(CDbl(StrConv(P_KANRIREC.MIN_URIAGE_NO, vbUnicode)), "00000000")   '�~�j�}�����㇂
            Else
                Text1(ptxMIN_URIAGE_NO).Text = "00000001"
            End If
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
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�Ǘ��}�X�^")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------���R�[�h���e�ҏW
    
    Call UniCode_Conv(P_KANRIREC.REC_NO, P_ST_KANRI_No)                                 'ں��އ�
    
    If IsNumeric(Text1(ptxNYUKO_S_RATE).Text) Then                                      '���Ɂ@�����[�g
        Call UniCode_Conv(P_KANRIREC.NYUKO_S_RATE, Format(CDbl(Text1(ptxNYUKO_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.NYUKO_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxNYUKO_R_RATE).Text) Then                                      '���Ɂ@�]�T��
        Call UniCode_Conv(P_KANRIREC.NYUKO_R_RATE, Format(CDbl(Text1(ptxNYUKO_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.NYUKO_R_RATE, "")
    End If
    
    
    If IsNumeric(Text1(ptxSYUKO_S_RATE).Text) Then                                      '�o�Ɂ@�����[�g
        Call UniCode_Conv(P_KANRIREC.SYUKO_S_RATE, Format(CDbl(Text1(ptxSYUKO_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKO_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxSYUKO_R_RATE).Text) Then                                      '�o�Ɂ@�]�T��
        Call UniCode_Conv(P_KANRIREC.SYUKO_R_RATE, Format(CDbl(Text1(ptxSYUKO_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKO_R_RATE, "")
    End If
    
    If IsNumeric(Text1(ptxSYUKA_S_RATE).Text) Then                                      '�o�ׁ@�����[�g
        Call UniCode_Conv(P_KANRIREC.SYUKA_S_RATE, Format(CDbl(Text1(ptxSYUKA_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKA_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxSYUKA_S_RATE).Text) Then                                      '�o�ׁ@�]�T��
        Call UniCode_Conv(P_KANRIREC.SYUKA_R_RATE, Format(CDbl(Text1(ptxSYUKA_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.SYUKA_R_RATE, "")
    End If
    
    If IsNumeric(Text1(ptxKOUTEI_LOT).Text) Then                                        '�H���@�O��H���W�����b�g
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LOT, Format(CDbl(Text1(ptxKOUTEI_LOT).Text), "000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LOT, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_S_RATE).Text) Then                                     '�H���@�����[�g
        Call UniCode_Conv(P_KANRIREC.KOUTEI_S_RATE, Format(CDbl(Text1(ptxKOUTEI_S_RATE).Text), "0000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_S_RATE, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_R_RATE).Text) Then                                     '�H���@�]�T��
        Call UniCode_Conv(P_KANRIREC.KOUTEI_R_RATE, Format(CDbl(Text1(ptxKOUTEI_R_RATE).Text), " 000.00"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_R_RATE, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_SHIZAI).Text) Then                                     '�H���@�����ފm�F�_��
        Call UniCode_Conv(P_KANRIREC.KOUTEI_SHIZAI, Format(CDbl(Text1(ptxKOUTEI_SHIZAI).Text), "000"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_SHIZAI, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_BUHIN).Text) Then                                      '�H���@�������i�m�F�_��
        Call UniCode_Conv(P_KANRIREC.KOUTEI_BUHIN, Format(CDbl(Text1(ptxKOUTEI_BUHIN).Text), "000"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_BUHIN, "")
    End If
    If IsNumeric(Text1(ptxKOUTEI_LABEL).Text) Then                                      '�H���@���x���\�t����
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LABEL, Format(CDbl(Text1(ptxKOUTEI_LABEL).Text), "000"))
    Else
        Call UniCode_Conv(P_KANRIREC.KOUTEI_LABEL, "")
    End If
                                                                                        '���Ϗ���
    Call UniCode_Conv(P_KANRIREC.MITSUMORI_NO, Format(CDbl(Text1(ptxMITSUMORI_NO).Text), "00000000"))
                                                                                        '��������
    Call UniCode_Conv(P_KANRIREC.SEIKYU_NO, Format(CDbl(Text1(ptxSEIKYU_NO).Text), "00000000"))
                                                                                        '�~�j�}�����㇂
    Call UniCode_Conv(P_KANRIREC.MIN_URIAGE_NO, Format(CDbl(Text1(ptxMIN_URIAGE_NO).Text), "00000000"))
    
    
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
            
            
            For i = ptxRec_No To ptxSEIKYU_NO
            
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
            Text1(ptxNYUKO_S_RATE).SetFocus
        Case P_CMD_DEL                      '�폜
        Case P_CMD_DSP                      '����/�\��
        
        Case 5                              '�����l�ݒ�   2008.02.13
        
            PM000103.Show vbModal
        
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        
        Case P_CMD_End                      '�I��
            Me.Visible = False
    End Select

End Sub


Private Sub Form_Activate()
                                '��ʏ����ݒ�
    If Item_Disp_Proc() Then
        Unload Me
    End If
                                
    Text1(ptxNYUKO_S_RATE).SetFocus

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
    PM000102.Caption = PM000102.Caption & LAST_UPDATE_DAY

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

