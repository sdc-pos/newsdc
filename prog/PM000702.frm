VERSION 5.00
Begin VB.Form PM000702 
   Caption         =   "�󕥐�}�X�^�����e�i���X"
   ClientHeight    =   7155
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   12645
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
   ScaleHeight     =   7155
   ScaleWidth      =   12645
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      ItemData        =   "PM000702.frx":0000
      Left            =   2160
      List            =   "PM000702.frx":0002
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   1320
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   9
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   11
      Top             =   5520
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   8
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   10
      Top             =   5040
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   8
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   4
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   6
      Top             =   2880
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   3
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   2
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1920
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      ItemData        =   "PM000702.frx":0004
      Left            =   2640
      List            =   "PM000702.frx":0006
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   840
      Width           =   2805
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
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
      TabIndex        =   23
      Top             =   6240
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6240
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
      Index           =   5
      Left            =   4800
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6240
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
      Index           =   3
      Left            =   2640
      TabIndex        =   15
      Top             =   6240
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6240
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
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�����敪"
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   36
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "�X�֔ԍ�"
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   35
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�Z���Q"
      Height          =   255
      Index           =   11
      Left            =   1320
      TabIndex        =   34
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "�Z���P"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   33
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "�i�u-�v�܂ށj"
      Height          =   255
      Index           =   8
      Left            =   4200
      TabIndex        =   32
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "FAX�ԍ�"
      Height          =   255
      Index           =   7
      Left            =   1200
      TabIndex        =   31
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�i�u-�v�܂ށj"
      Height          =   255
      Index           =   6
      Left            =   4200
      TabIndex        =   30
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "�d�b�ԍ�"
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   29
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�����^�c�Ə���"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   28
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label 
      Caption         =   "����"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   27
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "�󕥐於��"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   26
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "���x����"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�󕥐溰��"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   24
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "PM000702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�e�L�X�g�p�Y��
Private Const ptxUKEHARAI_CODE% = 0         '�󕥂��溰��
Private Const ptxSYUSHI_CODE% = 1           '���x����
Private Const ptxUKEHARAI_NAME% = 2         '�󕥐於��
Private Const ptxUKEHARAI_RNAME% = 3        '�󕥐旪��
Private Const ptxBUSHO_NAME% = 4            '��������
Private Const ptxTEL_NO% = 5                '�d�b�ԍ�
Private Const ptxFAX_NO% = 6                'FAX�ԍ�
Private Const ptxYUBIN_NO% = 7              '�X�֔ԍ�
Private Const ptxADDR1% = 8                 '�Z���P
Private Const ptxADDR2% = 9                 '�Z���Q

Private Const Mode_All% = 0
'�R���{�p�Y��
Private Const pcmbSYUSHI% = 0               '���x
Private Const pcmbTORI_KBN% = 1             '�����敪



Private INIT_FLG    As Boolean
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000701.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000701)


    PM000701.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer
    
Dim i       As Integer          '2019.04.04
    
    Error_Check_Proc = True
    
    
    
    Select Case Mode
        
        Case Mode_All, ptxUKEHARAI_CODE     '�󕥐�R�[�h
            
            Text1(Mode).Text = StrConv(RTrim(Text1(Mode).Text), vbUpperCase)
            
            
            
            If Trim(Text1(ptxUKEHARAI_CODE).Text) = "" Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxUKEHARAI_CODE).SetFocus
                Exit Function
            End If
            
        
            If G_SCREEN_FLG = G_SCREEN_INS Then
                '�V�K���͏d���`�F�b�N
                Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)
            
                sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                Select Case sts
                    Case BtNoErr
                        ans = MsgBox("���͂����R�[�h�́A�o�^�ςł��B�X�V�����Ƃ��Čp�����܂����H")
                        If ans = vbNo Then
                            Text1(ptxUKEHARAI_CODE).SetFocus
                            Exit Function
                        End If
                    
                        Call Item_Disp_Proc(Text1(ptxUKEHARAI_CODE).Text)
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
                        Exit Function
                End Select
            
            
                Text1(ptxUKEHARAI_CODE).BackColor = G_INPUT_NG
                Text1(ptxUKEHARAI_CODE).Locked = True
                Text1(ptxUKEHARAI_CODE).TabStop = False
            
            End If
        
        Case Mode_All, ptxSYUSHI_CODE       '���x����
        
'>>>>>>>>>>>>>> 2019.04.04
            For i = 0 To Combo1(pcmbSYUSHI).ListCount - 1
                    
                If Text1(ptxSYUSHI_CODE).Text = Right(Combo1(pcmbSYUSHI).List(i), 3) Then
                    Combo1(pcmbSYUSHI).ListIndex = i
                    Exit For
                End If
            
            Next i
        
            If i > (Combo1(pcmbSYUSHI).ListCount - 1) Then
                MsgBox ("���͂������x�R�[�h�́A���o�^�ł��B�ē��͂��ĉ������B")
                Text1(ptxSYUSHI_CODE).SetFocus
                Exit Function
            End If
'>>>>>>>>>>>>>> 2019.04.04
        
        
        
        
        
        
        Case Mode_All, ptxUKEHARAI_NAME     '�󕥐於��
        
            If G_SCREEN_FLG = G_SCREEN_INS Then
                Text1(ptxUKEHARAI_RNAME).Text = Text1(ptxUKEHARAI_NAME).Text
                    
            End If
        
        
        Case Mode_All, ptxUKEHARAI_RNAME    '�󕥐旪��
        Case Mode_All, ptxBUSHO_NAME        '�����^�c�Ə�
        Case Mode_All, ptxTEL_NO            '�d�b�ԍ�
        Case Mode_All, ptxFAX_NO            'FAX�ԍ�
        Case Mode_All, ptxYUBIN_NO          '�X�֔ԍ�
        Case Mode_All, ptxADDR1             '�Z���P
        Case Mode_All, ptxADDR2             '�Z���Q
              
        
    End Select
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    '�󕥐�Ͻ��ǂݍ���
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, CODE)
    
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
            'ں��ޓ��e�̕\��
                                            '�󕥐溰��
            Text1(ptxUKEHARAI_CODE).Text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode))
                                            '���x����
            Text1(ptxSYUSHI_CODE).Text = Trim(StrConv(P_UKEHARAIREC.SYUSHI_CODE, vbUnicode))
                                            '���x������
            For i = 0 To Combo1(pcmbSYUSHI).ListCount - 1
                If Right(Combo1(pcmbSYUSHI).List(i), 3) = Text1(ptxSYUSHI_CODE).Text Then
                    Combo1(pcmbSYUSHI).ListIndex = i
                    Exit For
                End If
            
            Next i
                                            '�󕥐於��
            Text1(ptxUKEHARAI_NAME).Text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode))
                                            '�󕥐旪��
            Text1(ptxUKEHARAI_RNAME).Text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))
                                            '�����^�c�Ə���
            Text1(ptxBUSHO_NAME).Text = Trim(StrConv(P_UKEHARAIREC.BUSHO_NAME, vbUnicode))
                                            '�d�b�ԍ�
            Text1(ptxTEL_NO).Text = Trim(StrConv(P_UKEHARAIREC.TEL_NO, vbUnicode))
                                            'FAX�ԍ�
            Text1(ptxFAX_NO).Text = Trim(StrConv(P_UKEHARAIREC.FAX_NO, vbUnicode))
                                            '�X�֔ԍ�
            Text1(ptxYUBIN_NO).Text = Trim(StrConv(P_UKEHARAIREC.YUBIN_NO, vbUnicode))
                                            '�Z���P
            Text1(ptxADDR1).Text = Trim(StrConv(P_UKEHARAIREC.ADDR1, vbUnicode))
                                            '�Z���Q
            Text1(ptxADDR2).Text = Trim(StrConv(P_UKEHARAIREC.ADDR2, vbUnicode))
                                            '�����敪
            Combo1(pcmbTORI_KBN).ListIndex = 0
            For i = 0 To Combo1(pcmbTORI_KBN).ListCount - 1
                If Right(Combo1(pcmbTORI_KBN).List(i), 1) = StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) Then
                    Combo1(pcmbTORI_KBN).ListIndex = i
                    Exit For
                End If
            
            Next i
        
        
        Case BtErrKeyNotFound
        
            MsgBox "���[���ŕύX����Ă��܂��B�O��ʂɖ߂�܂��B"
            PM000702.Visible = False
            INIT_FLG = False
            
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
            PM000702.Visible = False
            INIT_FLG = False
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^�o��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

    Update_Proc = True
    
    '�󕥐�}�X�^�@�ǂݍ���
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE))
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_UKEHARAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�󕥐�}�X�^")
                Exit Function
        
        End Select
    
    
    Loop
    '--------------------------------------------------------���R�[�h���e�ҏW
    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE).Text)        '�󕥐溰��
    Call UniCode_Conv(P_UKEHARAIREC.SYUSHI_CODE, Text1(ptxSYUSHI_CODE).Text)            '���x����
    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, Text1(ptxUKEHARAI_NAME).Text)        '�󕥐於
    Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, Text1(ptxUKEHARAI_RNAME).Text)      '�󕥗���
    Call UniCode_Conv(P_UKEHARAIREC.BUSHO_NAME, Text1(ptxBUSHO_NAME).Text)              '�����^�c�Ə�
    Call UniCode_Conv(P_UKEHARAIREC.TEL_NO, Text1(ptxTEL_NO).Text)                      '�d�b�ԍ�
    Call UniCode_Conv(P_UKEHARAIREC.FAX_NO, Text1(ptxFAX_NO).Text)                      'FAX�ԍ�
    Call UniCode_Conv(P_UKEHARAIREC.YUBIN_NO, Text1(ptxYUBIN_NO).Text)                  '�X�֔ԍ�
    Call UniCode_Conv(P_UKEHARAIREC.ADDR1, Text1(ptxADDR1).Text)                        '�Z���P
    Call UniCode_Conv(P_UKEHARAIREC.ADDR2, Text1(ptxADDR2).Text)                        '�Z���Q
    Call UniCode_Conv(P_UKEHARAIREC.TORI_KBN, Right(Combo1(pcmbTORI_KBN).Text, 1))      '�����敪
    
    
    
    
    Call UniCode_Conv(P_UKEHARAIREC.FILLER, "")                                         'Filler
    
    Call UniCode_Conv(P_UKEHARAIREC.UPD_TANTO, "")                                      '�X�V�S���Һ���
                                                                                        '�X�V����
    Call UniCode_Conv(P_UKEHARAIREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
    
    Do
        
        DoEvents
        
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_UKEHARAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "�󕥐�}�X�^")
                Exit Function
        End Select
    
    Loop
    
    
    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^�폜
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    Delete_Proc = True
    
    '�󕥐�}�X�^�@�ǂݍ���
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, Text1(ptxUKEHARAI_CODE))
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_UKEHARAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�󕥐�}�X�^")
                Exit Function
        
        End Select
    
    
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_UKEHRAI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "�󕥐�}�X�^")
                Exit Function
        End Select
    Loop


    Delete_Proc = False


End Function

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbSYUSHI     '���x
            Text1(ptxSYUSHI_CODE).Text = Right(Combo1(pcmbSYUSHI).Text, 3)
    
    End Select
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub


Private Sub Command1_Click(Index As Integer)

Dim yn As Integer

    Select Case Index
        Case P_CMD_Upd                      '�X�V
            If Error_Check_Proc(0) Then     '�G���[�`�F�b�N
                Exit Sub
            End If
            
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Update_Proc() Then
                    PM000702.Visible = False
                    INIT_FLG = False
                End If
            End If
            PM000702.Visible = False
            INIT_FLG = False
                    
        
        
        Case P_CMD_DEL                      '�폜
            yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Delete_Proc() Then
                    PM000702.Visible = False
                    INIT_FLG = False
                End If
            End If
            PM000702.Visible = False
            INIT_FLG = False
        Case P_CMD_DSP                      '����/�\��
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        
        Case P_CMD_End                      '�I��
            PM000702.Visible = False
            INIT_FLG = False
    End Select

End Sub

Private Sub Form_Activate()
    
Dim i       As Integer
Dim CODE    As String
    
    If INIT_FLG Then
        Exit Sub
    End If


    Select Case G_SCREEN_FLG
        Case G_SCREEN_INS       '�V�K
                
            Text1(ptxUKEHARAI_CODE).BackColor = G_INPUT_OK
            Text1(ptxUKEHARAI_CODE).TabStop = True
            Text1(ptxUKEHARAI_CODE).Locked = False
                
            For i = ptxUKEHARAI_CODE To ptxADDR2
                Text1(i).Text = ""
            Next i
            
            Combo1(pcmbTORI_KBN).ListIndex = 0
                
            Text1(ptxUKEHARAI_CODE).SetFocus
                
        
        Case G_SCREEN_UPD       '�X�V
    
                
    
    
            Text1(ptxUKEHARAI_CODE).BackColor = G_INPUT_NG
            Text1(ptxUKEHARAI_CODE).TabStop = False
            Text1(ptxUKEHARAI_CODE).Locked = True
    
            
            CODE = PM000701.txSEL_KEY
            
            If Item_Disp_Proc(CODE) Then
                Unload Me
            End If
    
            Text1(ptxSYUSHI_CODE).SetFocus
    
    End Select


    INIT_FLG = True

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

Dim com     As Integer
Dim sts     As Integer

    
    PM000702.Caption = PM000702.Caption & LAST_UPDATE_DAY
    
    
    '���x���e�̃Z�b�g
    If Code_Set_Proc(pcmbSYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
    
    '�����敪
    Combo1(pcmbTORI_KBN).Clear
    Combo1(pcmbTORI_KBN).AddItem P_TORI_GENERAL_N & "    " & P_TORI_GENERAL
    Combo1(pcmbTORI_KBN).AddItem P_TORI_NAISYOKU_N & "    " & P_TORI_NAISYOKU
    Combo1(pcmbTORI_KBN).AddItem P_TORI_GENKIN_N & "    " & P_TORI_GENKIN
    Combo1(pcmbTORI_KBN).AddItem P_TORI_SYANAI_N & "    " & P_TORI_SYANAI
    Combo1(pcmbTORI_KBN).AddItem P_TORI_ANOTHER_N & "    " & P_TORI_ANOTHER
    
    Combo1(pcmbTORI_KBN).AddItem P_TORI_JIKYU_N & "    " & P_TORI_JIKYU
    
    
    
    
    INIT_FLG = False
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
                                            
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000701 = Nothing
    Set PM000702 = Nothing

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
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case ptxUKEHARAI_CODE
            Text1(Index).Text = StrConv(RTrim(Text1(Index).Text), vbUpperCase)
    End Select
        
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    If Mode = 1 Then
        Combo1(Index).AddItem Space(Key_Len)
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    




End Function


