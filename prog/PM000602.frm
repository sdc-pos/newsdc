VERSION 5.00
Begin VB.Form PM000602 
   Caption         =   "���i���V�X�e���@�N���X�}�X�^�����e�i���X"
   ClientHeight    =   4530
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   11715
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
   ScaleHeight     =   4530
   ScaleWidth      =   11715
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7560
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   9960
      MaxLength       =   11
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   7560
      MaxLength       =   11
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5100
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   2940
      MaxLength       =   11
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   2040
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2805
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   1
      Left            =   4680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   840
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2535
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
      TabIndex        =   19
      Top             =   3600
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   11
      Top             =   3600
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3600
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
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�d���H��"
      Height          =   255
      Index           =   5
      Left            =   6405
      TabIndex        =   26
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Label Label 
      Caption         =   "���̑�"
      Height          =   255
      Index           =   4
      Left            =   9120
      TabIndex        =   25
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "����H��"
      Height          =   255
      Index           =   3
      Left            =   6405
      TabIndex        =   24
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Label Label 
      Caption         =   "�H��"
      Height          =   255
      Index           =   2
      Left            =   4500
      TabIndex        =   23
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "���i�����i"
      Height          =   255
      Index           =   1
      Left            =   1620
      TabIndex        =   22
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "�d������"
      Height          =   255
      Index           =   9
      Left            =   960
      TabIndex        =   21
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "�N���X"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   20
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "PM000602"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�e�L�X�g�p�Y��
Private Const ptxCLASS_CODE% = 0            '�N���X�R�[�h
Private Const ptxCLASS_NAME% = 1            '�Ăі�

Private Const ptxTANKA% = 2                 '���i�����i
Private Const ptxKOUSU% = 3                 '�H��
Private Const ptxKOURYOU% = 4               '�H��
Private Const ptxETC% = 5                   '���̑�

Private Const ptxURI_KOURYOU% = 6           '����H��   2007.01.11



'�R���{�p�Y��
Private Const pcmbSHIMUKE% = 0              '�d������

Private INIT_FLG    As Boolean

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000602.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000602)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000602)


    PM000602.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim com     As Integer
Dim ans     As Integer
Dim sts     As Integer

Dim i       As Integer
    
    Error_Check_Proc = True
    
    
    Select Case Mode
        
        Case ptxCLASS_CODE         '�i��
            
            
            If Trim(Text1(ptxCLASS_CODE).Text) = "" Then
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxCLASS_CODE).SetFocus
                Exit Function
            End If
            
        
            If G_SCREEN_FLG = G_SCREEN_INS And _
                Not Text1(ptxCLASS_CODE).Locked Then
                '�V�K���͏d���`�F�b�N
                
                Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxCLASS_CODE).Text)
                sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                Select Case sts
                    Case BtNoErr
                        ans = MsgBox("���͂����R�[�h�́A�o�^�ςł��B�X�V�����Ƃ��Čp�����܂����H", vbYesNo, "�m�F����")
                        If ans = vbNo Then
                            Text1(ptxCLASS_CODE).SetFocus
                            Exit Function
                        End If
                
                        Call Item_Disp_Proc(Right(Combo1(pcmbSHIMUKE), 3) & Text1(ptxCLASS_CODE).Text)
                    
                    
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�N���X�}�X�^")
                        Exit Function
                End Select
            
            
                Combo1(pcmbSHIMUKE).BackColor = G_INPUT_NG
                Combo1(pcmbSHIMUKE).Locked = True
                Combo1(pcmbSHIMUKE).TabStop = False
            
            
                Text1(ptxCLASS_CODE).BackColor = G_INPUT_NG
                Text1(ptxCLASS_CODE).Locked = True
                Text1(ptxCLASS_CODE).TabStop = False
            
            
            End If
        
        Case ptxTANKA              '���i���P��
        
        
            If Text1(ptxTANKA).Text = "" Then
                Text1(ptxTANKA).Text = "0"
            End If
        
        
            If Not IsNumeric(Text1(ptxTANKA).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxTANKA).SetFocus
                    Exit Function
            Else
                
                Text1(ptxTANKA).Text = Format(CDbl(Text1(ptxTANKA).Text), "#0.00")
                
                If CDbl(Text1(ptxTANKA).Text) < 0 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxTANKA).SetFocus
                    Exit Function
                End If
            
            End If
        
        Case ptxKOUSU          '�H��
        
            If Text1(ptxKOUSU).Text = "" Then
                Text1(ptxKOUSU).Text = "0"
            End If
        
        
            If Not IsNumeric(Text1(ptxKOUSU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxKOUSU).SetFocus
                    Exit Function
            Else
                
                Text1(ptxKOUSU).Text = Format(CDbl(Text1(ptxKOUSU).Text), "#0.000")
                
                If CDbl(Text1(ptxKOUSU).Text) < 0 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxKOUSU).SetFocus
                    Exit Function
                End If
            
            End If
        
        
        Case ptxKOURYOU           '�H��
        
            If Text1(ptxKOURYOU).Text = "" Then
                Text1(ptxKOURYOU).Text = "0"
            End If
        
        
            If Not IsNumeric(Text1(ptxKOURYOU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxKOURYOU).SetFocus
                    Exit Function
            Else
                
                Text1(ptxKOURYOU).Text = Format(CDbl(Text1(ptxKOURYOU).Text), "#0.00")
                
                If CDbl(Text1(ptxKOURYOU).Text) < 0 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxKOURYOU).SetFocus
                    Exit Function
                End If
            
            End If
        
        Case ptxETC             '���̑�
        
            If Text1(ptxETC).Text = "" Then
                Text1(ptxETC).Text = "0"
            End If
        
        
            If Not IsNumeric(Text1(ptxETC).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxETC).SetFocus
                    Exit Function
            Else
                
                Text1(ptxETC).Text = Format(CDbl(Text1(ptxETC).Text), "#0.00")
                
                If CDbl(Text1(ptxETC).Text) < 0 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxETC).SetFocus
                    Exit Function
                End If
            
            End If
        
        Case ptxURI_KOURYOU           '����H�� 2007.01.11
        
            If Text1(ptxURI_KOURYOU).Text = "" Then
                Text1(ptxURI_KOURYOU).Text = "0"
            End If
        
        
            If Not IsNumeric(Text1(ptxURI_KOURYOU).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxKOURYOU).SetFocus
                    Exit Function
            Else
                
                Text1(ptxURI_KOURYOU).Text = Format(CDbl(Text1(ptxURI_KOURYOU).Text), "#0.00")
                
                If CDbl(Text1(ptxURI_KOURYOU).Text) < 0 Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxKOURYOU).SetFocus
                    Exit Function
                End If
            
            End If
        
        
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
    
    '�׽Ͻ��ǂݍ���
    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(CODE, 1, 2))
    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Mid(CODE, 3, 20))
    
    
    sts = BTRV(BtOpGetGreaterEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    Select Case sts
        Case BtNoErr
            
            
            'ں��ޓ��e�̕\��
            For i = 0 To Combo1(pcmbSHIMUKE).ListCount - 1
            
                If StrConv(P_CLASSREC.SHIMUKE_CODE, vbUnicode) = Left(Right(Combo1(pcmbSHIMUKE).List(i), 4), 2) Then
            
                    Combo1(pcmbSHIMUKE).ListIndex = i
                    
                    Exit For
            
                End If
            
            Next
                                            '�׽����
            Text1(ptxCLASS_CODE).Text = Trim(StrConv(P_CLASSREC.CLASS_CODE, vbUnicode))
                                            '�Ăі�
            Text1(ptxCLASS_NAME).Text = Trim(StrConv(P_CLASSREC.CLASS_NAME, vbUnicode))
                                            '���i�����i
            Text1(ptxTANKA).Text = Format(CDbl(StrConv(P_CLASSREC.TANKA, vbUnicode)), "#0.00")
                                            '�H��
            Text1(ptxKOUSU).Text = Format(CDbl(StrConv(P_CLASSREC.KOUSU, vbUnicode)), "#0.000")
                                            '�H��
            Text1(ptxKOURYOU).Text = Format(CDbl(StrConv(P_CLASSREC.KOURYOU, vbUnicode)), "#0.00")
                                            '���̑�
            Text1(ptxETC).Text = Format(CDbl(StrConv(P_CLASSREC.ETC, vbUnicode)), "#0.00")
        
                                            '����H��
            If IsNumeric(StrConv(P_CLASSREC.URI_KOURYOU, vbUnicode)) Then
                Text1(ptxURI_KOURYOU).Text = Format(CDbl(StrConv(P_CLASSREC.URI_KOURYOU, vbUnicode)), "#0.00")
            Else
                Text1(ptxURI_KOURYOU).Text = "0.00"
            End If
        
        
        Case BtErrKeyNotFound
        
            MsgBox "���[���ŕύX����Ă��܂��B�O��ʂɖ߂�܂��B"
            PM000602.Visible = False
            INIT_FLG = False
            
            Exit Function
                    
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�N���X�}�X�^")
            PM000602.Visible = False
            INIT_FLG = False
            Exit Function
    
    End Select

    Item_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �N���X�}�X�^�o��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
Dim i       As Integer

    Update_Proc = True
    
    
    
    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxCLASS_CODE).Text)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CLASS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�N���X�}�X�^")
                Exit Function
        
        End Select
    
    
    Loop
    
    
    If com = BtOpInsert Then
        Call UniCode_Conv(P_CLASSREC.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
        Call UniCode_Conv(P_CLASSREC.CLASS_CODE, Text1(ptxCLASS_CODE).Text)
    
        Call UniCode_Conv(P_CLASSREC.FILLER, "")
    
    
    End If


    Call UniCode_Conv(P_CLASSREC.CLASS_NAME, Text1(ptxCLASS_NAME).Text)                         '�Ăі�
    Call UniCode_Conv(P_CLASSREC.TANKA, Format(CDbl(Text1(ptxTANKA).Text), "00000000.00"))      '���i�����i
    Call UniCode_Conv(P_CLASSREC.KOUSU, Format(CDbl(Text1(ptxKOUSU).Text), "000.000"))          '�H��
    Call UniCode_Conv(P_CLASSREC.KOURYOU, Format(CDbl(Text1(ptxKOURYOU).Text), "00000000.00"))  '�H��
    Call UniCode_Conv(P_CLASSREC.ETC, Format(CDbl(Text1(ptxETC).Text), "00000000.00"))          '���̑�
                                                                                                '����H��   2007.01.11
    Call UniCode_Conv(P_CLASSREC.URI_KOURYOU, Format(CDbl(Text1(ptxURI_KOURYOU).Text), "00000000.00"))

    Call UniCode_Conv(P_CLASSREC.UPD_TANTO, "")                                                 '�X�V�S���Һ���
    Call UniCode_Conv(P_CLASSREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS")) '�X�V����

    Do
        
        DoEvents
        
        sts = BTRV(com, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CLASS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, com, "�N���X�}�X�^")
                Exit Function
        End Select
    
    Loop
    
    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   �N���X�}�X�^�폜
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer

    Delete_Proc = True
    
    
    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))
    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxCLASS_CODE).Text)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CLASS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�N���X�}�X�^")
                Exit Function
        
        End Select

    Loop


    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CLASS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "�N���X�}�X�^")
                Exit Function
        End Select
    Loop


    Delete_Proc = False


End Function



Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub


Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer
Dim i       As Integer

    Select Case Index
        Case P_CMD_Upd                      '�X�V
            
            For i = ptxCLASS_CODE To ptxURI_KOURYOU
                If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                    Exit Sub
                End If
            Next i
            
            Beep
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    PM000602.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
                                
            PM000602.Visible = False
            INIT_FLG = False
                                
                                
        
        Case P_CMD_DEL                      '�폜
            ans = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Delete_Proc() Then
                    PM000602.Visible = False
                    INIT_FLG = False
                End If
            Else
                Exit Sub
            End If
        
        
            PM000602.Visible = False
            INIT_FLG = False
        
        
        Case P_CMD_DSP                      '����/�\��
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        
        Case P_CMD_End                      '�I��
            PM000602.Visible = False
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
                
                
                
                
                
            Combo1(pcmbSHIMUKE).BackColor = G_INPUT_OK
            Combo1(pcmbSHIMUKE).TabStop = True
            Combo1(pcmbSHIMUKE).Locked = False
                
                
            Text1(ptxCLASS_CODE).BackColor = G_INPUT_OK
            Text1(ptxCLASS_CODE).TabStop = True
            Text1(ptxCLASS_CODE).Locked = False
                
            For i = ptxCLASS_CODE To ptxKOURYOU
                Text1(i).Text = ""
            Next i
                
                
            Combo1(pcmbSHIMUKE).SetFocus
            Combo1(pcmbSHIMUKE).ListIndex = 0
                
                
                
        
        Case G_SCREEN_UPD       '�X�V
    
            Combo1(pcmbSHIMUKE).BackColor = G_INPUT_NG
            Combo1(pcmbSHIMUKE).TabStop = False
            Combo1(pcmbSHIMUKE).Locked = True
                
    
    
            Text1(ptxCLASS_CODE).BackColor = G_INPUT_NG
            Text1(ptxCLASS_CODE).TabStop = False
            Text1(ptxCLASS_CODE).Locked = True
    
    
            
            CODE = PM000601.txSEL_KEY.Text
            
            If Item_Disp_Proc(CODE) Then
                Exit Sub
            End If
    
            Text1(ptxCLASS_NAME).SetFocus
    
    End Select


    INIT_FLG = True

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

    
    '�d�����於�̃Z�b�g
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD) Then
        Unload Me
    End If
    
    
    INIT_FLG = False
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
                                            '�N���X�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�N���X�}�X�^")
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
    Set PM000601 = Nothing
    Set PM000602 = Nothing

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
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Function Code_Set_Proc(Index As Integer, KBN As String) As Integer
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

