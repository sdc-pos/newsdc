VERSION 5.00
Begin VB.Form F1011001 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�I�ԕʌ���������N�ݒ�"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
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
   ScaleHeight     =   7200
   ScaleWidth      =   11385
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   840
      Width           =   2052
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I�@��"
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6600
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
      Index           =   8
      Left            =   7800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�}�b�v"
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�Q ��"
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6600
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
      Index           =   3
      Left            =   2640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6600
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Caption         =   "�ǁ^��"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6600
      Width           =   855
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H000000FF&
      Height          =   5775
      Left            =   5520
      TabIndex        =   29
      Top             =   360
      Width           =   5655
      Begin VB.TextBox Text 
         Height          =   375
         Index           =   6
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "�`�a�b�����N"
         Height          =   4695
         Left            =   600
         TabIndex        =   30
         Top             =   720
         Width           =   4575
         Begin VB.OptionButton Option1 
            Caption         =   "�`�|�P"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�`�|�Q"
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�a�|�P"
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�a�|�Q"
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�b�|�P"
            Height          =   495
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   2400
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�b�|�Q"
            Height          =   495
            Index           =   5
            Left            =   240
            TabIndex        =   11
            Top             =   2880
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�c"
            Height          =   495
            Index           =   6
            Left            =   240
            TabIndex        =   12
            Top             =   3360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "�d"
            Height          =   495
            Index           =   7
            Left            =   240
            TabIndex        =   13
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label ABC_KBN 
            Height          =   375
            Index           =   0
            Left            =   1560
            TabIndex        =   38
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label ABC_KBN 
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   37
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label ABC_KBN 
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   36
            Top             =   1560
            Width           =   2775
         End
         Begin VB.Label ABC_KBN 
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   35
            Top             =   2040
            Width           =   2775
         End
         Begin VB.Label ABC_KBN 
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   34
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label ABC_KBN 
            Height          =   375
            Index           =   5
            Left            =   1560
            TabIndex        =   33
            Top             =   3000
            Width           =   2775
         End
         Begin VB.Label ABC_KBN 
            Caption         =   "�ߋ��Q�N�ԂŎ��т���"
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   32
            Top             =   3480
            Width           =   2775
         End
         Begin VB.Label ABC_KBN 
            Caption         =   "�s�ړ�"
            Height          =   375
            Index           =   7
            Left            =   1560
            TabIndex        =   31
            Top             =   3960
            Width           =   2775
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         Caption         =   "������"
         Height          =   375
         Index           =   11
         Left            =   600
         TabIndex        =   39
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   42
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      Caption         =   "�A"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   41
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   40
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   27
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      Caption         =   "�q��"
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   26
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "F1011001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSoko_No% = 0       '�q�ɇ�
Private Const ptxSoko_Name% = 1     '�q�ɖ���
Private Const ptxstRetu% = 2        '��
Private Const ptxenRetu% = 3        '��
Private Const ptxstRen% = 4         '�A
Private Const ptxenRen% = 5         '�A
Private Const ptxPacking_No% = 6    '������

Private Const Text_Max% = 6

Private Const ABC_KBN_MAX% = 7

Private Sub Command_Click(Index As Integer)
    
Dim i       As Integer
Dim yn      As Integer
Dim sts     As Integer
    
    
    Select Case Index
        
        Case 0          '�ǉ��^����
            
            sts = Err_Check_Proc(0)
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("[�ǉ��^����]���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                sts = Update_Proc
                Select Case sts
                    Case False
                    Case True
                        Unload Me
                    Case SYS_CANCEL
                End Select
                
            End If
            
            Text(ptxstRetu).SetFocus
        
        Case 3          '�폜
            sts = Err_Check_Proc(1)
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            Beep
            yn = MsgBox("[�폜]���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                sts = Delete_Proc
                Select Case sts
                    Case False
                    Case True
                        Unload Me
                    Case SYS_CANCEL
                End Select
                
            End If
            
            Text(ptxstRetu).SetFocus
        
        Case 4          '�Q��
        
            If Text(ptxSoko_No).Text = "" Then
            
                Beep
                MsgBox "���͂������ڂ̓G���[�ł�� �i�K�{���́j"
                Text(ptxSoko_No).SetFocus
                Exit Sub
            
            End If
        
            F1011002.Text(0).Text = Text(ptxSoko_No).Text
            F1011002.Text(1).Text = Text(ptxSoko_Name).Text
            F1011002.Text(2).Text = Text(ptxstRetu).Text
            F1011002.Text(3).Text = Text(ptxenRetu).Text
            F1011002.Text(4).Text = Text(ptxstRen).Text
            F1011002.Text(5).Text = Text(ptxenRen).Text
        
            F1011002.Text(6).Text = ""
        
            F1011002.Show vbModal
        
        
        Case 7          '�}�b�v
            
            F1011003.Text(0).Text = ""
            F1011003.Text(1).Text = ""
            F1011003.Text(2).Text = ""
            
            For i = 0 To 19
                F1011003.Text1(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text2(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text3(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text4(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text5(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text6(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text7(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text8(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text9(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text10(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text11(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text12(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text13(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text14(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text15(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text16(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text17(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text18(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text19(i).BackColor = F1011003.Label5(8).BackColor
                F1011003.Text20(i).BackColor = F1011003.Label5(8).BackColor
            Next i
            
            
            F1011003.Show vbModal
        
        Case 11         '�I��
            
            Unload Me
    
    End Select

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

Dim c       As String * 128
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
                                
                                
    If Not GetIni("F101100", "COLOR1", "SYS", c) Then
        F1011003.Label5(0).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR2", "SYS", c) Then
        F1011003.Label5(1).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR3", "SYS", c) Then
        F1011003.Label5(2).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR4", "SYS", c) Then
        F1011003.Label5(3).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR5", "SYS", c) Then
        F1011003.Label5(4).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR6", "SYS", c) Then
        F1011003.Label5(5).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR7", "SYS", c) Then
        F1011003.Label5(6).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR8", "SYS", c) Then
        F1011003.Label5(7).BackColor = CLng(RTrim(c))
    End If
    
    If Not GetIni("F101100", "COLOR9", "SYS", c) Then
        F1011003.Label5(8).BackColor = CLng(RTrim(c))
    End If
    
                                
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����}�X�^�n�o�d�m
    If PACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�ʌ����}�X�^�n�o�d�m
    If TPACKING_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
                                            '�����}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
                                            '�I�ʌ����}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�ʌ����}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    
    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If
    
    Set F1011001 = Nothing

    End

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i           As Integer
Dim j           As Integer
Dim sts         As Integer
Dim Edit        As String


    If KeyCode <> vbKeyReturn Then Exit Sub
            
    Select Case Index
        
        Case ptxSoko_No
        
            Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSoko_No))
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                                
                    If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_JITU Then
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł�� "
                        Text(Index).SetFocus
                        Exit Sub
                    End If
                
                    Text(ptxSoko_Name).Text = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                
                Case BtErrKeyNotFound
            
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� �i���o�^�j"
                    Text(Index).SetFocus
                    Exit Sub
            
                Case Else
                    
                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                    Unload Me
            
            End Select
        
        Case ptxstRetu, ptxstRen, ptxenRetu, ptxenRen
            If Len(Trim(Text(Index).Text)) = 0 Then
            Else
                If Not IsNumeric(Text(Index).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� "
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
    
            If Index = ptxenRetu Then
                If Len(Trim(Text(Index).Text)) = 0 Then
                    Text(Index).Text = "99"
                End If
            End If
            If Index = ptxenRen Then
                If Len(Trim(Text(Index).Text)) = 0 Then
                    Text(Index).Text = "99"
                End If
            End If
    
    
            If Index = ptxenRen Then
                
                If Text(ptxstRetu).Text & Text(ptxstRen).Text > Text(ptxenRetu).Text & Text(ptxenRen).Text Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� "
                    Text(ptxstRetu).SetFocus
                    Exit Sub
                End If
                
            End If
    
        Case ptxPacking_No
                                        '������
            Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(ptxPacking_No).Text)
            sts = BTRV(BtOpGetEqual, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
            
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� �i���o�^�j"
                    Text(Index).SetFocus
                    Exit Sub
            
            
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�����}�X�^")
                    Unload Me
            End Select
    
            ABC_KBN(0).Caption = "[" & StrConv(Format(CLng(StrConv(PACKINGREC.RANK_A1, vbUnicode)), "#0"), vbWide) & "�`"
            ABC_KBN(1).Caption = "[" & StrConv(Format(CLng(StrConv(PACKINGREC.RANK_A2, vbUnicode)), "#0"), vbWide) & "�`"
            ABC_KBN(2).Caption = "[" & StrConv(Format(CLng(StrConv(PACKINGREC.RANK_B1, vbUnicode)), "#0"), vbWide) & "�`"
            ABC_KBN(3).Caption = "[" & StrConv(Format(CLng(StrConv(PACKINGREC.RANK_B2, vbUnicode)), "#0"), vbWide) & "�`"
            ABC_KBN(4).Caption = "[" & StrConv(Format(CLng(StrConv(PACKINGREC.RANK_C1, vbUnicode)), "#0"), vbWide) & "�`"
            ABC_KBN(5).Caption = "[" & StrConv(Format(CLng(StrConv(PACKINGREC.RANK_C2, vbUnicode)), "#0"), vbWide) & "�`"
    
    End Select
    
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Function Update_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Upd         As Integer

Dim ans         As Integer
        
Dim WK_Soko_No  As String * 2
Dim WK_Retu     As String * 2
Dim WK_Ren      As String * 2
    
Dim i           As Integer
    
    Update_Proc = True
    
    Call Input_Lock
    Frame2.Enabled = False
    
    Call UniCode_Conv(K0_TANA.Soko_No, Text(ptxSoko_No).Text)
    Call UniCode_Conv(K0_TANA.Retu, Text(ptxstRetu).Text)
    Call UniCode_Conv(K0_TANA.Ren, Text(ptxstRen).Text)
    Call UniCode_Conv(K0_TANA.Dan, "")
    
    
    com = BtOpGetGreaterEqual
    WK_Soko_No = ""
    
    Do
        DoEvents
        sts = BTRV(com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(TANAREC.Soko_No, vbUnicode) <> Text(ptxSoko_No).Text) Or _
                   (StrConv(TANAREC.Retu, vbUnicode) > Text(ptxenRetu).Text) Then
                    If WK_Soko_No <> "" Then
                    
                        sts = TPACKING_Update_Proc(WK_Soko_No, WK_Retu, WK_Ren)
                        
                        Select Case sts
                            Case False
                            Case True
                                Exit Function
                            Case SYS_CANCEL
                                Update_Proc = False
                                Exit Function
                        End Select
                    
                    End If
                    
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Exit Function
        End Select
    
    
        If (StrConv(TANAREC.Ren, vbUnicode) < Text(ptxstRen).Text) Or _
           (StrConv(TANAREC.Ren, vbUnicode) > Text(ptxenRen).Text) Then
    
        Else
            If com = BtOpGetGreaterEqual Then
            
                WK_Soko_No = StrConv(TANAREC.Soko_No, vbUnicode)
                WK_Retu = StrConv(TANAREC.Retu, vbUnicode)
                WK_Ren = StrConv(TANAREC.Ren, vbUnicode)
        
            End If
    
            If (WK_Soko_No <> StrConv(TANAREC.Soko_No, vbUnicode) Or _
                WK_Retu <> StrConv(TANAREC.Retu, vbUnicode) Or _
                WK_Ren <> StrConv(TANAREC.Ren, vbUnicode)) Then
            
                sts = TPACKING_Update_Proc(WK_Soko_No, WK_Retu, WK_Ren)
                Select Case sts
                    Case False
                    Case True
                        Exit Function
                    Case SYS_CANCEL
                        Update_Proc = False
                        Exit Function
                End Select
        
            End If
    
            WK_Soko_No = StrConv(TANAREC.Soko_No, vbUnicode)
            WK_Retu = StrConv(TANAREC.Retu, vbUnicode)
            WK_Ren = StrConv(TANAREC.Ren, vbUnicode)
        
        End If
        
        com = BtOpGetNext
    
    Loop
    
    
    
    Text(ptxPacking_No).Text = ""
    
    For i = 0 To ABC_KBN_MAX
        Option1(i).Value = False
    Next i

    Update_Proc = False
    
    Beep
    MsgBox "�������ݏ���������ɏI�����܂����B"
    
    Frame2.Enabled = True
    Call Input_UnLock


End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1011001.MousePointer = vbHourglass

    Call Ctrl_Lock(F1011001)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1011001)


    F1011001.MousePointer = vbDefault

End Sub



Private Function Err_Check_Proc(Mode As Integer) As Integer
    
Dim sts As Integer
Dim i   As Integer
    
    Err_Check_Proc = True

    Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSoko_No))
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
                
            Text(ptxSoko_Name).Text = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                
        Case BtErrKeyNotFound
            
            Beep
            MsgBox "���͂������ڂ̓G���[�ł�� �i���o�^�j"
            Text(ptxSoko_No).SetFocus
            Exit Function
            
        Case Else
                    
            Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
            Err_Check_Proc = SYS_ERR
            Exit Function
    End Select

    If Text(ptxstRetu).Text = "" Then
        Text(ptxstRetu).Text = "99"
    End If

    If Text(ptxstRen).Text = "" Then
        Text(ptxstRen).Text = "99"
    End If


    If (Text(ptxstRetu).Text & Text(ptxstRen).Text) > (Text(ptxenRetu).Text & Text(ptxenRen).Text) Then
            
        Beep
        MsgBox "���͂������ڂ̓G���[�ł�� "
        Text(ptxstRetu).SetFocus
        Exit Function
            
    End If


    If Mode = 0 Then

        Call UniCode_Conv(K0_PACKING.PACKING_NO, Text(ptxPacking_No).Text)
        sts = BTRV(BtOpGetEqual, PACKING_POS, PACKINGREC, Len(PACKINGREC), K0_PACKING, Len(K0_PACKING), 0)
        Select Case sts
            Case BtNoErr

            Case BtErrKeyNotFound

                Beep
                MsgBox "���͂������ڂ̓G���[�ł�� �i���o�^�j"
                Text(ptxPacking_No).SetFocus
                Exit Function

            Case Else

                Call File_Error(sts, BtOpGetEqual, "�����}�X�^")
                Err_Check_Proc = SYS_ERR
                Exit Function
        End Select

    Else

        If Len(Trim(Text(ptxPacking_No).Text)) = 0 Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł�� �i���͕K�{�j"
            Text(ptxPacking_No).SetFocus
            Exit Function
        End If

    End If

    For i = 0 To ABC_KBN_MAX
        If Option1(i).Value Then
            Exit For
        End If
    Next i

    If i > ABC_KBN_MAX Then
        Beep
        MsgBox "�����N��I�����Ă���������i�I��K�{�j"
        Text(ptxPacking_No).SetFocus
        Exit Function
    End If


    Err_Check_Proc = False


End Function

Private Function TPACKING_Update_Proc(WK_Soko_No As String, WK_Retu As String, WK_Ren As String) As Integer

Dim sts     As Integer
Dim com     As Integer

Dim ans     As Integer

Dim i       As Integer


    TPACKING_Update_Proc = True


    Call UniCode_Conv(K0_TPACKING.Soko_No, WK_Soko_No)
    Call UniCode_Conv(K0_TPACKING.Retu, WK_Retu)
    Call UniCode_Conv(K0_TPACKING.Ren, WK_Ren)
    Call UniCode_Conv(K0_TPACKING.PACKING_NO, Text(ptxPacking_No).Text)
    For i = 0 To ABC_KBN_MAX
        If Option1(i).Value Then
            Call UniCode_Conv(K0_TPACKING.RANK, StrConv(Option1(i).Caption, vbNarrow))
            Exit For
        End If
    Next i

    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANAPACKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    TPACKING_Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                    
                Call File_Error(sts, BtOpGetEqual, "�I�ʌ����}�X�^")
                Exit Function
        
        End Select
    
    Loop


    Call UniCode_Conv(TPACKINGREC.Soko_No, WK_Soko_No)
    Call UniCode_Conv(TPACKINGREC.Retu, WK_Retu)
    Call UniCode_Conv(TPACKINGREC.Ren, WK_Ren)
    Call UniCode_Conv(TPACKINGREC.PACKING_NO, Text(ptxPacking_No).Text)
    For i = 0 To ABC_KBN_MAX
        If Option1(i).Value Then
            Call UniCode_Conv(TPACKINGREC.RANK, StrConv(Option1(i).Caption, vbNarrow))
            Exit For
        End If
    Next i

    Do
        
        sts = BTRV(com, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANAPACKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_PACKING), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "�I�ʌ����}�X�^")
                        Exit Function
                    End If
                    TPACKING_Update_Proc = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, com, "�I�ʌ����}�X�^")
                Exit Function
        End Select
    
    Loop


    TPACKING_Update_Proc = False


End Function

Private Function Delete_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Upd         As Integer

Dim ans         As Integer
Dim i           As Integer

Dim wkRANK      As String


    Delete_Proc = True


    Call Input_Lock
    Frame2.Enabled = False

    wkRANK = ""
    For i = 0 To ABC_KBN_MAX
        If Option1(i).Value Then
            wkRANK = StrConv(Option1(i).Caption, vbNarrow)
            Exit For
        End If
    Next i

    Call UniCode_Conv(K0_TPACKING.Soko_No, Text(ptxSoko_No).Text)
    Call UniCode_Conv(K0_TPACKING.Retu, Text(ptxstRetu).Text)
    Call UniCode_Conv(K0_TPACKING.Ren, Text(ptxstRen).Text)
    Call UniCode_Conv(K0_TPACKING.PACKING_NO, "")
    Call UniCode_Conv(K0_TPACKING.RANK, "")

    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        
        Do
            sts = BTRV(com + BtSNoWait, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANAPACKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Delete_Proc = SYS_CANCEL
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�I�ʌ����}�X�^")
                    Exit Function
            End Select

        Loop

        If sts = BtNoErr Then
            If (StrConv(TPACKINGREC.Soko_No, vbUnicode) <> Text(ptxSoko_No).Text) Or _
                ((StrConv(TPACKINGREC.Retu, vbUnicode) & StrConv(TPACKINGREC.Ren, vbUnicode)) > _
                (Text(ptxenRetu).Text & Text(ptxenRen).Text)) Then

                sts = BTRV(BtOpUnlock, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
                If sts Then
                    Call File_Error(sts, BtOpUnlock, "�I�ʌ����}�X�^")
                    Exit Function
                End If

                Exit Do

            End If
        Else
            Exit Do
        End If

        If StrConv(TPACKINGREC.PACKING_NO, vbUnicode) = Text(ptxPacking_No).Text And _
           Trim(StrConv(TPACKINGREC.RANK, vbUnicode)) = wkRANK Then
            Do
                sts = BTRV(BtOpDelete, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANAPACKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            sts = BTRV(BtOpUnlock, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_PACKING), 0)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "�I�ʌ����}�X�^")
                                Exit Function
                            End If
                            Delete_Proc = SYS_CANCEL
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "�I�ʌ����}�X�^")
                        Exit Function
                End Select
            Loop
        End If

        com = BtOpGetNext

    Loop



    Text(ptxPacking_No).Text = ""
    
    For i = 0 To ABC_KBN_MAX
        Option1(i).Value = False
    Next i


    Delete_Proc = False


    Beep
    MsgBox "�������ݏ���������ɏI�����܂����B"

    Frame2.Enabled = True
    Call Input_UnLock

End Function
