VERSION 5.00
Begin VB.Form F1020601 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���ח\��o�^"
   ClientHeight    =   6960
   ClientLeft      =   2550
   ClientTop       =   2715
   ClientWidth     =   12360
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
   ScaleHeight     =   6960
   ScaleWidth      =   12360
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   10800
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   17
      Left            =   10800
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   15
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   11
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   10
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   9
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   10800
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1200
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text 
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1560
      Width           =   2535
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   7
      Left            =   6480
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   4
      Left            =   3960
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   3
      Left            =   2640
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X  �V"
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "POS�݌�"
      Height          =   255
      Index           =   20
      Left            =   9840
      TabIndex        =   52
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "H�݌�"
      Height          =   255
      Index           =   19
      Left            =   9840
      TabIndex        =   51
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�����́F�������ԁj"
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   49
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TEXTNo"
      Height          =   255
      Index           =   18
      Left            =   360
      TabIndex        =   48
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�����́F�������ԁj"
      Height          =   255
      Index           =   17
      Left            =   2640
      TabIndex        =   47
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�S��"
      Height          =   255
      Index           =   14
      Left            =   480
      TabIndex        =   46
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblTanto_Name 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   1920
      TabIndex        =   45
      Top             =   240
      Width           =   2415
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
      TabIndex        =   44
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   16
      Left            =   6480
      TabIndex        =   43
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   13
      Left            =   5760
      TabIndex        =   42
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   12
      Left            =   5040
      TabIndex        =   41
      Top             =   4200
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�z�X�g�I��"
      Height          =   255
      Index           =   11
      Left            =   3240
      TabIndex        =   40
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�z�X�g�q��"
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   39
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�\�Z�P�ʁi��j"
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   38
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�\�Z�P�ʁi���j"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   37
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[��"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   36
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   35
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   34
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[���t"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ɐ�"
      Height          =   255
      Index           =   0
      Left            =   9840
      TabIndex        =   32
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���O"
      Height          =   255
      Index           =   15
      Left            =   600
      TabIndex        =   31
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   30
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i  ��"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   29
      Top             =   1680
      Width           =   735
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1020601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO As String

Private Const pcmbNAIGAI% = 0           '�����O

Private Const ptxTanto_Code% = 0        '�S���҃R�[�h
Private Const ptxID_NO% = 1             '�h�c��
Private Const ptxHin_Gai% = 2           '�i��
Private Const ptxHin_Name% = 3          '�i��
Private Const ptxYotei_Qty% = 4         '���ɐ�
Private Const ptxDEN_DT_YY% = 5         '�`�[���t�@�N
Private Const ptxDEN_DT_MM% = 6         '�`�[���t�@��
Private Const ptxDEN_DT_DD% = 7         '�`�[���t�@��
Private Const ptxDEN_NO% = 8            '�`�[��
Private Const ptxYOSAN_FROM% = 9        '�\�Z�P��FROM
Private Const ptxYOSAN_TO% = 10         '�\�Z�P��TO
Private Const ptxHOST_SOKO% = 11        '�z�X�g�q��
Private Const ptxST_SOKO% = 12          '�W���I�� �q��
Private Const ptxST_RETU% = 13          '�W���I�� ��
Private Const ptxST_REN% = 14           '�W���I�� �A
Private Const ptxST_DAN% = 15           '�W���I�� �i
    
Private Const ptxPOS_ZAIQTY% = 16       'POS�݌�
Private Const ptxHS_ZAIQTY% = 17        'νč݌�
    
    
Private Const Text_Max% = 17


Private MEMO_TEXT   As String           '��������
Private KASO_NYUKA  As String * 2       '���בq��
                                    
Private SOKO_GOODS_ON_F As String * 1
                                    
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer
    
    For i = Mode To Text_Max
        
        If i = 5 Or i = 6 Or i = 7 Then
        Else
            Text(i).Text = ""
        End If
    Next i
    
End Sub
                                    '�i�ڃ}�X�^���e���ڂ�\������
Private Function Item_Read_Proc() As Integer

Dim sts     As Integer
Dim NAIGAI  As String * 1

    Item_Read_Proc = True
                                                '�����O�̔���
    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                                
                                                '�܂��O���i�Ԃœǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxHin_Gai).Text)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            
            Text(ptxHin_Name).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            Text(ptxHOST_SOKO).Text = RTrim(StrConv(ITEMREC.BIKOU_SOKO, vbUnicode))
            Text(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
            Text(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
            Text(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
            Text(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
        
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
                    Text(ptxHOST_SOKO).Text = RTrim(StrConv(ITEMREC.BIKOU_SOKO, vbUnicode))
                    Text(ptxST_SOKO).Text = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    Text(ptxST_RETU).Text = StrConv(ITEMREC.ST_RETU, vbUnicode)
                    Text(ptxST_REN).Text = StrConv(ITEMREC.ST_REN, vbUnicode)
                    Text(ptxST_DAN).Text = StrConv(ITEMREC.ST_DAN, vbUnicode)
        
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

'                                       ���͍��ڂ̃G���[�`�F�b�N
Private Function Err_Chk() As Integer
            
Dim sts     As Integer
Dim i       As Integer
Dim RetBuf  As String
Dim c       As String * 128

    Err_Chk = True
                                        '�S���҂̃`�F�b�N
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTanto_Code).Text)
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            lblTanto_Name.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Beep
            MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
            Text(ptxTanto_Code).SetFocus
            Exit Function
        Case Else
           Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Err_Chk = SYS_ERR
            Exit Function
    End Select
                                        '�i�ԃ`�F�b�N
    sts = Item_Read_Proc()
    Select Case sts
        Case False
        Case True
            Beep
            MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
            Text(ptxHin_Gai).SetFocus
            Exit Function
        Case Else
            Err_Chk = sts
            Exit Function
    End Select
                                        '���ɐ��ʃ`�F�b�N
    If Len(Text(ptxYotei_Qty).Text) = 0 Then
        Text(ptxYotei_Qty).Text = "0"
    End If
    If Not IsNumeric(Text(ptxYotei_Qty).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxYotei_Qty).SetFocus
        Exit Function
    Else
                
        Text(ptxYotei_Qty).Text = Format(CLng(Text(ptxYotei_Qty).Text), "#0")
        If CLng(Text(ptxYotei_Qty).Text) <= 0 Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(ptxYotei_Qty).SetFocus
            Exit Function
        End If
    End If
                                            '�`�[���t
    For i = ptxDEN_DT_YY To ptxDEN_DT_DD
        If Len(Text(i).Text) = 0 Then
            Text(i).Text = "0"
        End If
        
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(i).SetFocus
            Exit Function
        Else
            RetBuf = Format(CInt(Text(i).Text), "0000")
            Text(i).Text = Right(RetBuf, Text(i).MaxLength)
        End If
    Next i
    
    If Not IsDate(Text(ptxDEN_DT_YY).Text & "/" & Text(ptxDEN_DT_MM).Text & "/" & Text(ptxDEN_DT_DD).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxDEN_DT_YY).SetFocus
        Exit Function
    End If
                                '�h�c��
    If Len(Text(ptxID_NO).Text) = 0 Then
    Else
                                                '�������ԈȊO�����͂��ꂽ��o�^�ς݃`�F�b�N
        If Not IsNumeric(Text(ptxID_NO).Text) Then
'            Beep                               '�p�����G���[�ɂ��Ȃ�
'            MsgBox "���͂������ڂ̓G���[�ł��B"
'            Text(ptxDEN_NO).SetFocus
'            Exit Function
        Else
            Text(ptxID_NO).Text = Format(CLng(Text(ptxID_NO).Text), "00000000")
        End If
        
        Call UniCode_Conv(K0_Y_NYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_Y_NYU.SYUKA_YMD, Text(ptxDEN_DT_YY).Text & Text(ptxDEN_DT_MM).Text & Text(ptxDEN_DT_DD).Text)
        Call UniCode_Conv(K0_Y_NYU.TEXT_NO, Text(ptxID_NO).Text)
        sts = BTRV(BtOpGetEqual, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
        Select Case sts
            Case BtNoErr
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i���ח\��o�^�ς݁j"
                Text(ptxID_NO).SetFocus
                Exit Function
            Case BtErrKeyNotFound
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ח\��}�X�^")
                Err_Chk = SYS_ERR
                Exit Function
        End Select
    
    End If
                                '�`�[��
    If Len(Text(ptxDEN_NO).Text) = 0 Then
    Else
        If Not IsNumeric(Text(ptxDEN_NO).Text) Then
'            Beep                               '�p�����G���[�ɂ��Ȃ�
'            MsgBox "���͂������ڂ̓G���[�ł��B"
'            Text(ptxDEN_NO).SetFocus
'            Exit Function
        Else
            Text(ptxDEN_NO).Text = Format(CLng(Text(ptxDEN_NO).Text), "000000")
        End If
    
    End If
                                        '�z�X�g�I��
'�`�F�b�N�����悤���킩���
    Err_Chk = False
    
End Function

                                            '���ח\��̒ǉ��^����
Private Function Update_Proc() As Integer
                                            
Dim sts         As Integer
Dim NAIGAI      As String * 1
Dim DEN_NO      As String * 6
Dim ID_NO       As String * 9
Dim ans         As Integer
                                            
Dim SUMI_QTY    As Long
Dim MI_QTY    As Long
                                            
    Update_Proc = True
                                        
    Call Input_Lock

    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then    '�����O�̔���
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If
                                                '�݌��ް��̃��b�N
'    sts = Zaiko_Lock_Proc(KASO_NYUKA_Soko & "01" & "01" & "01", _
'                            Last_JGYOBU, _
'                            NAIGAI, _
'                            Text(ptxHin_Gai).Text, _
'                            WS_NO)
'
'    If sts Then
'        Update_Proc = sts
'        GoTo Abort_Tran
'    End If
                                            
    SUMI_QTY = 0
    MI_QTY = 0


    If SOKO_GOODS_ON_F = GOODS_ON Then
        SUMI_QTY = CLng(Text(ptxYotei_Qty).Text)
    Else
        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
            MI_QTY = CLng(Text(ptxYotei_Qty).Text)
        Else
            SUMI_QTY = CLng(Text(ptxYotei_Qty).Text)
        End If
    End If
                                            
                                            '���ח\��ҏW
    Call UniCode_Conv(Y_NYUREC.KAN_KBN, KAN_KBN_FIN)            '�����敪
    Call UniCode_Conv(Y_NYUREC.DT_SYU, "R")                     '�f�[�^���
    Call UniCode_Conv(Y_NYUREC.JGYOBU, Last_JGYOBU)             '���ƕ�
    Call UniCode_Conv(Y_NYUREC.NAIGAI, NAIGAI)                  '�����O
    Call UniCode_Conv(Y_NYUREC.JGYOBA, "")                      '���Ə�
    Call UniCode_Conv(Y_NYUREC.DATA_KBN, "")                    '�f�[�^�敪
    Call UniCode_Conv(Y_NYUREC.TORI_KBN, "")                    '����敪
                                                                '�h�c��
    If Len(Trim(Text(ptxID_NO).Text)) <> 0 Then
        Call UniCode_Conv(Y_NYUREC.ID_NO, Text(ptxID_NO).Text)
    Else
        sts = Den_No_Set_Proc(11, Last_JGYOBU, ID_NO)
        If sts Then
            Update_Proc = sts
            GoTo Abort_Tran
        End If
    
        Call UniCode_Conv(Y_NYUREC.ID_NO, ID_NO)
        
    End If
        
    Call UniCode_Conv(Y_NYUREC.HIN_NO, Text(ptxHin_Gai).Text)   '�i�ڔԍ�
                                                                
                                                                '�`�[��
    If Len(Trim(Text(ptxDEN_NO).Text)) <> 0 Then
        Call UniCode_Conv(Y_NYUREC.DEN_NO, Text(ptxID_NO).Text)
    Else
        sts = Den_No_Set_Proc(10, Last_JGYOBU, DEN_NO)
        If sts Then
            Update_Proc = sts
            GoTo Abort_Tran
        End If
    
        Call UniCode_Conv(Y_NYUREC.DEN_NO, DEN_NO)
        
    End If
                                                                '�\�萔��
    Call UniCode_Conv(Y_NYUREC.SURYO, Format(CLng(Text(ptxYotei_Qty).Text), "0000000"))
    Call UniCode_Conv(Y_NYUREC.MUKE_CODE, "")                   '�o�ɐ�
    Call UniCode_Conv(Y_NYUREC.SYUKO_SYUSI, "")                 '�o�Ɏ��x
                                                                '�o�ɓ��t
    Call UniCode_Conv(Y_NYUREC.SYUKO_YMD, Format(Now, "YYYYMMDD"))
    Call UniCode_Conv(Y_NYUREC.TANKA, "")                       '�P��
    Call UniCode_Conv(Y_NYUREC.ODER_NO, "")                     '�I�[�_�[�ԍ�
    Call UniCode_Conv(Y_NYUREC.ITEM_NO, "")                     '�A�C�e���ԍ�
    Call UniCode_Conv(Y_NYUREC.ODER_NO_R, "")                   '�I�[�_�[����
    Call UniCode_Conv(Y_NYUREC.KOSO_KEITAI, "")                 '���`��
                                                                '�o�ד�
    Call UniCode_Conv(Y_NYUREC.SYUKA_YMD, Text(ptxDEN_DT_YY).Text & Text(ptxDEN_DT_MM).Text & Text(ptxDEN_DT_DD).Text)
                                                                '�I�ԂP
    Call UniCode_Conv(Y_NYUREC.TANABAN1, (Text(ptxST_SOKO).Text & Text(ptxST_RETU).Text & Text(ptxST_REN).Text & Text(ptxST_DAN).Text))
        
    Call UniCode_Conv(Y_NYUREC.TANABAN2, "")                    '�I�ԂQ
    Call UniCode_Conv(Y_NYUREC.TANABAN3, "")                    '�I�ԂR
    Call UniCode_Conv(Y_NYUREC.MUKE_NAME, "")                   '�o�ɐ於��
    Call UniCode_Conv(Y_NYUREC.CYU_KBN, "")                     '�����敪
    Call UniCode_Conv(Y_NYUREC.CYU_KBN_NAME, "")                '�����敪����
    Call UniCode_Conv(Y_NYUREC.ORIGIN1, "")                     '���Y���P
    Call UniCode_Conv(Y_NYUREC.ORIGIN2, "")                     '���Y���Q
    Call UniCode_Conv(Y_NYUREC.BIKOU2, "")                      '���l�Q
    Call UniCode_Conv(Y_NYUREC.HAN_KBN, "")                     '�̔��敪
    Call UniCode_Conv(Y_NYUREC.CHOKU_KBN, "")                   '�����敪
    Call UniCode_Conv(Y_NYUREC.UNIT_ID_NO, "")                  '�ƯďC��ID-NO
    Call UniCode_Conv(Y_NYUREC.ZAIKO_HIKIATE, "")               '�݌Ɉ�������
    Call UniCode_Conv(Y_NYUREC.GOKON_KANRI_NO, "")              '�����Ǘ��ԍ�
    Call UniCode_Conv(Y_NYUREC.JYUCHU_ZAN, "")                  '�󒍎c����
    Call UniCode_Conv(Y_NYUREC.KYOKYU_KBN, "")                  '�����敪
    Call UniCode_Conv(Y_NYUREC.SHOHIN_SYUSI, "")                '���i���[������x
    Call UniCode_Conv(Y_NYUREC.BIKOU1, "")                      '���l�P
    Call UniCode_Conv(Y_NYUREC.CHOHA_KBN, "")                   '���[�敪
    Call UniCode_Conv(Y_NYUREC.JYU_HIN_NO, "")                  '�󒍕i�ڔԍ�
                                                                '�i��
    Call UniCode_Conv(Y_NYUREC.HIN_NAME, Text(ptxHin_Name).Text)
    Call UniCode_Conv(Y_NYUREC.HIN_CHANGE_KBN, "")              '�i�ԕύX�敪
    Call UniCode_Conv(Y_NYUREC.MODULE_EXCHANGE, "")             '���W���[�������敪
    Call UniCode_Conv(Y_NYUREC.ZAIKO_SYUSI, "")                 '�c�݌ɂ܂Ƃߍ݌Ɏ��x�R�[�h
    Call UniCode_Conv(Y_NYUREC.NOUKI_YMD, "")                   '�w��[��
    Call UniCode_Conv(Y_NYUREC.SERVICE_KANRI_NO, "")            '�T�[�r�X��ЊǗ��ԍ�
    Call UniCode_Conv(Y_NYUREC.KI_HIN_NO, "")                   '�@��i�ڃR�[�h
    Call UniCode_Conv(Y_NYUREC.ENVIRONMENT_KBN, "")             '���K�i���i�敪
    Call UniCode_Conv(Y_NYUREC.KAN_DT, Format(Now, "YYYYMMDD")) '�������t
    Call UniCode_Conv(Y_NYUREC.BEF_NYU_QTY, "00000000")         '��s���א�
    Call UniCode_Conv(Y_NYUREC.FILLER, "")
    
    Do
        sts = BTRV(BtOpInsert, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
        
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = SYS_CANCEL
                    GoTo Abort_Tran
                End If
            Case BtErrDuplicates
                If Len(Trim(Text(ptxID_NO).Text)) = 0 Then
                                            '�������ԃf�[�^�d���͍Ĕ��s
                    sts = Den_No_Set_Proc(11, Last_JGYOBU, ID_NO)
                    If sts Then
                        Update_Proc = sts
                        GoTo Abort_Tran
                    End If
    
                    Call UniCode_Conv(Y_NYUREC.TEXT_NO, ID_NO)
                
                Else
                    Call File_Error(sts, BtOpInsert, "���ח\��f�[�^")
                    GoTo Abort_Tran
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "���ח\��f�[�^")
                GoTo Abort_Tran
        End Select
    Loop
                            
    sts = Nyuko_Update_Proc(Last_JGYOBU, _
                                    NAIGAI, _
                                    Text(ptxHin_Gai).Text, _
                                    (Text(ptxDEN_DT_YY).Text & Text(ptxDEN_DT_MM).Text & Text(ptxDEN_DT_DD).Text), _
                                    (KASO_NYUKA & "01" & "01" & "01"), _
                                    YOIN_TU_NYUKA, _
                                    SUMI_QTY, _
                                    MI_QTY, _
                                    WS_NO, _
                                    Text(ptxTanto_Code).Text, , _
                                    MEMO_TEXT)
                            
                            
    If sts Then
        Update_Proc = sts
        GoTo Abort_Tran
    End If
                                        '�g�����U�N�V�����I��
End_Tran:
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock

    
    Update_Proc = False
    
    Exit Function

Abort_Tran:
    Call Input_UnLock
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If
End Function

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i   As Integer
'----------------------------------------------------------------------------
'                   �R���{�{�b�N�X���́i�j�����c�������j����
'----------------------------------------------------------------------------
    If KeyCode = vbKeyReturn Then
        If Index = pcmbNAIGAI Then
            Call Clear_Field(0)
    
            For i = ptxHin_Gai To Text_Max
                If Text(i).Visible And Text(i).Enabled Then
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        End If
    End If
End Sub


Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            '�G���[�`�F�b�N
            sts = Err_Chk()
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                sts = Update_Proc()
                Select Case sts
                    Case False
                    Case SYS_ERR, True
                        Unload Me
                    Case SYS_CANCEL
                End Select
            
            End If
            
            Call Clear_Field(1)
            
            Text(ptxID_NO).SetFocus
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
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
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer
    
Dim sBuffer As String * 255
Dim com     As String
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
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
            F1020601.Caption = "���ח\��o�^�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)



                                '���בq�ɔԍ���荞��
    If GetIni(App.EXEName, "KASO_NYUKA", "SYS", c) Then
        Beep
        MsgBox "�i���z�j���בq�ɔԍ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    KASO_NYUKA = RTrim(c)
                                        
                                        
                                        
                                        
                                        
                                        '�u�ʏ���ׁv�̗v��
    If GetIni("YOIN", "YOIN_TU_NYUKA", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_TU_NYUKA] READ ERROR")
        MsgBox "�ʏ���חp�v���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_TU_NYUKA = Trim(c)
                                
                                
                                '����������荞��
    If GetIni(App.EXEName, "MEMO", "SYS", c) Then
        MEMO_TEXT = ""
    Else
        MEMO_TEXT = RTrim(c)
    End If

'���z�q�ɔԍ��ԍ���荞��
'    If Kaso_Soko_No_Set() Then
'        Unload Me
'    End If
                                '�V�X�e���\��ϗv����荞��
'    If SYSTEM_YOIN_Set() Then
'        Beep
'        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
'        End
'    End If
'�[���ԍ���荞��
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m�i�f�[�^�X�V�p�j
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ԃ}�X�^�n�o�d�m
    If HATUBAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���ח\��f�[�^�t�@�C���n�o�d�m
    If Y_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
'---------------------------------------------- '�q��Ͻ���菤�i���L�����l��
    Call UniCode_Conv(K0_SOKO.Soko_No, KASO_NYUKA)
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            SOKO_GOODS_ON_F = StrConv(SOKOREC.GOODS_ON_F, vbUnicode)
        
        Case BtErrKeyNotFound
            MsgBox "���חp�q�ɖ��o�^(" & KASO_NYUKA & ")"
            Unload Me
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
            Unload Me
    End Select


'---------------------------------------------- '��Ǝ���۸ނn�o�d�m
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                '��ʏ����ݒ�
    
    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
    Combo(pcmbNAIGAI).Text = NAIGAI1
        
    Call Clear_Field(0)
        
    Text(ptxTanto_Code).SetFocus

    End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�i�f�[�^�X�V�p�j�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '���ԃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, HATUBAN_POS, HATUBANREC, Len(HATUBANREC), K0_HATUBAN, Len(K0_HATUBAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ԃ}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '���ח\��f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, Y_NYU_POS, Y_NYUREC, Len(Y_NYUREC), K0_Y_NYU, Len(K0_Y_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ח\��f�[�^�t�@�C��")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, Y_NYUREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1020601 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1020601.Caption = "���ח\��o�^�i" + RTrim(JGYOBU_T(Index).NAME) + "�j"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

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
Dim sts         As Integer

Dim SUMI_QTY    As Long
Dim MI_QTY      As Long

    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        Case ptxTanto_Code          '�S���҃R�[�h
        
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTanto_Code).Text)
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    lblTanto_Name.Caption = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                    Text(ptxTanto_Code).SetFocus
                    Exit Sub
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Unload Me
            End Select
        
        Case ptxHin_Gai             '�i��
            sts = Item_Read_Proc()
            Select Case sts
                Case False
                Case True
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                    Text(Index).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
                    
                    

            If Zaiko_Syukei_Proc(SUMI_QTY, MI_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                       StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                       StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                    Unload Me
            
            End If
                        
                        
                                    

            
            Text(ptxPOS_ZAIQTY).Text = Format(SUMI_QTY + MI_QTY, "#,##0")
                    
                    
'-------------  �݌ɏW�v�ް����
                                            
                                            '�݌ɏW�v�f�[�^���z�X�g���_�݌Ɋl��
'            Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
'            Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
'            Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
'            sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
'
'            Select Case sts
'                Case BtNoErr
'                    Text(ptxHS_ZAIQTY).Text = Format(CLng(StrConv(SUMZREC.HS_ZAIQTY, vbUnicode)), "#,##0")
'                Case BtErrKeyNotFound
'                    Text(ptxHS_ZAIQTY).Text = ""
'                Case Else
'                    Call File_Error(sts, BtOpGetEqual, "�݌ɏW�v�f�[�^")
'                    Unload Me
'            End Select
                    


'-------------  �i�ڃ}�X�^�r�Q���
                    
            If IsNumeric(StrConv(ITEMREC.G_S2_ZAI_QTY, vbUnicode)) Then
                Text(ptxHS_ZAIQTY).Text = Format(CLng(StrConv(ITEMREC.G_S2_ZAI_QTY, vbUnicode)), "#,##0")
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

    F1020601.MousePointer = vbHourglass

    Call Ctrl_Lock(F1020601)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1020601)


    F1020601.MousePointer = vbDefault

End Sub

