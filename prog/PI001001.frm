VERSION 5.00
Begin VB.Form PI001001 
   Caption         =   "�\���}�X�^�[�i�Ԉꊇ�ύX"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16845
   ClipControls    =   0   'False
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
   ScaleHeight     =   10995
   ScaleWidth      =   16845
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txtTANTO_NAME 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   43
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�N���A"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   41
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ǉ��i�ԁ�"
      Height          =   3015
      Index           =   1
      Left            =   9480
      TabIndex        =   26
      Top             =   2640
      Width           =   7335
      Begin VB.TextBox txtJGYOBU 
         Height          =   375
         IMEMode         =   3  '�̌Œ�
         Left            =   3600
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox combo1 
         Appearance      =   0  '�ׯ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         ItemData        =   "PI001001.frx":0000
         Left            =   1200
         List            =   "PI001001.frx":0002
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   3
         Top             =   480
         Width           =   1260
      End
      Begin VB.ComboBox combo1 
         Appearance      =   0  '�ׯ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         ItemData        =   "PI001001.frx":0004
         Left            =   1200
         List            =   "PI001001.frx":0006
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   5
         Top             =   960
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   8
         Left            =   1200
         TabIndex        =   9
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   360
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '�E����
         Height          =   360
         IMEMode         =   3  '�̌Œ�
         Index           =   7
         Left            =   1200
         TabIndex        =   8
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   375
         Index           =   6
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   1  '�E����
         Caption         =   "���ƕ�"
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
         Left            =   2760
         TabIndex        =   45
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "��@��"
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
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "��@��"
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
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "�i�@��"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "���@��"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "���@�l�@"
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
         Left            =   240
         TabIndex        =   38
         Top             =   2400
         Width           =   735
      End
   End
   Begin VB.ListBox List1 
      Height          =   7020
      ItemData        =   "PI001001.frx":0008
      Left            =   240
      List            =   "PI001001.frx":000A
      OLEDragMode     =   1  '����
      OLEDropMode     =   1  '�蓮
      TabIndex        =   34
      Top             =   2520
      Width           =   9135
   End
   Begin VB.TextBox txtKEN_SU 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   9720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "���ύX�O�\���i�Ԏw�聄"
      Height          =   2895
      Index           =   0
      Left            =   9480
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   7335
      Begin VB.ComboBox combo1 
         Appearance      =   0  '�ׯ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         ItemData        =   "PI001001.frx":000C
         Left            =   1200
         List            =   "PI001001.frx":000E
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   11
         Top             =   960
         Width           =   1260
      End
      Begin VB.ComboBox combo1 
         Appearance      =   0  '�ׯ�
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         ItemData        =   "PI001001.frx":0010
         Left            =   1200
         List            =   "PI001001.frx":0012
         Style           =   2  '��ۯ���޳� ؽ�
         TabIndex        =   10
         Top             =   480
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   15
         Top             =   2400
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         Height          =   375
         Index           =   2
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   3
         Left            =   1200
         TabIndex        =   14
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "���@�l�@"
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
         Left            =   240
         TabIndex        =   39
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "��@��"
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
         Left            =   240
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "���@��"
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
         Left            =   240
         TabIndex        =   25
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "�i�@��"
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
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "��@��"
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
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.ComboBox combo1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      ItemData        =   "PI001001.frx":0014
      Left            =   1200
      List            =   "PI001001.frx":0016
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   1080
      Width           =   2100
   End
   Begin VB.ComboBox combo1 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1215
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   600
      Width           =   2100
   End
   Begin VB.TextBox txtTANTO_CODE 
      Appearance      =   0  '�ׯ�
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I�@�@��"
      Height          =   495
      Index           =   0
      Left            =   13440
      TabIndex        =   18
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X�@�V"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "status"
      Height          =   240
      Index           =   7
      Left            =   7800
      TabIndex        =   44
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "̧�ٖ�"
      Height          =   240
      Index           =   6
      Left            =   3840
      TabIndex        =   42
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���ǉ�/�폜�������͊Y������q�i�ԂƂȂ�܂�"
      Height          =   240
      Index           =   5
      Left            =   360
      TabIndex        =   40
      Top             =   10200
      Width           =   5160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ύX��i��"
      Height          =   240
      Index           =   4
      Left            =   5280
      TabIndex        =   37
      Top             =   2280
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ύX�O�i�ԁi���j"
      Height          =   240
      Index           =   3
      Left            =   2760
      TabIndex        =   36
      Top             =   2280
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�i�@��"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   35
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lblEXCEL_FILE 
      Height          =   255
      Left            =   4920
      TabIndex        =   33
      Top             =   1080
      Width           =   9975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���@��"
      Height          =   240
      Index           =   1
      Left            =   5520
      TabIndex        =   30
      Top             =   9840
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ǝw��"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�d����"
      Height          =   240
      Index           =   110
      Left            =   360
      TabIndex        =   20
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�S����"
      Height          =   240
      Index           =   111
      Left            =   360
      TabIndex        =   19
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "PI001001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'EXCEL ��ԍ�
Private Const exeHIN_GAI% = 1
Private Const exeBEF_HIN_GAI% = 2
Private Const exeAFT_HIN_GAI% = 3


'�e�L�X�g�p�Y��

Private Const ptxBEF_HIN_GAI% = 1           '�C���O�i��
Private Const ptxBEF_HIN_NAME% = 2          '�C���O�i��
Private Const ptxBEF_QTY% = 3               '�C���O����
Private Const ptxBEF_BIKOU% = 4             '�C���O���l

Private Const ptxAFT_HIN_GAI% = 5           '�C����i��
Private Const ptxAFT_HIN_NAME% = 6          '�C����i��
Private Const ptxAFT_QTY% = 7               '�C�������
Private Const ptxAFT_BIKOU% = 8             '�C������l




'�R���{�p�Y��
Private Const pcmbSHIMUKE% = 0              '�d������
Private Const pcmbSHORI% = 1                '����

Private Const pcmbBEF_DATA_KBN% = 2         '�C����@�敪
Private Const pcmbBEF_SYUBETSU% = 3         '�C����@���

Private Const pcmbAFT_DATA_KBN% = 4         '�C����@�敪
Private Const pcmbAFT_SYUBETSU% = 5         '�C����@���





Private KUBUN_CODE_TBL()    As String * 1   '�敪����
Private KUBUN_NAME_TBL()    As String * 4   '�敪����



Dim HIN_INV_F       As Integer
Dim BIKOU_F         As Integer
Dim BIKOU_SET_F     As Integer



Dim PI00100_LOG_F   As String

'Private Const Last_Update_day$ = "�\���}�X�^�[�i�Ԉꊇ�ύX (PI00100 2019.04.24 15:30)"
'Private Const Last_Update_day$ = "�\���}�X�^�[�i�Ԉꊇ�ύX (PI00100 2019.04.26 13:30)"
'Private Const Last_Update_day$ = "�\���}�X�^�[�i�Ԉꊇ�ύX (PI00100 2019.05.24 16:56)"
'Private Const Last_Update_day$ = "�\���}�X�^�[�i�Ԉꊇ�ύX (PI00100 2019.06.04 19:30)"
'Private Const Last_Update_day$ = "�\���}�X�^�[�i�Ԉꊇ�ύX (PI00100 2019.06.05 11:42)"
Private Const Last_Update_day$ = "�\���}�X�^�[�i�Ԉꊇ�ύX (PI00100 2019.06.05 15:40)"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI001001.MousePointer = vbHourglass

    Call Ctrl_Lock(PI001001)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI001001)


    PI001001.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim sts     As Integer
    
    Error_Check_Proc = True









    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxAFT_HIN_GAI).Text)


    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
                
            Text1(ptxAFT_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            txtJGYOBU = Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1)
            
                
                
        Case BtErrKeyNotFound
            
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxAFT_HIN_GAI).Text)
        
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                        
                    Text1(ptxAFT_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    txtJGYOBU = SHIZAI
                    
                        
                        
                Case BtErrKeyNotFound
                    
                    
                    
                    
                    Text1(ptxAFT_HIN_NAME).Text = "���o�^�i�Ԃł��B"
                        
                        
                    If HIN_INV_F = 1 Then
                        MsgBox "�i�ڃ}�X�^���o�^�ł�"
                        Text1(ptxAFT_HIN_GAI).SetFocus
                        
                        Exit Function
                    End If
                        
                        
                Case Else
                            
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    Exit Function
                    
        
            End Select
            
            
            
            
                
        Case Else
                    
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
            Exit Function
            

    End Select







    If Right(combo1(pcmbSHORI).Text, 1) = "1" Then              '�ǉ��̎�
        If Not IsNumeric(Text1(ptxAFT_QTY).Text) Then
            MsgBox "�����͐��l����͂��ĉ������B"
            Text1(ptxAFT_QTY).SetFocus
            
            Exit Function
        End If
        
        If BIKOU_F = 1 Then
        
            If Text1(ptxAFT_BIKOU).Text = "" Then
                MsgBox "���l�͕K�{���͂ł��B"
                 Text1(ptxAFT_BIKOU).SetFocus
                    
                Exit Function
            End If
        End If
    
    
    End If




    Error_Check_Proc = False


End Function


Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �\���}�X�^�����i���w���ް��o��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim com         As Integer

Dim SEQNO       As Integer

Dim i           As Integer
Dim j           As Integer

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�\���}�X�^�@�X�V�����J�n", Me.hwnd, 0)



    Update_Proc = True

                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If



'>>>>>>�\���}�X�^(���ި)�o��


    
    
    
    
    Select Case Right(combo1(pcmbSHORI).Text, 1)
        
        Case 1      '�ǉ�
        
            
            
            
            For i = 0 To List1.ListCount - 1
                SEQNO = 0
            
                If Right(List1.List(i), 3) = "Err" Then
                    If Trim(PI00100_LOG_F) <> "" Then
                        Call LOG_OUT(PI00100_LOG_F, List1.List(i))
                    
                    End If
                Else
                    
                    
                    com = BtOpGetGreaterEqual
                
                    Do
                    
                        DoEvents
                    
                        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(combo1(pcmbSHIMUKE), 4), 1, 2))
                        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1))
                        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
                        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Left(List1.List(i), 20))
                            
                        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, Right(combo1(pcmbAFT_DATA_KBN).Text, 1))
                        
                        SEQNO = SEQNO + 10
                        
                        Call UniCode_Conv(K0_P_COMPO.SEQNO, Format(SEQNO, "000"))
                            
                        sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Exit Do
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                GoTo Abort_Tran
                        End Select
                    
                    Loop
            
                            
                                                                                                '�d�����溰��
                    Call UniCode_Conv(P_COMPO_K_REC.SHIMUKE_CODE, Mid(Right(combo1(pcmbSHIMUKE).Text, 4), 1, 2))
                                                                                                '���ƕ�
                    Call UniCode_Conv(P_COMPO_K_REC.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE).Text, 4), 3, 1))
                                                                                                '�����O
                    Call UniCode_Conv(P_COMPO_K_REC.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE).Text, 4), 4, 1))
                                                                                                '�i�ԁi�e�j
                    Call UniCode_Conv(P_COMPO_K_REC.HIN_GAI, Left(List1.List(i), 20))
                                                                                                '�f�[�^�敪
                    Call UniCode_Conv(P_COMPO_K_REC.DATA_KBN, Right(combo1(pcmbAFT_DATA_KBN).Text, 1))
                                                                                                
                                                                                                
                                                                                                
                                                                                                
                                                                                                
                    Call UniCode_Conv(P_COMPO_K_REC.SEQNO, Format(SEQNO, "000"))                                '�ǔ�
                    Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, Right(combo1(pcmbAFT_SYUBETSU).Text, 2))       '���
                    Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, txtJGYOBU.Text)                                  '���ƕ�
                    Call UniCode_Conv(P_COMPO_K_REC.KO_NAIGAI, NAIGAI_NAI)                                      '�����O
                    Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(ptxAFT_HIN_GAI).Text)                     '�i��
                    Call UniCode_Conv(P_COMPO_K_REC.KO_QTY, Format(Val(Text1(ptxAFT_QTY).Text), "000.00"))      '����                                                                                    '����
                    Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, Text1(ptxAFT_BIKOU).Text)                         '���l
    
                    Call UniCode_Conv(P_COMPO_K_REC.CLASS_CODE, "")
                    Call UniCode_Conv(P_COMPO_K_REC.FILLER, "")
                    Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, txtTANTO_CODE.Text)                              '�X�V�S���Һ���
                                                                                                                '�X�V����
                    Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
    
            
                    sts = BTRV(BtOpInsert, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                    Select Case sts
                        Case BtNoErr
                        Case Else
                            Call File_Error(sts, BtOpInsert, "�\���}�X�^")
                            GoTo Abort_Tran
                    End Select
        
        
                    If Trim(PI00100_LOG_F) <> "" Then
                        Call LOG_OUT(PI00100_LOG_F, List1.List(i) & " INS")
                    
                    End If
        
        
        
                End If
            Next i
        
        
        Case 2      '�ύX
        
        
            For i = 0 To List1.ListCount - 1
        
                If Right(List1.List(i), 3) = "Err" Then
        
                    If Trim(PI00100_LOG_F) <> "" Then
                        Call LOG_OUT(PI00100_LOG_F, List1.List(i) & "INS")
                    
                    End If
        
                Else
        
                    com = BtOpGetGreaterEqual
                
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(combo1(pcmbSHIMUKE), 4), 1, 2))
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1))
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Left(List1.List(i), 20))
                        
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, Right(combo1(pcmbAFT_DATA_KBN).Text, 1))
                    '2019.05.24 ����
                    '               �݌��l����̎w���Łu�敪�͖��֌W�v�Ƃ����B
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
                    
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                
                
                
                    Do
                    
                        DoEvents
                    
                            
                        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            
                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(combo1(pcmbSHIMUKE), 4), 1, 2) Then
                                    Exit Do
                                End If
                                
                                If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Left(List1.List(i), 20)) Then
                                    Exit Do
                                End If
                                
                    '2019.05.24 ����
                    '               �݌��l����̎w���Łu�敪�͖��֌W�v�Ƃ����B
'                                If Trim(StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)) <> Right(combo1(pcmbAFT_DATA_KBN).Text, 1) Then
'                                    Exit Do
'                                End If
                                
                            
                            
                            
                                If Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) = Trim(Mid(List1.List(i), 22, 20)) Then
                                
                                    Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, txtJGYOBU)
                                    Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(ptxAFT_HIN_GAI).Text)
                                    Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, Text1(ptxAFT_BIKOU).Text)
                                
                                
                                    Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, txtTANTO_CODE.Text)                              '�X�V�S���Һ���
                                                                                                                                '�X�V����
                                    Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                    
                            
                                    sts = BTRV(BtOpUpdate, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        Case Else
                                            Call File_Error(sts, BtOpUpdate, "�\���}�X�^")
                                            GoTo Abort_Tran
                                    End Select
                                
                                
                                End If
                            
                            
                            
                                If Trim(PI00100_LOG_F) <> "" Then
                                    Call LOG_OUT(PI00100_LOG_F, List1.List(i) & " UPD")
                                
                                End If
                            
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                GoTo Abort_Tran
                        End Select
                    
            
            
                        com = BtOpGetNext
                
                
                    Loop
                End If
        
            Next i
        
        
        
        
        Case 3      '�폜
    
    
    
            For i = 0 To List1.ListCount - 1
        
                If Right(List1.List(i), 3) = "Err" Then
    
                    If Trim(PI00100_LOG_F) <> "" Then
                        Call LOG_OUT(PI00100_LOG_F, List1.List(i) & "INS")
                    
                    End If
                Else
    
    
                    com = BtOpGetGreaterEqual
                
                
                
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(combo1(pcmbSHIMUKE), 4), 1, 2))
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1))
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Left(List1.List(i), 20))
                        
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, Right(combo1(pcmbAFT_DATA_KBN).Text, 1))
                    '2019.05.24 ����
                    '               �݌��l����̎w���Łu�敪�͖��֌W�v�Ƃ����B
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "")
                    
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")
                
                
                
                    Do
                    
                        DoEvents
                    
                            
                        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            
                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(combo1(pcmbSHIMUKE), 4), 1, 2) Then
                                    Exit Do
                                End If
                                
                                If Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Left(List1.List(i), 20)) Then
                                    Exit Do
                                End If
                                
                    '2019.05.24 ����
                    '               �݌��l����̎w���Łu�敪�͖��֌W�v�Ƃ����B
'                                If Trim(StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode)) <> Right(combo1(pcmbAFT_DATA_KBN).Text, 1) Then
'                                    Exit Do
'                                End If
                                
                            
                            
                            
                                If Trim(StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)) = Trim(Mid(List1.List(i), 22, 20)) Then
                                    
                                    '2019.05.24 ����
                                    '                   ���L�͍폜����̂ŁA�ҏW�s�v�I
'                                    Call UniCode_Conv(P_COMPO_K_REC.KO_JGYOBU, txtJGYOBU)
'                                    Call UniCode_Conv(P_COMPO_K_REC.KO_HIN_GAI, Text1(ptxAFT_HIN_GAI).Text)
'                                    Call UniCode_Conv(P_COMPO_K_REC.KO_BIKOU, Text1(ptxAFT_BIKOU).Text)
'
'
'                                    Call UniCode_Conv(P_COMPO_K_REC.UPD_TANTO, txtTANTO_CODE.Text)                              '�X�V�S���Һ���
'                                                                                                                                '�X�V����
'                                    Call UniCode_Conv(P_COMPO_K_REC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
                    
                            
                                    sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        Case Else
                                            Call File_Error(sts, BtOpDelete, "�\���}�X�^")
                                            GoTo Abort_Tran
                                    End Select
                                
                                
                                    If Trim(PI00100_LOG_F) <> "" Then
                                        Call LOG_OUT(PI00100_LOG_F, List1.List(i) & " DEL")
                                    
                                    End If
                                
                                
                                
                                End If
                            
                            
                            
                            
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                GoTo Abort_Tran
                        End Select
                    
            
            
                        com = BtOpGetNext
                    Loop
            
                End If
            Next i
    
    
    End Select









































End_Tran:
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If


    MsgBox "�f�[�^�X�V���I�����܂���"


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�\���}�X�^�@�X�V��������I��", Me.hwnd, 0)


    Update_Proc = False

    Exit Function

Abort_Tran:

    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    MsgBox "�f�[�^�X�V���I�����܂���"


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�\���}�X�^�@�ُ폈������I��", Me.hwnd, 0)

End Function








Private Sub combo1_Click(Index As Integer)
Dim W_STR       As String


    Select Case Index
        Case pcmbSHIMUKE
'�R���{�p�Y��
'Private Const pcmbSHIMUKE% = 0              '�d������
'Private Const pcmbSHORI% = 1                '����
'
'Private Const pcmbBEF_DATA_KBN% = 2         '�C����@�敪
'Private Const pcmbBEF_SYUBETSU% = 3         '�C����@���
'
'Private Const pcmbAFT_DATA_KBN% = 4         '�C����@�敪
'Private Const pcmbAFT_SYUBETSU% = 5         '�C����@���
'            combo1(pcmbSHORI%).SetFocus                         '2019.06.03 ����
            Call Tab_Ctrl(0)
            
        Case pcmbSHORI
            Select Case Right(combo1(Index).Text, 1)
                Case 1

                    Frame1(1).Enabled = True


                    Label2(3).Visible = True
                    Label2(5).Visible = True
                    Label2(7).Visible = True
                    Label2(8).Visible = True
                    Text1(ptxAFT_BIKOU).Visible = True


                    Text1(ptxAFT_QTY).Visible = True


                    combo1(pcmbAFT_DATA_KBN).Visible = True

                    combo1(pcmbAFT_SYUBETSU).Visible = True
                    
                    DoEvents
                    
                    combo1(pcmbAFT_DATA_KBN).SetFocus           '2019.06.03 ����

                Case 2
                    Frame1(1).Enabled = True


                    Label2(3).Visible = False
                    Label2(5).Visible = False
                    Label2(7).Visible = False
                    Label2(8).Visible = True
                    Text1(ptxAFT_BIKOU).Visible = True


                    Text1(ptxAFT_QTY).Visible = False

                    combo1(pcmbAFT_DATA_KBN).Visible = False

                    combo1(pcmbAFT_SYUBETSU).Visible = False
                    DoEvents
                    Text1(ptxAFT_BIKOU).SetFocus

                Case 3
                    Frame1(1).Enabled = False

                    Label2(3).Visible = False
                    Label2(5).Visible = False
                    Label2(7).Visible = False
                    Label2(8).Visible = False
                    Text1(ptxAFT_BIKOU).Visible = False


                    Text1(ptxAFT_QTY).Visible = False


                    combo1(pcmbAFT_DATA_KBN).Visible = False

                    combo1(pcmbAFT_SYUBETSU).Visible = False
                    DoEvents
            End Select


        Case pcmbAFT_DATA_KBN
            W_STR = Right(combo1(Index).Text, 1)
'[PI00100]
'KUBUN=������,1,�O������,2,�\�����i,3                 '�݌�������̗v�]
            Select Case W_STR
                Case "1", "2"
                    Text1(ptxAFT_BIKOU) = ""            '  = 8             '�C������l)
                    '2019.06.05 ���͕s�I�ɂ����i���X�̗v�]�̂悤�ł��j
                    Text1(ptxAFT_BIKOU).Locked = True
                    
                Case Else
                    '2019.06.05 ���R�A���̋敪�̎��́ALock�����I
                    Text1(ptxAFT_BIKOU).Locked = False
            End Select
            
'            combo1(pcmbAFT_SYUBETSU).SetFocus           '2019.06.03 �ǉ��@����
            
            txtJGYOBU.SetFocus                          '2019.06.03 �ǉ��@����
        

        Case pcmbAFT_SYUBETSU
            Text1(ptxAFT_QTY%).SetFocus                 '2019.06.03 �ǉ��@����


    End Select

End Sub


Private Sub combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim W_STR       As String



    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If


    
    Select Case Index
        Case pcmbSHIMUKE
'�R���{�p�Y��
'Private Const pcmbSHIMUKE% = 0              '�d������
'Private Const pcmbSHORI% = 1                '����
'
'Private Const pcmbBEF_DATA_KBN% = 2         '�C����@�敪
'Private Const pcmbBEF_SYUBETSU% = 3         '�C����@���
'
'Private Const pcmbAFT_DATA_KBN% = 4         '�C����@�敪
'Private Const pcmbAFT_SYUBETSU% = 5         '�C����@���
            
'            '2019.06.04 �ǉ��i�֑��H�j
'            If Trim(txtJGYOBU) = "" Then
'                txtJGYOBU = Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1)
'            End If
            
'            combo1(pcmbSHORI%).SetFocus                         '2019.06.03 ����
            Call Tab_Ctrl(Shift)
            
        Case pcmbSHORI
            Select Case Right(combo1(Index).Text, 1)
                Case 1
                                
                    Frame1(1).Enabled = True
                
                
                    Label2(3).Visible = True
                    Label2(5).Visible = True
                    Label2(7).Visible = True
                    Label2(8).Visible = True
                    Text1(ptxAFT_BIKOU).Visible = True
                    
                    
                    Text1(ptxAFT_QTY).Visible = True
                    
                    
                    combo1(pcmbAFT_DATA_KBN).Visible = True
                
                    combo1(pcmbAFT_SYUBETSU).Visible = True
                    
                    combo1(pcmbAFT_DATA_KBN).SetFocus           '2019.06.03 ����
                    
                Case 2
                    Frame1(1).Enabled = True
                
                
                    Label2(3).Visible = False
                    Label2(5).Visible = False
                    Label2(7).Visible = False
                    Label2(8).Visible = True
                    Text1(ptxAFT_BIKOU).Visible = True
                    
                    
                    Text1(ptxAFT_QTY).Visible = False
                    
                    combo1(pcmbAFT_DATA_KBN).Visible = False
                
                    combo1(pcmbAFT_SYUBETSU).Visible = False
                
                    Text1(ptxAFT_BIKOU).SetFocus
                    
                Case 3
                    Frame1(1).Enabled = False

                    Label2(3).Visible = False
                    Label2(5).Visible = False
                    Label2(7).Visible = False
                    Label2(8).Visible = False
                    Text1(ptxAFT_BIKOU).Visible = False
                    
                    
                    Text1(ptxAFT_QTY).Visible = False
                    
                    
                    combo1(pcmbAFT_DATA_KBN).Visible = False
                
                    combo1(pcmbAFT_SYUBETSU).Visible = False
            
            End Select
        
        
        Case pcmbAFT_DATA_KBN
            W_STR = Right(combo1(Index).Text, 1)
'[PI00100]
'KUBUN=������,1,�O������,2,�\�����i,3                 '�݌�������̗v�]
            Select Case W_STR
                Case "1", "2"
                    Text1(ptxAFT_BIKOU) = ""            '  = 8             '�C������l)
                    '2019.06.05 ���͕s�I�ɂ����i���X�̗v�]�̂悤�ł��j
                    Text1(ptxAFT_BIKOU).Locked = True
                    
                Case Else
                    '2019.06.05 ���R�A���̋敪�̎��́ALock�����I
                    Text1(ptxAFT_BIKOU).Locked = False
            End Select
            
'            combo1(pcmbAFT_SYUBETSU).SetFocus           '2019.06.03 �ǉ��@����
            
            txtJGYOBU.SetFocus                          '2019.06.03 �ǉ��@����
        
        Case pcmbAFT_SYUBETSU
            Text1(ptxAFT_QTY%).SetFocus                 '2019.06.03 �ǉ��@����
            
        
    End Select
    
End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans     As Integer


    Select Case Index
        Case 0
            Unload Me
    
    
        Case 1
                
            If List1.ListCount <= 0 Then
                Exit Sub
            End If
        
            If Error_Check_Proc() Then
                Exit Sub
            End If
        
        
            If Update_Proc() Then
                Exit Sub
            End If
        
        
        Case 2
        
            Call Init_Proc
        
    
    End Select



End Sub


Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer

Dim wkVAL       As Variant


Show    '2015.03.26


    PI001001.Caption = Last_Update_day      '2016.02.10

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�\���}�X�^�[�i�Ԉꊇ�ύX�@�u�N���������v", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
'    Me.Enabled = False
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.10.09
                                
                                
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LOG_F = RTrim(c)











'                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If


    '����Ͻ���`
    Call P_CODE_TBL_Proc






    '�d������̃Z�b�g
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If

    '���
'    If Code_Set_Proc(pcmbBEF_SYUBETSU, P_KBN06_CD, 1) Then
    '2019.06.03 �w���ɂ��A�󔒁��s�Ƃ���
    If Code_Set_Proc(pcmbBEF_SYUBETSU, P_KBN06_CD, 0) Then
    
        Unload Me
    End If
    
'    If Code_Set_Proc(pcmbAFT_SYUBETSU, P_KBN06_CD, 1) Then
    '2019.06.03 �w���ɂ��A�󔒁��s�Ƃ���
    If Code_Set_Proc(pcmbAFT_SYUBETSU, P_KBN06_CD, 0) Then
        Unload Me
    End If

    combo1(pcmbBEF_SYUBETSU).ListIndex = 0
    combo1(pcmbAFT_SYUBETSU).ListIndex = 0




    '�敪
    If GetIni(App.EXEName, "KUBUN", App.EXEName, c) Then
        c = "*,*"
    End If
    wkVAL = Split(Trim(c), ",", -1)
    
    i = 0
    j = 0
    Do
            
        If i > UBound(wkVAL) Then
            Exit Do
        End If
        
        ReDim Preserve KUBUN_CODE_TBL(0 To j)
        ReDim Preserve KUBUN_NAME_TBL(0 To j)
            
        KUBUN_NAME_TBL(j) = wkVAL(i)
        KUBUN_CODE_TBL(j) = wkVAL(i + 1)
            
            
        i = i + 2
        j = j + 1
    
    Loop
    
    
    
    If GetIni(App.EXEName, "HIN_INV", App.EXEName, c) Then
        HIN_INV_F = 0
    Else
        If Trim(c) = "1" Then
            HIN_INV_F = 1
        Else
            HIN_INV_F = 0
        End If
    End If
    
    
    If GetIni(App.EXEName, "BIKOU_F", App.EXEName, c) Then
        BIKOU_F = 0
    Else
        If Trim(c) = "1" Then
            BIKOU_F = 1
        Else
            BIKOU_F = 0
        End If
    End If
    
    
    If GetIni(App.EXEName, "BIKOU_SET_F", App.EXEName, c) Then
        BIKOU_SET_F = 0
    Else
        If Trim(c) = "1" Then
            BIKOU_SET_F = 1
        Else
            BIKOU_SET_F = 0
        End If
    End If
    
    
    If GetIni(App.EXEName, "PI00100_LOG", App.EXEName, c) Then
        PI00100_LOG_F = ""
    Else
        PI00100_LOG_F = Trim(c)
    End If
    
    
    
    combo1(pcmbBEF_DATA_KBN).Clear
    combo1(pcmbAFT_DATA_KBN).Clear
    
    
    For i = 0 To UBound(KUBUN_CODE_TBL)
    
        If KUBUN_CODE_TBL(i) = "*" Then
        Else
            combo1(pcmbBEF_DATA_KBN).AddItem KUBUN_NAME_TBL(i) & "     " & KUBUN_CODE_TBL(i)
            combo1(pcmbAFT_DATA_KBN).AddItem KUBUN_NAME_TBL(i) & "     " & KUBUN_CODE_TBL(i)
        End If
    


    Next i
    combo1(pcmbBEF_DATA_KBN).ListIndex = 0
    combo1(pcmbAFT_DATA_KBN).ListIndex = 0
    
    
    '��ʏ����ݒ�
    If Init_Proc() Then
        Unload Me
    End If

'2009.03.25

    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�\���}�X�^�[�i�Ԉꊇ�ύX�@�u���������v", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    Me.Enabled = True
    
    '���
    combo1(pcmbSHORI).AddItem "�ǉ�" & Space(20) & "1"
    combo1(pcmbSHORI).AddItem "�ύX" & Space(20) & "2"
    combo1(pcmbSHORI).AddItem "�폜" & Space(20) & "3"
    combo1(pcmbSHORI).ListIndex = 0
    
    
    combo1(pcmbSHIMUKE).ListIndex = 0
    combo1(pcmbSHIMUKE).SetFocus
    DoEvents
    txtTANTO_CODE.SetFocus
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer



    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "", 0)
    End If
    Set PI001001 = Nothing

    End
End Sub


Private Function Init_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���͉�ʂ̏����ݒ�
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer


    Init_Proc = True




    List1.Clear


    combo1(pcmbAFT_DATA_KBN).ListIndex = 0
    combo1(pcmbAFT_SYUBETSU).ListIndex = 0
    

    For i = pcmbAFT_DATA_KBN To ptxAFT_BIKOU
        
        Text1(i).Text = ""
    
    Next i
    
    txtJGYOBU = ""
    
    Init_Proc = False

End Function

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

Dim WK As String

Dim i           As Integer

Start_Proc0:        '2015.03.26ok

    Code_Set_Proc = True

    combo1(Index).Clear

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
        combo1(Index).AddItem Space(Key_Len)
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
        
        combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption


        com = BtOpGetNext

    Loop

    Code_Set_Proc = False




End Function



Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �uEXCEL�v�Ǎ��ݏ���
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long



Dim wkHIN_GAI       As String * 20
Dim wkBEF_HIN_GAI   As String * 20
Dim wkAFT_HIN_GAI   As String * 20

Dim i               As Long
Dim j               As Long

Dim TEXT_BEF        As String

Dim SvBEF_HIN_GAI   As String * 20
Dim SvAFT_HIN_GAI   As String * 20


Dim Err_Mark        As String * 3


    List_Disp_Proc = True

'    Call Input_Lock


    PI001001.Enabled = False

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�捞�݃f�[�^�@�\�������J�n", Me.hwnd, 0)





    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (lblEXCEL_FILE.Caption), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0
    
    
    Row = 0
    
    
    Call Init_Proc
    
    
    
    Set xlSheet = xlApp.Worksheets(1)
    xlSheet.Activate
    
    
    For i = 1 To 1048576
            
            
        If Trim(xlSheet.Application.Cells(i, exeHIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, exeBEF_HIN_GAI)) = "" And _
            Trim(xlSheet.Application.Cells(i, exeAFT_HIN_GAI)) = "" Then
            Exit For
        End If
        
        
        If i > 1 Then
        
        
        
        
        
        '�i��
            wkHIN_GAI = Trim(xlSheet.Application.Cells(i, exeHIN_GAI))
            '�C���O�i��
            wkBEF_HIN_GAI = Trim(xlSheet.Application.Cells(i, exeBEF_HIN_GAI))
            '�C����i��
            wkAFT_HIN_GAI = Trim(xlSheet.Application.Cells(i, exeAFT_HIN_GAI))
    
    
            Err_Mark = ""
    
    
    
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, wkHIN_GAI)


            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
            
                    Err_Mark = "Err"
            
                
                Case Else
                    
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    Exit Function
            

            End Select

    
            Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(combo1(pcmbSHIMUKE), 4), 1, 2))
            Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_P_COMPO.HIN_GAI, wkHIN_GAI)
        
            Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
            Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
        
            sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            Select Case sts
                Case BtNoErr
                    
                
                
                
                Case BtErrKeyNotFound
                
                    Err_Mark = "Err"
                        
        
                Case Else
                        
                    Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                    Exit For
            
            
            End Select
    
    
    
    
    
            List1.AddItem wkHIN_GAI & " " & wkBEF_HIN_GAI & " " & wkAFT_HIN_GAI & " " & Err_Mark
        
        
        
        
        
            If Right(combo1(pcmbSHORI).Text, 1) = "2" Then
                If Trim(xlSheet.Application.Cells(i, exeAFT_HIN_GAI)) = "" Then
            
                    MsgBox "�d�w�b�d�k�f�[�^�̓��e���s���ł��B���e���m�F���ĉ������B"
                    Exit For
                    
                End If
            
            
            Else
                If Trim(xlSheet.Application.Cells(i, exeAFT_HIN_GAI)) <> "" Then
            
                    MsgBox "�d�w�b�d�k�f�[�^�̓��e���s���ł��B���e���m�F���ĉ������B"
                    Exit For
                    
                End If
            End If
        
        
        
            If i = 2 Then
                SvBEF_HIN_GAI = wkBEF_HIN_GAI
                SvAFT_HIN_GAI = wkAFT_HIN_GAI
            End If
        
        
            If SvBEF_HIN_GAI <> wkBEF_HIN_GAI Then
                MsgBox "�قȂ�q�i�ԁi�ύX�O�j�͎w��o���܂���B�f�[�^�𕪊����ĉ������B"
                Exit For
            End If
        
            If SvAFT_HIN_GAI <> wkAFT_HIN_GAI Then
                MsgBox "�قȂ�q�i�ԁi�ύX��j�͎w��o���܂���B�f�[�^�𕪊����ĉ������B"
                Exit For
            End If
        
        
        
        
        
        
        
        
        
        
        
        
            SvBEF_HIN_GAI = wkBEF_HIN_GAI
            SvAFT_HIN_GAI = wkAFT_HIN_GAI
        
            Row = Row + 1
    
        End If
    
    
    
    
    Next i


    



    txtKEN_SU.Text = Format(Row, "#0") & "��"

'    Call SendMessage(Combo2.hwnd, CB_SHOWDROPDOWN, True, 0)


    PI001001.Enabled = True




    If Right(combo1(pcmbSHORI).Text, 1) = "2" Then
        If Item_Disp_Proc(wkAFT_HIN_GAI) Then
        End If
    Else
        If Item_Disp_Proc(wkBEF_HIN_GAI) Then
        End If
    End If



    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit 'EXCEL�����
    Set xlApp = Nothing




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�捞�݃f�[�^�@�\�������I��", Me.hwnd, 0)



'    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 �t�@�C�����܂��͔ԍ����s���ł��B
        '53 �t�@�C����������܂���B
        '54 �t�@�C�� ���[�h���s���ł��B
        '55 �t�@�C���͊��ɊJ����Ă��܂��B
        '57 �f�o�C�X I/O �G���[�ł��B
        '59 ���R�[�h������v���܂���B
        '61 �f�B�X�N�̋󂫗e�ʂ��s�����Ă��܂��B
        '62 �t�@�C���ɂ���ȏ�f�[�^������܂���B
        '63 ���R�[�h�ԍ����s���ł��B
        '68 �f�o�C�X����������Ă��܂���B
        '70 �������݂ł��܂���B
        '71 �f�B�X�N����������Ă��܂���B
        '75 �p�X���������ł��B
        '76 �p�X��������܂���B
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "�w��̃t�@�C����������܂���B" & Chr(13) & Chr(10) & "�������t�@�C��������͂��Ă��������B"
            
            
            
            List_Disp_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "�Ǎ��ݑΏۂ�EXCEL�f�[�^�Ɉُ킪�L��܂��B���e���m�F��A�Ď��s���Ă��������B"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCEL�����
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_Proc = False      '



        Case Else
    End Select
 '   Call Input_UnLock

End Function


Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sts     As Integer
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    
    Select Case sts
        Case BtNoErr
            txtTANTO_NAME.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            txtTANTO_NAME.Text = ""

            MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
            txtTANTO_CODE.SetFocus
            Exit Sub
        
        Case Else
            
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Unload Me

    End Select
    
    
    
    
    lblEXCEL_FILE.Caption = Trim(Data.Files(1))


    If List_Disp_Proc() Then
        Unload Me
    End If

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index).Text = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case ptxAFT_QTY '= 7               '�C�������
            
            If Right(combo1(pcmbSHORI).Text, 1) = "1" Then              '�ǉ��̎�
                If Trim(Text1(Index)) = "" Then
                    MsgBox "�������w�肵�ĉ������B", vbExclamation
                    Text1(Index).SetFocus
                    Exit Sub
                Else
                    If Not IsNumeric(Trim(Text1(Index))) Then
                        MsgBox "�����͐��l����͂��ĉ������B", vbExclamation
                        Text1(Index).SelStart = 0
                        Text1(Index).SelLength = Len(Text1(Index))
                        Text1(Index).SetFocus
                        Exit Sub
                    End If
                    
                    If CLng(Trim(Text1(Index))) = 0 Then
                        MsgBox "�[���G���[�I", vbExclamation
                        Text1(Index).SelStart = 0
                        Text1(Index).SelLength = Len(Text1(Index))
                        Text1(Index).SetFocus
                        Exit Sub
                    End If
                    
                End If
            End If
            
        Case Else
    
    
    End Select
    
    
    Call Tab_Ctrl(Shift)        '�ړ�
    

End Sub

Private Sub txtJGYOBU_GotFocus()
    txtJGYOBU.Text = Trim(txtJGYOBU.Text)
    txtJGYOBU.SelStart = 0
    txtJGYOBU.SelLength = 1
    
End Sub

Private Sub txtJGYOBU_KeyDown(KeyCode As Integer, Shift As Integer)
Dim WK      As String

    If KeyCode <> vbKeyReturn Then Exit Sub
    
        
    
'    If Trim(txtJGYOBU) = "" Then
'        MsgBox "�����ꂩ�̎��ƕ����w�肵�ĉ������B", vbExclamation
'        txtJGYOBU.SetFocus
'        Exit Sub
'    End If
    
    WK = UCase(txtJGYOBU)
    txtJGYOBU = WK
    DoEvents
    
    Select Case txtJGYOBU
        Case "B", "S"
        
        Case Else
            
    End Select
    
    combo1(pcmbAFT_SYUBETSU).SetFocus           '2019.06.03 �ǉ��@����
    
End Sub

Private Sub txtTANTO_CODE_GotFocus()
    If txtTANTO_CODE.TabStop = True Then
        txtTANTO_CODE.Text = Trim(txtTANTO_CODE.Text)
        txtTANTO_CODE.SelStart = 0
        txtTANTO_CODE.SelLength = Len(txtTANTO_CODE.Text)
    End If

End Sub

Private Sub txtTANTO_CODE_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
Dim sts As Integer
    
    
    
    If KeyCode <> vbKeyReturn Then Exit Sub


                
                
                
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    
    Select Case sts
        Case BtNoErr
            txtTANTO_NAME.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            txtTANTO_NAME.Text = ""

            MsgBox "���͂������ڂ̓G���[�ł��B(�S����)"
            txtTANTO_CODE.SelStart = 0
            txtTANTO_CODE.SelLength = Len(txtTANTO_CODE.Text)
            txtTANTO_CODE.SetFocus
            Exit Sub
        Case Else
            
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Unload Me

    End Select




    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Function Item_Disp_Proc(HIN_GAI As String) As Integer
'----------------------------------------------------------------------------
'                   ���ו\��
'----------------------------------------------------------------------------
Dim sts     As Integer

'    If Right(combo1(pcmbSHORI).Text, 1) <> "1" Then
'        Exit Function
'    End If
    
        
    Item_Disp_Proc = True
    
    Text1(ptxAFT_HIN_GAI).Text = HIN_GAI


    Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1))
    Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)


    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
                
            Text1(ptxAFT_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
            txtJGYOBU = Mid(Right(combo1(pcmbSHIMUKE), 4), 3, 1)
            
                
                
        Case BtErrKeyNotFound
            
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, HIN_GAI)
        
        
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                        
                    Text1(ptxAFT_HIN_NAME).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    txtJGYOBU = SHIZAI
                    
                        
                        
                Case BtErrKeyNotFound
                    
                     Text1(ptxAFT_HIN_NAME).Text = "���o�^�i�Ԃł��B"
                        
                Case Else
                            
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    Exit Function
                    
        
            End Select
            
            
            
            
                
        Case Else
                    
            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
            Exit Function
            

    End Select

    If BIKOU_SET_F = 1 Then
        Text1(ptxAFT_BIKOU).Text = Text1(ptxAFT_HIN_NAME).Text
    End If



    Item_Disp_Proc = False

End Function
