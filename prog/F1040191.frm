VERSION 5.00
Begin VB.Form F1040191 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�݌ɒP�������e�i���X"
   ClientHeight    =   6915
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   12555
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   12555
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   46
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   13
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   12
      Left            =   1365
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3480
      Width           =   660
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   1365
      MaxLength       =   8
      TabIndex        =   40
      Top             =   2880
      Width           =   1080
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      Height          =   360
      Index           =   0
      Left            =   2100
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   38
      Top             =   2280
      Width           =   2850
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   1365
      MaxLength       =   5
      TabIndex        =   37
      Top             =   2280
      Width           =   660
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   2835
      MaxLength       =   2
      TabIndex        =   35
      Top             =   1680
      Width           =   345
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   2205
      MaxLength       =   2
      TabIndex        =   33
      Top             =   1680
      Width           =   345
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   1365
      MaxLength       =   4
      TabIndex        =   31
      Top             =   1680
      Width           =   555
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   3255
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1080
      Width           =   345
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   2625
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1080
      Width           =   345
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   1995
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1080
      Width           =   345
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   1365
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1080
      Width           =   345
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1365
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '�ׯ�
      Height          =   3870
      ItemData        =   "F1040191.frx":0000
      Left            =   5055
      List            =   "F1040191.frx":0007
      TabIndex        =   4
      Top             =   1500
      Width           =   7365
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   5355
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   7770
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   3165
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
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
      Left            =   10290
      TabIndex        =   16
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
      Left            =   9450
      TabIndex        =   15
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
      Left            =   8610
      TabIndex        =   14
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
      Left            =   7770
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��  �V"
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
      Left            =   6510
      TabIndex        =   12
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
      Left            =   5670
      TabIndex        =   11
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
      Left            =   4830
      TabIndex        =   10
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
      Left            =   3990
      TabIndex        =   9
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
      Left            =   2625
      TabIndex        =   8
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
      Left            =   1785
      TabIndex        =   7
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
      Left            =   945
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Left            =   105
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Y��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   540
      TabIndex        =   50
      Top             =   4140
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�݌ɐ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   9900
      TabIndex        =   49
      Top             =   1260
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�d���P��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   8730
      TabIndex        =   48
      Top             =   1260
      Width           =   960
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�d����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   7785
      TabIndex        =   47
      Top             =   1260
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ד�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   6630
      TabIndex        =   45
      Top             =   1260
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   5160
      TabIndex        =   43
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�݌ɐ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   525
      TabIndex        =   41
      Top             =   3600
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�d���P��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   315
      TabIndex        =   39
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�d����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   525
      TabIndex        =   36
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2625
      TabIndex        =   34
      Top             =   1800
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1995
      TabIndex        =   32
      Top             =   1800
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ד�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   525
      TabIndex        =   30
      Top             =   1800
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3045
      TabIndex        =   28
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2415
      TabIndex        =   26
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1785
      TabIndex        =   24
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   735
      TabIndex        =   22
      Top             =   1200
      Width           =   540
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
      TabIndex        =   21
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   525
      TabIndex        =   20
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�����j"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3990
      TabIndex        =   19
      Top             =   720
      Width           =   1485
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7140
      TabIndex        =   18
      Top             =   720
      Width           =   750
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i��(�O��)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   17
      Top             =   720
      Width           =   1275
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1040191"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbNAIGAI% = 0           '�����O

Private Const ptxHin_Gai% = 0           '�i�ԁi�O���j
Private Const ptxHin_Nai% = 1           '�i�ԁi�����j
Private Const ptxHin_Name% = 2          '�i��

Private Const ptxSOKO% = 3              '�q��
Private Const ptxRETU% = 4              '��
Private Const ptxREN% = 5               '�A
Private Const ptxDAN% = 6               '�i
    
Private Const ptxNYUKA_YY% = 7          '���ד��@�N
Private Const ptxNYUKA_MM% = 8          '���ד��@��
Private Const ptxNYUKA_DD% = 9          '���ד��@��
Private Const ptxSHIIRE_CODE% = 10      '�d����
Private Const ptxSHIIRE_TANKA% = 11     '�d���P��

Private Const ptxZAIKO_QTY% = 12        '�݌ɐ�

Private Const ptxGOODS_ON% = 13         '���i��

Private Const ptxGENSANKOKU% = 14       '���Y��


Private Const Text_Max% = 14

    
    
    
Private Const pcmbUKEHARAI% = 0         '�d�������
    
    
Private Const pLstZaiko% = 0            '�݌�ؽ�
    

Dim WS_NO   As String * 3


Private Const LAST_UPDATE_DAY$ = "[���Y���Ή�](F104019 2010.08.24 10:00)"


Private Function List_Dsp() As Integer

Dim sts             As Integer
Dim com             As Integer
Dim i               As Integer
Dim NAIGAI          As String * 1
Dim RetBuf          As String
Dim Edit            As String
    
    
    List_Dsp = True
    
    Call Input_Lock
    
    List1.Clear
    
    
    
    If Combo(pcmbNAIGAI).Text = NAIGAI1 Then
        NAIGAI = NAIGAI_NAI
    Else
        NAIGAI = NAIGAI_GAI
    End If
        
                                            '�݌Ƀf�[�^�Ǎ���
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI)
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, RTrim(Text(ptxHin_Gai).Text))
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")
    
    com = BtOpGetGreaterEqual
    
    
    Do
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> NAIGAI Or _
                    RTrim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> RTrim(Text(ptxHin_Gai).Text) Then
                    Exit Do
                End If
                        
                        
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                    Edit = "*"
                Else
                    Edit = " "
                End If
                            
                        
                Edit = Edit & StrConv(ZAIKOREC.Soko_No, vbUnicode) & "-" & StrConv(ZAIKOREC.Retu, vbUnicode) & "-" & StrConv(ZAIKOREC.Ren, vbUnicode) & "-" & StrConv(ZAIKOREC.Dan, vbUnicode) & " "
                Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2) & "  "
                
                Edit = Edit & StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode) & "  "
                
                If IsNumeric(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)) Then
                    RetBuf = Format(CCur(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)), "#,##0.00")
                Else
                    RetBuf = "0.00"
                End If
                
                If Len(Trim(RetBuf)) < 7 Then
                    RetBuf = Space(7 - Len(Trim(RetBuf))) & Trim(RetBuf)
                End If
                Edit = Edit + RetBuf + "  "
                
                
                
                
                
                RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
                If Len(Trim(RetBuf)) < 7 Then
                    RetBuf = Space(7 - Len(Trim(RetBuf))) & Trim(RetBuf)
                End If
                Edit = Edit + RetBuf + "  "
                
                '2010.08.23
                Edit = Edit & Trim(StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))
                
                
                List1.AddItem Edit
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                List_Dsp = True
                Exit Function
        End Select
        
        
        com = BtOpGetNext
    
    Loop
    
    
    Call Input_UnLock
    
    List_Dsp = False

End Function
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)

Dim i   As Integer


    For i = Mode To Text_Max
            Text(i).Text = ""
    Next i
    
    List1.Clear
End Sub

                                    '�i�ڃ}�X�^���e���ڂ�\������
Private Function Item_Read_Proc() As Integer

Dim sts     As Integer
Dim NAIGAI  As String * 1
Dim i       As Integer


    Item_Read_Proc = True
                                                
                                                
    For i = ptxSOKO To ptxGENSANKOKU
    
        Text(i).Text = ""
    
    Next i

    Combo1(pcmbUKEHARAI).ListIndex = -1
                                                
                                                
                                                
                                                
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
            
        
            Text(ptxHin_Nai).Text = Trim(StrConv(ITEMREC.HIN_NAI, vbUnicode))
            Text(ptxHin_Name).Text = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))

        
        Case BtErrKeyNotFound
                    
    
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Item_Read_Proc = SYS_ERR
            Exit Function
    End Select
            
            
            
            
    Item_Read_Proc = False

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1040191.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040191)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1040191)


    F1040191.MousePointer = vbDefault

End Sub


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbNAIGAI
            Call Clear_Field(0)
            Text(ptxHin_Gai).SetFocus
    End Select

End Sub



Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
        
    Select Case Index
        Case pcmbUKEHARAI       '��z��
            Text(ptxSHIIRE_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).Text, 5))
            Text(ptxSHIIRE_TANKA).SetFocus
    End Select

End Sub

Private Sub Command_Click(Index As Integer)

Dim yn  As Integer
Dim sts As Integer
    
Dim i   As Integer
    
Dim svGENSANKOKU    As String
    
    Select Case Index
        
        
                
        Case 0                              '�X�V
        
            For i = ptxHin_Gai To ptxGENSANKOKU
            
                If Error_Check_Proc(i) Then
                    Exit Sub
                End If
            
            Next i
        
            
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                                                        '�݌Ƀf�[�^�t�@�C���ǂݍ���
                Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)             '���ƕ�
                Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)              '�����O
                Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHin_Gai).Text)  '�i�ԁi�O���j
                                                                            '���i�^�����i
                Call UniCode_Conv(K1_ZAIKO.GOODS_ON, Text(ptxGOODS_ON).Text)
                                                                            '���ד�
                Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, StrConv(ZAIKOREC.NYUKA_DT, vbUnicode))
                Call UniCode_Conv(K1_ZAIKO.Soko_No, Text(ptxSOKO).Text)     '�I�ԁ@�q��
                Call UniCode_Conv(K1_ZAIKO.Retu, Text(ptxRETU).Text)        '      ��
                Call UniCode_Conv(K1_ZAIKO.Ren, Text(ptxREN).Text)          '      �A
                Call UniCode_Conv(K1_ZAIKO.Dan, Text(ptxDAN).Text)          '      �i
            
                sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
                Select Case sts
                    Case BtNoErr
            
                        svGENSANKOKU = StrConv(ZAIKOREC.GENSANKOKU, vbUnicode)
            
                        sts = BTRV(BtOpDelete, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
            
                        If sts <> BtNoErr Then
                            Call File_Error(sts, BtOpDelete, "�݌Ƀf�[�^")
            
                            Unload Me
                        End If
            
            
            
            
            
            
                    Case BtErrKeyNotFound
                        Beep
                        MsgBox "�f�[�^���e���ύX����Ă��܂��B"
            
            
                        If List_Dsp() Then
                            Unload Me
                        End If
                        
                        If List1.ListCount < 1 Then
                            Beep
                            MsgBox "���͂������ڂ̓G���[�ł�� (�Y���݌ɂȂ�)"
                            Text(Index).SetFocus
                            Exit Sub
                        End If
                        
                        List1.SetFocus
                        List1.ListIndex = 0
            
            
            
            
            
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
        
                        Unload Me
        
                End Select
                
                
                
                
                Call UniCode_Conv(ZAIKOREC.NYUKA_DT, Text(ptxNYUKA_YY).Text & Text(ptxNYUKA_MM).Text & Text(ptxNYUKA_DD).Text)
                Call UniCode_Conv(ZAIKOREC.SHIIRE_CODE, Text(ptxSHIIRE_CODE).Text)
                
                Call UniCode_Conv(ZAIKOREC.SHIIRE_TANKA, Format(CCur(Text(ptxSHIIRE_TANKA).Text), "00000000.00"))
                
            
            
                Call UniCode_Conv(ZAIKOREC.GENSANKOKU, Text(ptxGENSANKOKU).Text) '2010.08.23
            
                sts = BTRV(BtOpInsert, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    
                If sts <> BtNoErr Then
                    Call File_Error(sts, BtOpInsert, "�݌Ƀf�[�^")
    
                    Unload Me
                End If
                
                
                Call LOG_OUT(App.EXEName & ".txt", StrConv(ZAIKOREC.JGYOBU, vbUnicode) & "-" & _
                                                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) & "-" & _
                                                    StrConv(ZAIKOREC.HIN_GAI, vbUnicode) & " " & _
                                                    StrConv(ZAIKOREC.Soko_No, vbUnicode) & "-" & _
                                                    StrConv(ZAIKOREC.Retu, vbUnicode) & "-" & _
                                                    StrConv(ZAIKOREC.Ren, vbUnicode) & "-" & _
                                                    StrConv(ZAIKOREC.Dan, vbUnicode) & " " & _
                                                    svGENSANKOKU & "-->" & _
                                                    StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))


                If List_Dsp() Then
                    Unload Me
                End If
                
                If List1.ListCount < 1 Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� (�Y���݌ɂȂ�)"
                    Text(Index).SetFocus
                    Exit Sub
                End If
                
                List1.SetFocus
                List1.ListIndex = 0
            
            
            
            
                For i = ptxSOKO To ptxGENSANKOKU
                
                    Text(i).Text = ""
                
                Next i
            
                Combo1(pcmbUKEHARAI).ListIndex = -1
            
            
            
            
            
            End If
        
        
        
        
        
        
        
        
        
        Case 7                              '�ŐV�\��
            
            
            For i = ptxSOKO To ptxGENSANKOKU
            
                Text(i).Text = ""
            
            Next i
        
            Combo1(pcmbUKEHARAI).ListIndex = -1
            
            
    
            Text(ptxHin_Gai).Text = StrConv(RTrim(Text(ptxHin_Gai).Text), vbUpperCase)
    
            
            
            sts = Item_Read_Proc()
            Select Case sts
                Case False
                Case True
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� (���o�^�G���[)"
                    Text(ptxHin_Gai).SetFocus
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
                    
            If List_Dsp() Then
                Unload Me
            End If
            
            If List1.ListCount < 1 Then
                Beep
                MsgBox "���͂������ڂ̓G���[�ł�� (�Y���݌ɂȂ�)"
                Text(Index).SetFocus
                Exit Sub
            End If
            
            List1.SetFocus
            List1.ListIndex = 0
                        
        Case 11                             '�I��
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
    
    Show
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

        Last_JGYOBU = SHIZAI

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1040191.Caption = "�݌ɒP�������e�i���X�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

'�[���ԍ���荞��
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = RTrim(com)
                                
                                '�݌Ƀf�[�^OPEN
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^OPEN
    If wZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '���Y���}�X�^�n�o�d�m       '2010.08.23
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '��ʏ����ݒ�
    Call Clear_Field(0)
    
    
                                '�����O��荞��
    Combo(pcmbNAIGAI).AddItem NAIGAI1
    Combo(pcmbNAIGAI).AddItem NAIGAI2
    Combo(pcmbNAIGAI).Text = NAIGAI1
    
    
    '�󕥐�
    If Ukeharai_Set_Proc() Then
        Unload Me
    End If
    
    
    
    Combo(pcmbNAIGAI).SetFocus
    
    End Sub



Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            
                                            
                                            
                                        '�݌Ƀf�[�^�g�p������
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
    End If
                                            
                                            
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    If wZAIKO_CLOSE() Then
    End If

'�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�󕥐�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1040191 = Nothing

    End
End Sub



Private Sub List1_DblClick()

Dim LOCATION    As String * 8
Dim END_FLG     As Boolean
Dim sts         As Integer

Dim NYUKA_YMD   As String * 8
Dim i           As Integer


    Call Input_Lock
                                        
                                        
    END_FLG = False
                                        
                                        '�g�����U�N�V�����J�n
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Unload Me
    End If
                                        '�݌Ƀf�[�^�g�p������
    If Zaiko_UNLock_Proc("", "", "", "", WS_NO) Then
        END_FLG = True
        GoTo Abort_Tran
    End If
                                        '���P�[�V�����l��
    
    If Mid(List1.List(List1.ListIndex), 1, 1) = "*" Then
        Text(ptxGOODS_ON).Text = GOODS_ON
    Else
        Text(ptxGOODS_ON).Text = GOODS_OFF
    End If
    
    LOCATION = Mid(List1.List(List1.ListIndex), 2, 2) & _
                Mid(List1.List(List1.ListIndex), 5, 2) & _
                Mid(List1.List(List1.ListIndex), 8, 2) & _
                Mid(List1.List(List1.ListIndex), 11, 2)

    sts = Zaiko_Lock_Proc(LOCATION, Last_JGYOBU, NAIGAI_NAI, Text(ptxHin_Gai).Text, WS_NO)
    Select Case sts
        Case False
        Case True, SYS_CANCEL
            GoTo Abort_Tran
        Case SYS_ERR
            END_FLG = True
            GoTo Abort_Tran
    End Select
                                                
    NYUKA_YMD = Mid(List1.List(List1.ListIndex), 14, 4) & _
                                            Mid(List1.List(List1.ListIndex), 19, 2) & _
                                            Mid(List1.List(List1.ListIndex), 22, 2)
                                                
                                                '�݌Ƀf�[�^�t�@�C���ǂݍ���
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)                 '���ƕ�
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, NAIGAI_NAI)                  '�����O
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Text(ptxHin_Gai).Text)      '�i�ԁi�O���j
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, Text(ptxGOODS_ON).Text)    '���i�^�����i
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, NYUKA_YMD)                 '���ד�
    Call UniCode_Conv(K1_ZAIKO.Soko_No, Mid(LOCATION, 1, 2))        '�I�ԁ@�q��
    Call UniCode_Conv(K1_ZAIKO.Retu, Mid(LOCATION, 3, 2))           '      ��
    Call UniCode_Conv(K1_ZAIKO.Ren, Mid(LOCATION, 5, 2))            '      �A
    Call UniCode_Conv(K1_ZAIKO.Dan, Mid(LOCATION, 7, 2))            '      �i
        
    sts = BTRV(BtOpGetEqual, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
    Select Case sts
        Case BtNoErr
        
            Text(ptxSOKO).Text = StrConv(ZAIKOREC.Soko_No, vbUnicode)
            Text(ptxRETU).Text = StrConv(ZAIKOREC.Retu, vbUnicode)
            Text(ptxREN).Text = StrConv(ZAIKOREC.Ren, vbUnicode)
            Text(ptxDAN).Text = StrConv(ZAIKOREC.Dan, vbUnicode)
        
            Text(ptxNYUKA_YY).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4)
            Text(ptxNYUKA_MM).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2)
            Text(ptxNYUKA_DD).Text = Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2)
        
            Text(ptxSHIIRE_CODE).Text = RTrim(StrConv(ZAIKOREC.SHIIRE_CODE, vbUnicode))
                    
            Combo1(pcmbUKEHARAI).ListIndex = -1
            For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
            
Debug.Print RTrim(Right(Combo1(pcmbUKEHARAI).List(i), 5))
            
                If RTrim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) = Text(ptxSHIIRE_CODE).Text Then
                    Combo1(pcmbUKEHARAI).ListIndex = i
                    Exit For
                End If
            
            
            Next i
        
        
        
        
            If IsNumeric(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)) Then
                Text(ptxSHIIRE_TANKA).Text = Format(CCur(StrConv(ZAIKOREC.SHIIRE_TANKA, vbUnicode)), "#0.00")
            Else
                Text(ptxSHIIRE_TANKA).Text = "0.00"
            End If
        
            Text(ptxZAIKO_QTY).Text = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
        
            '2010.08.23
            Text(ptxGENSANKOKU).Text = Trim(StrConv(ZAIKOREC.GENSANKOKU, vbUnicode))
        
        
        Case BtErrKeyNotFound
            Beep
            MsgBox "�f�[�^���e���ύX����Ă��܂��B�u�ŐV�v�\����I�����Ă��������B"
            GoTo Abort_Tran
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
            END_FLG = True
            GoTo Abort_Tran
    End Select
                                        '�g�����U�N�V�����I��

End_Tran:
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        END_FLG = True
        GoTo Abort_Tran
    End If
    


    Call Input_UnLock

    Text(ptxNYUKA_YY).SetFocus

    Exit Sub

Abort_Tran:
    Call Input_UnLock
    
    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
        Unload Me
    End If

    If END_FLG Then
        Unload Me
    End If
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
    F1040191.Caption = "�݌ɒP�������e�i���X�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
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

Dim i As Integer
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub

    If Index = 0 Or Index = 1 Then
    
        Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
    
    End If



    Select Case Index
        Case ptxHin_Gai             '�O���i��
            
            If Len(Trim(Text(ptxHin_Gai).Text)) <> 0 Then
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
                        
                If List_Dsp() Then
                    Unload Me
                End If
                        
                If List1.ListCount < 1 Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� (�Y���݌ɂȂ�)"
                    Text(Index).SetFocus
                    Exit Sub
                End If
                
                List1.SetFocus
                List1.ListIndex = 0
            
                Exit Sub
            
            End If
    
    
    
        Case ptxSHIIRE_CODE
            Combo1(pcmbUKEHARAI).ListIndex = -1
            
            If Trim(Text(ptxSHIIRE_CODE).Text) = "" Then
            Else
               For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                   If Text(ptxSHIIRE_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                       Combo1(pcmbUKEHARAI).ListIndex = i
                       Exit For
                   End If
               
               Next i
        
               If i > Combo1(pcmbUKEHARAI).ListCount - 1 Then
                   MsgBox "�d����G���[�ł��B"
                   Text(ptxSHIIRE_CODE).SetFocus
                   Exit Sub
               End If
            End If
    
        Case ptxGENSANKOKU      '���Y�� 2010.08.23
            
            If Trim(Text(ptxGENSANKOKU).Text) = "" Then
            Else
    
                Call UniCode_Conv(K0_GENSAN.JGYOBU, (StrConv(ITEMREC.JGYOBU, vbUnicode)))
                Call UniCode_Conv(K0_GENSAN.NAIGAI, (StrConv(ITEMREC.NAIGAI, vbUnicode)))
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, (StrConv(ITEMREC.HIN_GAI, vbUnicode)))
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, Text(ptxGENSANKOKU).Text)

                sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                    
                        MsgBox "���Y��Ͻ����o�^�ł��B"
                        Text(Index).SetFocus
                        Exit Sub
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���Y��Ͻ�")
                        Unload Me
                End Select
            End If
    
    
    
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).TabStop And Text(i).Visible Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub


Private Sub Text_LostFocus(Index As Integer)
    
    If Index = 0 Or Index = 1 Then
    
        Text(Index).Text = StrConv(RTrim(Text(Index).Text), vbUpperCase)
    
    End If

End Sub

Private Function Ukeharai_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   �󕥐�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer




Dim i           As Integer
    
    Ukeharai_Set_Proc = True
    
    Combo1(pcmbUKEHARAI).Clear
    
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�󕥐�}�X�^")
                Exit Function
        
        End Select

        
        
        Combo1(pcmbUKEHARAI).AddItem StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode) & "            " & _
                                StrConv(P_UKEHARAIREC.UKEHARAI_CODE, vbUnicode)
        
        com = BtOpGetNext
    
    Loop

    Ukeharai_Set_Proc = False
    



End Function


Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts     As Integer
Dim i       As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxHin_Gai    '�i��
        
            If RTrim(Text(Mode).Text) <> RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                MsgBox "�Ώەi�Ԃ��ύX����Ă܂���ēx���͂��s���Ă�������� """
                Text(Mode).SetFocus
                Exit Function
            End If
        
            If List1.ListCount = 0 Then
                MsgBox "�Ώەi�Ԃ��ύX����Ă܂��B�ēx���͂��s���Ă��������B"
                Text(Mode).SetFocus
                Exit Function
            End If
        
        
        Case ptxSOKO, ptxRETU, ptxREN, ptxDAN
            If Trim(Text(Mode)) = "" Then
                MsgBox "�Ώۍ݌ɂ��ύX����Ă܂���ēx���͂��s���Ă�������� """
                
                
                If List1.ListCount < 1 Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł�� (�Y���݌ɂȂ�)"
                    Text(ptxHin_Gai).SetFocus
                    Exit Function
                End If
                
                List1.SetFocus
                List1.ListIndex = 0
                
                
                Exit Function
            End If
        
        
        Case ptxNYUKA_YY
        Case ptxNYUKA_MM
            If IsNumeric(Text(Mode).Text) Then
                Text(Mode).Text = Format(Val(Text(Mode).Text), "00")
            End If
        Case ptxNYUKA_DD
        
            If IsNumeric(Text(Mode).Text) Then
                Text(Mode).Text = Format(Val(Text(Mode).Text), "00")
            End If
        
        
        
            If Not IsDate(Text(ptxNYUKA_YY).Text & "/" & Text(ptxNYUKA_MM).Text & "/" & Text(ptxNYUKA_DD).Text) Then
            
                MsgBox "���ד��G���[�ł�"
                Text(ptxNYUKA_YY).SetFocus
                Exit Function
            
            
            End If
        
            If (Text(ptxNYUKA_YY).Text & Text(ptxNYUKA_MM).Text & Text(ptxNYUKA_DD).Text) <> StrConv(ZAIKOREC.NYUKA_DT, vbUnicode) Then
            
            
            
            
            
                Call UniCode_Conv(K1_wZAIKO.JGYOBU, Last_JGYOBU)                '���ƕ�
                Call UniCode_Conv(K1_wZAIKO.NAIGAI, NAIGAI_NAI)                 '�����O
                Call UniCode_Conv(K1_wZAIKO.HIN_GAI, Text(ptxHin_Gai).Text)     '�i�ԁi�O���j
                                                                                '���i�^�����i
                Call UniCode_Conv(K1_wZAIKO.GOODS_ON, Text(ptxGOODS_ON).Text)
                                                                                '���ד�
                Call UniCode_Conv(K1_wZAIKO.NYUKA_DT, (Text(ptxNYUKA_YY).Text & Text(ptxNYUKA_MM).Text & Text(ptxNYUKA_DD).Text))
                Call UniCode_Conv(K1_wZAIKO.Soko_No, Text(ptxSOKO).Text)        '�I�ԁ@�q��
                Call UniCode_Conv(K1_wZAIKO.Retu, Text(ptxRETU).Text)           '      ��
                Call UniCode_Conv(K1_wZAIKO.Ren, Text(ptxREN).Text)             '      �A
                Call UniCode_Conv(K1_wZAIKO.Dan, Text(ptxDAN).Text)             '      �i
                    
                sts = BTRV(BtOpGetEqual, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K1_wZAIKO, Len(K1_wZAIKO), 1)
                Select Case sts
                    Case BtNoErr
                    
                        Beep
                        MsgBox "���ד��o�^�ςł��B"
                    
                        Text(ptxNYUKA_YY).SetFocus
                        Exit Function
                    
                    
                    Case BtErrKeyNotFound
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�݌Ƀf�[�^")
                        Exit Function
                End Select
            
            
            
            
            
            End If
        Case ptxSHIIRE_CODE   '��z��
            
            Combo1(pcmbUKEHARAI).ListIndex = -1
            If Trim(Text(ptxSHIIRE_CODE).Text) = "" Then
            Else
            
            
               For i = 0 To Combo1(pcmbUKEHARAI).ListCount - 1
                   If Text(ptxSHIIRE_CODE).Text = Trim(Right(Combo1(pcmbUKEHARAI).List(i), 5)) Then
                       Combo1(pcmbUKEHARAI).ListIndex = i
                       Exit For
                   End If
               
               Next i
        
               If i > Combo1(pcmbUKEHARAI).ListCount - 1 Then
                   MsgBox "�d����G���[�ł��B"
                   Text(Mode).SetFocus
                   Exit Function
               End If
            End If
        
        Case ptxSHIIRE_TANKA    '�P��
        
            If Trim(Trim(Text(Mode).Text)) = "" Then
            Else
                    
                If Not IsNumeric(Trim(Text(Mode).Text)) Then
                    MsgBox "�d����P���G���[�ł��B"
                    Text(Mode).SetFocus
                    Exit Function
                End If
        
                Text(Mode).Text = Format(CCur(Text(Mode).Text), "0.00")
                
                
                If CCur(Text(Mode).Text) < 0 Then
                    MsgBox "�d����P���G���[�ł��B"
                    Text(Mode).SetFocus
                    Exit Function
                End If
            End If
    
        Case ptxGENSANKOKU      '���Y�� 2010.08.23
    
    
            If Trim(Text(ptxGENSANKOKU).Text) = "" Then
            Else
    
                Call UniCode_Conv(K0_GENSAN.JGYOBU, (StrConv(ITEMREC.JGYOBU, vbUnicode)))
                Call UniCode_Conv(K0_GENSAN.NAIGAI, (StrConv(ITEMREC.NAIGAI, vbUnicode)))
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, (StrConv(ITEMREC.HIN_GAI, vbUnicode)))
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, Text(ptxGENSANKOKU).Text)

                sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                    
                        MsgBox "���Y��Ͻ����o�^�ł��B"
                        Text(Mode).SetFocus
'                        Exit Function
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���Y��Ͻ�")
                        Exit Function
                End Select
            End If
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function


