VERSION 5.00
Begin VB.Form PI000192 
   Appearance      =   0  '�ׯ�
   Caption         =   "���i���w�}�����s���\�����i�o�^��"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15945
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
   ScaleHeight     =   10305
   ScaleWidth      =   15945
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   174
      Left            =   10440
      TabIndex        =   199
      Top             =   9000
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   173
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   198
      TabStop         =   0   'False
      Top             =   9000
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   172
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   197
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   171
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   196
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   170
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   195
      Top             =   9000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   169
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   9000
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   168
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   193
      Top             =   9000
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   24
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   192
      Top             =   9000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   167
      Left            =   10440
      TabIndex        =   191
      Top             =   8640
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   166
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   8640
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   165
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   8640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   164
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   8640
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   163
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   187
      Top             =   8640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   162
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   186
      TabStop         =   0   'False
      Top             =   8640
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   161
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   185
      Top             =   8640
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   23
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   184
      Top             =   8640
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   160
      Left            =   10440
      TabIndex        =   183
      Top             =   8280
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   159
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   182
      TabStop         =   0   'False
      Top             =   8280
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   158
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   181
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   157
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   180
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   156
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   179
      Top             =   8280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   155
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   178
      TabStop         =   0   'False
      Top             =   8280
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   154
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   177
      Top             =   8280
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   22
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   176
      Top             =   8280
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   153
      Left            =   10440
      TabIndex        =   175
      Top             =   7920
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   152
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   174
      TabStop         =   0   'False
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   151
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   173
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   150
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   172
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   149
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   171
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   148
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   170
      TabStop         =   0   'False
      Top             =   7920
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   147
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   169
      Top             =   7920
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   21
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   168
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   146
      Left            =   10440
      TabIndex        =   167
      Top             =   7560
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   145
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   144
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   143
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   164
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   142
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   163
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   141
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   162
      TabStop         =   0   'False
      Top             =   7560
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   140
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   161
      Top             =   7560
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   20
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   160
      Top             =   7560
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   139
      Left            =   10440
      TabIndex        =   159
      Top             =   7200
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   138
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   137
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   136
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   135
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   155
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   134
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   7200
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   133
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   153
      Top             =   7200
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   19
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   152
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   132
      Left            =   10440
      TabIndex        =   151
      Top             =   6840
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   131
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   150
      TabStop         =   0   'False
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   130
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   129
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   148
      TabStop         =   0   'False
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   128
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   147
      Top             =   6840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   127
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   6840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   126
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   145
      Top             =   6840
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   18
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   144
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   125
      Left            =   10440
      TabIndex        =   143
      Top             =   6480
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   124
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   123
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   122
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   121
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   139
      Top             =   6480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   120
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   6480
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   119
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   137
      Top             =   6480
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   17
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   136
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   118
      Left            =   10440
      TabIndex        =   135
      Top             =   6120
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   117
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   116
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   115
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   114
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   131
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   113
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   6120
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   112
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   129
      Top             =   6120
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   16
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   128
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   111
      Left            =   10440
      TabIndex        =   127
      Top             =   5760
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   110
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   109
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   108
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   107
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   123
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   106
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   5760
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   105
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   121
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   15
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   120
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   104
      Left            =   10440
      TabIndex        =   119
      Top             =   5400
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   103
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   102
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   101
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   100
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   115
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   99
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   98
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   113
      Top             =   5400
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   14
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   112
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   97
      Left            =   10440
      TabIndex        =   111
      Top             =   5040
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   96
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   95
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   94
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   93
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   107
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   92
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   91
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   105
      Top             =   5040
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   13
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   104
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   90
      Left            =   10440
      TabIndex        =   103
      Top             =   4680
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   89
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   88
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   87
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   86
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   99
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   85
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   84
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   97
      Top             =   4680
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   12
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   96
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   83
      Left            =   10440
      TabIndex        =   95
      Top             =   4320
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   82
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   81
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   80
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   79
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   91
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   78
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   77
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   89
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   11
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   88
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   76
      Left            =   10440
      TabIndex        =   87
      Top             =   3960
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   75
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   74
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   73
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   72
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   83
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   71
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   70
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   81
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   10
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   80
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   69
      Left            =   10440
      TabIndex        =   79
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   68
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   67
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   66
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   65
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   75
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   64
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   63
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   73
      Top             =   3600
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   9
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   72
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   62
      Left            =   10440
      TabIndex        =   71
      Top             =   3240
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   61
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   60
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   59
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   58
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   67
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   57
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   56
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   65
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   8
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   64
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   55
      Left            =   10440
      TabIndex        =   63
      Top             =   2880
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   54
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   53
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   52
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   51
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   59
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   50
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   49
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   57
      Top             =   2880
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   7
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   56
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   48
      Left            =   10440
      TabIndex        =   55
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   47
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   46
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   45
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   44
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   51
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   43
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   42
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   49
      Top             =   2520
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   6
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   48
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   41
      Left            =   10440
      TabIndex        =   47
      Top             =   2160
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   40
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   39
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   38
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   37
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   43
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   36
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   35
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   41
      Top             =   2160
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   5
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   40
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   34
      Left            =   10440
      TabIndex        =   39
      Top             =   1800
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   33
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   32
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   31
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   30
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   35
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   29
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   28
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   33
      Top             =   1800
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   4
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   32
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   27
      Left            =   10440
      TabIndex        =   31
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   26
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   25
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   24
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   23
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   27
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   22
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   21
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   25
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   3
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   24
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   20
      Left            =   10440
      TabIndex        =   23
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   19
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   18
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   17
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   16
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   15
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   14
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   17
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   2
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   16
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   13
      Left            =   10440
      TabIndex        =   15
      Top             =   720
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   12
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   11
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   10
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   8
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   9
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   1
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   1
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   3
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   4
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      Index           =   5
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   10440
      TabIndex        =   7
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�߂�"
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
      Left            =   10440
      TabIndex        =   211
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9600
      TabIndex        =   210
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8760
      TabIndex        =   209
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   7920
      TabIndex        =   208
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   6600
      TabIndex        =   207
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "26-50"
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
      Left            =   5760
      TabIndex        =   206
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1-25"
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
      Left            =   4920
      TabIndex        =   205
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4080
      TabIndex        =   204
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   2760
      TabIndex        =   203
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   1920
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   9840
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
      Left            =   1080
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   9840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�m �F"
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
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   9840
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���"
      Height          =   255
      Index           =   12
      Left            =   480
      TabIndex        =   244
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      Height          =   255
      Index           =   14
      Left            =   1680
      TabIndex        =   243
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   255
      Index           =   15
      Left            =   6720
      TabIndex        =   242
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�݌�"
      Height          =   255
      Index           =   18
      Left            =   9960
      TabIndex        =   241
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "���l"
      Height          =   255
      Index           =   19
      Left            =   11160
      TabIndex        =   240
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�i��"
      Height          =   255
      Index           =   21
      Left            =   3960
      TabIndex        =   239
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����"
      Height          =   255
      Index           =   22
      Left            =   7800
      TabIndex        =   238
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�I��"
      Height          =   255
      Index           =   16
      Left            =   8760
      TabIndex        =   237
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   24
      Left            =   120
      TabIndex        =   236
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   23
      Left            =   120
      TabIndex        =   235
      Top             =   8760
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   234
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   233
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   232
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   231
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   230
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   229
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   228
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   227
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   226
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   225
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   224
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   223
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   222
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   221
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   220
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   219
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   218
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   217
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   216
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   215
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   214
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   213
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblNumber 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   212
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "PI000192"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Init_Flg    As Boolean

'�e�L�X�g�p�Y��
Private Const ptxD_HIN_GAI01% = 0           '1�@�����^�\���i��
Private Const ptxD_HIN_NAME01% = 1          '1�@�����^�\���i��
Private Const ptxD_QTY01% = 2               '1�@����
Private Const ptxD_SHIJI_QTY01% = 3         '1�@����
Private Const ptxD_ST_LOCATION01% = 4       '1�@�I��
Private Const ptxD_ZAIKO_QTY01% = 5         '1�@�݌ɐ�
Private Const ptxD_BIKOU01% = 6             '1�@���l

Private Const ptxD_HIN_GAI02% = 7           '2�@�����^�\���i��
Private Const ptxD_HIN_NAME02% = 8          '2�@�����^�\���i��
Private Const ptxD_QTY02% = 9               '2�@����
Private Const ptxD_SHIJI_QTY02% = 10        '2�@����
Private Const ptxD_ST_LOCATION02% = 11      '2�@�I��
Private Const ptxD_ZAIKO_QTY02% = 12        '2�@�݌ɐ�
Private Const ptxD_BIKOU02% = 13            '2�@���l

Private Const ptxD_HIN_GAI03% = 14          '3�@�����^�\���i��
Private Const ptxD_HIN_NAME03% = 15         '3�@�����^�\���i��
Private Const ptxD_QTY03% = 16              '3�@����
Private Const ptxD_SHIJI_QTY03% = 17        '3�@����
Private Const ptxD_ST_LOCATION03% = 18      '3�@�I��
Private Const ptxD_ZAIKO_QTY03% = 19        '3�@�݌ɐ�
Private Const ptxD_BIKOU03% = 20            '3�@���l

Private Const ptxD_HIN_GAI04% = 21          '4  �����^�\���i��
Private Const ptxD_HIN_NAME04% = 22         '4�@�����^�\���i��
Private Const ptxD_QTY04% = 23              '4�@����
Private Const ptxD_SHIJI_QTY04% = 24        '4�@����
Private Const ptxD_ST_LOCATION04% = 25      '4�@�I��
Private Const ptxD_ZAIKO_QTY04% = 26        '4�@�݌ɐ�
Private Const ptxD_BIKOU04% = 27            '4�@���l

Private Const ptxD_HIN_GAI05% = 28          '5  �����^�\���i��
Private Const ptxD_HIN_NAME05% = 29         '5�@�����^�\���i��
Private Const ptxD_QTY05% = 30              '5�@����
Private Const ptxD_SHIJI_QTY05% = 31        '5�@����
Private Const ptxD_ST_LOCATION05% = 32      '5�@�I��
Private Const ptxD_ZAIKO_QTY05% = 33        '5�@�݌ɐ�
Private Const ptxD_BIKOU05% = 34            '5�@���l

Private Const ptxD_HIN_GAI06% = 35          '6  �����^�\���i��
Private Const ptxD_HIN_NAME06% = 36         '6�@�����^�\���i��
Private Const ptxD_QTY06% = 37              '6�@����
Private Const ptxD_SHIJI_QTY06% = 38        '6�@����
Private Const ptxD_ST_LOCATION06% = 39      '6�@�I��
Private Const ptxD_ZAIKO_QTY06% = 40        '6�@�݌ɐ�
Private Const ptxD_BIKOU06% = 41            '6�@���l

Private Const ptxD_HIN_GAI07% = 42          '7  �����^�\���i��
Private Const ptxD_HIN_NAME07% = 43         '7�@�����^�\���i��
Private Const ptxD_QTY07% = 44              '7�@����
Private Const ptxD_SHIJI_QTY07% = 45        '7�@����
Private Const ptxD_ST_LOCATION07% = 46      '7�@�I��
Private Const ptxD_ZAIKO_QTY07% = 47        '7�@�݌ɐ�
Private Const ptxD_BIKOU07% = 48            '7�@���l

Private Const ptxD_HIN_GAI08% = 49          '8  �����^�\���i��
Private Const ptxD_HIN_NAME08% = 50         '8�@�����^�\���i��
Private Const ptxD_QTY08% = 51              '8�@����
Private Const ptxD_SHIJI_QTY08% = 52        '8�@����
Private Const ptxD_ST_LOCATION08% = 53      '8�@�I��
Private Const ptxD_ZAIKO_QTY08% = 54        '8�@�݌ɐ�
Private Const ptxD_BIKOU08% = 55            '8�@���l

Private Const ptxD_HIN_GAI09% = 56          '9  �����^�\���i��
Private Const ptxD_HIN_NAME09% = 57         '9�@�����^�\���i��
Private Const ptxD_QTY09% = 58              '9�@����
Private Const ptxD_SHIJI_QTY09% = 59        '9�@����
Private Const ptxD_ST_LOCATION09% = 60      '9�@�I��
Private Const ptxD_ZAIKO_QTY09% = 61        '9�@�݌ɐ�
Private Const ptxD_BIKOU09% = 62            '9  ���l

Private Const ptxD_HIN_GAI10% = 63          '10  �����^�\���i��
Private Const ptxD_HIN_NAME10% = 64         '10�@�����^�\���i��
Private Const ptxD_QTY10% = 65              '10�@����
Private Const ptxD_SHIJI_QTY10% = 66        '10�@����
Private Const ptxD_ST_LOCATION10% = 67      '10�@�I��
Private Const ptxD_ZAIKO_QTY10% = 68        '10�@�݌ɐ�
Private Const ptxD_BIKOU10% = 69            '10  ���l

Private Const ptxD_HIN_GAI11% = 70          '11  �����^�\���i��
Private Const ptxD_HIN_NAME11% = 71         '11�@�����^�\���i��
Private Const ptxD_QTY11% = 72              '11�@����
Private Const ptxD_SHIJI_QTY11% = 73        '11�@����
Private Const ptxD_ST_LOCATION11% = 74      '11�@�I��
Private Const ptxD_ZAIKO_QTY11% = 75        '11�@�݌ɐ�
Private Const ptxD_BIKOU11% = 76            '11  ���l

Private Const ptxD_HIN_GAI12% = 77          '12  �����^�\���i��
Private Const ptxD_HIN_NAME12% = 78         '12�@�����^�\���i��
Private Const ptxD_QTY12% = 79              '12�@����
Private Const ptxD_SHIJI_QTY12% = 80        '12�@����
Private Const ptxD_ST_LOCATION12% = 81      '12�@�I��
Private Const ptxD_ZAIKO_QTY12% = 82        '12�@�݌ɐ�
Private Const ptxD_BIKOU12% = 83            '12  ���l

Private Const ptxD_HIN_GAI13% = 84          '13  �����^�\���i��
Private Const ptxD_HIN_NAME13% = 85         '13�@�����^�\���i��
Private Const ptxD_QTY13% = 86              '13�@����
Private Const ptxD_SHIJI_QTY13% = 87        '13�@����
Private Const ptxD_ST_LOCATION13% = 88      '13�@�I��
Private Const ptxD_ZAIKO_QTY13% = 89        '13�@�݌ɐ�
Private Const ptxD_BIKOU13% = 90            '13  ���l

Private Const ptxD_HIN_GAI14% = 91          '14  �����^�\���i��
Private Const ptxD_HIN_NAME14% = 92         '14�@�����^�\���i��
Private Const ptxD_QTY14% = 93              '14�@����
Private Const ptxD_SHIJI_QTY14% = 94        '14�@����
Private Const ptxD_ST_LOCATION14% = 95      '14�@�I��
Private Const ptxD_ZAIKO_QTY14% = 96        '14�@�݌ɐ�
Private Const ptxD_BIKOU14% = 97            '14  ���l

Private Const ptxD_HIN_GAI15% = 98          '15  �����^�\���i��
Private Const ptxD_HIN_NAME15% = 99         '15�@�����^�\���i��
Private Const ptxD_QTY15% = 100             '15�@����
Private Const ptxD_SHIJI_QTY15% = 101       '15�@����
Private Const ptxD_ST_LOCATION15% = 102     '15�@�I��
Private Const ptxD_ZAIKO_QTY15% = 103       '15�@�݌ɐ�
Private Const ptxD_BIKOU15% = 104           '15  ���l


Private Const ptxD_HIN_GAI16% = 105         '16  �����^�\���i��
Private Const ptxD_HIN_NAME16% = 106        '16�@�����^�\���i��
Private Const ptxD_QTY16% = 107             '16�@����
Private Const ptxD_SHIJI_QTY16% = 108       '16�@����
Private Const ptxD_ST_LOCATION16% = 109     '16�@�I��
Private Const ptxD_ZAIKO_QTY16% = 110       '16�@�݌ɐ�
Private Const ptxD_BIKOU16% = 111           '16  ���l

Private Const ptxD_HIN_GAI17% = 112         '17  �����^�\���i��
Private Const ptxD_HIN_NAME17% = 113        '17�@�����^�\���i��
Private Const ptxD_QTY17% = 114             '17�@����
Private Const ptxD_SHIJI_QTY17% = 115       '17�@����
Private Const ptxD_ST_LOCATION17% = 116     '17�@�I��
Private Const ptxD_ZAIKO_QTY17% = 117       '17�@�݌ɐ�
Private Const ptxD_BIKOU17% = 118           '17  ���l

Private Const ptxD_HIN_GAI18% = 119         '18  �����^�\���i��
Private Const ptxD_HIN_NAME18% = 120        '18�@�����^�\���i��
Private Const ptxD_QTY18% = 121             '18�@����
Private Const ptxD_SHIJI_QTY18% = 122       '18�@����
Private Const ptxD_ST_LOCATION18% = 123     '18�@�I��
Private Const ptxD_ZAIKO_QTY18% = 124       '18�@�݌ɐ�
Private Const ptxD_BIKOU18% = 125           '18  ���l

Private Const ptxD_HIN_GAI19% = 126         '19  �����^�\���i��
Private Const ptxD_HIN_NAME19% = 127        '19�@�����^�\���i��
Private Const ptxD_QTY19% = 128             '19�@����
Private Const ptxD_SHIJI_QTY19% = 129       '19�@����
Private Const ptxD_ST_LOCATION19% = 130     '19�@�I��
Private Const ptxD_ZAIKO_QTY19% = 131       '19�@�݌ɐ�
Private Const ptxD_BIKOU19% = 132           '19  ���l

Private Const ptxD_HIN_GAI20% = 133         '20  �����^�\���i��
Private Const ptxD_HIN_NAME20% = 134        '20�@�����^�\���i��
Private Const ptxD_QTY20% = 135             '20�@����
Private Const ptxD_SHIJI_QTY20% = 136       '20�@����
Private Const ptxD_ST_LOCATION20% = 137     '20�@�I��
Private Const ptxD_ZAIKO_QTY20% = 138       '20�@�݌ɐ�
Private Const ptxD_BIKOU20% = 139           '20  ���l

Private Const ptxD_HIN_GAI21% = 140         '21  �����^�\���i��
Private Const ptxD_HIN_NAME21% = 141        '21�@�����^�\���i��
Private Const ptxD_QTY21% = 142             '21�@����
Private Const ptxD_SHIJI_QTY21% = 143       '21�@����
Private Const ptxD_ST_LOCATION21% = 144     '21�@�I��
Private Const ptxD_ZAIKO_QTY21% = 145       '21�@�݌ɐ�
Private Const ptxD_BIKOU21% = 146           '21  ���l

Private Const ptxD_HIN_GAI22% = 147         '22  �����^�\���i��
Private Const ptxD_HIN_NAME22% = 148        '22�@�����^�\���i��
Private Const ptxD_QTY22% = 149             '22�@����
Private Const ptxD_SHIJI_QTY22% = 150       '22�@����
Private Const ptxD_ST_LOCATION22% = 151     '22�@�I��
Private Const ptxD_ZAIKO_QTY22% = 152       '22�@�݌ɐ�
Private Const ptxD_BIKOU22% = 153           '22  ���l

Private Const ptxD_HIN_GAI23% = 154         '23  �����^�\���i��
Private Const ptxD_HIN_NAME23% = 155        '23�@�����^�\���i��
Private Const ptxD_QTY23% = 156             '23�@����
Private Const ptxD_SHIJI_QTY23% = 157       '23�@����
Private Const ptxD_ST_LOCATION23% = 158     '23�@�I��
Private Const ptxD_ZAIKO_QTY23% = 159       '23�@�݌ɐ�
Private Const ptxD_BIKOU23% = 160           '23  ���l

Private Const ptxD_HIN_GAI24% = 161         '24  �����^�\���i��
Private Const ptxD_HIN_NAME24% = 162        '24�@�����^�\���i��
Private Const ptxD_QTY24% = 163             '24�@����
Private Const ptxD_SHIJI_QTY24% = 164       '24�@����
Private Const ptxD_ST_LOCATION24% = 165     '24�@�I��
Private Const ptxD_ZAIKO_QTY24% = 166       '24�@�݌ɐ�
Private Const ptxD_BIKOU24% = 167           '24  ���l

Private Const ptxD_HIN_GAI25% = 168         '25  �����^�\���i��
Private Const ptxD_HIN_NAME25% = 169        '25�@�����^�\���i��
Private Const ptxD_QTY25% = 170             '25�@����
Private Const ptxD_SHIJI_QTY25% = 171       '25�@����
Private Const ptxD_ST_LOCATION25% = 172     '25�@�I��
Private Const ptxD_ZAIKO_QTY25% = 173       '25�@�݌ɐ�
Private Const ptxD_BIKOU25% = 174           '25  ���l




'�R���{�p�Y��
Private Const pcmbD_SYUBETSU01% = 0         '�@�@���
Private Const pcmbD_SYUBETSU02% = 1         '�A�@���
Private Const pcmbD_SYUBETSU03% = 2         '�B�@���
Private Const pcmbD_SYUBETSU04% = 3         '�C�@���
Private Const pcmbD_SYUBETSU05% = 4         '�D�@���
Private Const pcmbD_SYUBETSU06% = 5         '�E�@���
Private Const pcmbD_SYUBETSU07% = 6         '�F�@���
Private Const pcmbD_SYUBETSU08% = 7         '�G�@���
Private Const pcmbD_SYUBETSU09% = 8         '�H�@���
Private Const pcmbD_SYUBETSU10% = 9         '�I�@���
Private Const pcmbD_SYUBETSU11% = 10        '�J�@���
Private Const pcmbD_SYUBETSU12% = 11        '�K�@���
Private Const pcmbD_SYUBETSU13% = 12        '�L�@���
Private Const pcmbD_SYUBETSU14% = 13        '�M�@���
Private Const pcmbD_SYUBETSU15% = 14        '�N�@���
Private Const pcmbD_SYUBETSU16% = 15        '�O�@���
Private Const pcmbD_SYUBETSU17% = 16        '�P�@���
Private Const pcmbD_SYUBETSU18% = 17        '�Q�@���
Private Const pcmbD_SYUBETSU19% = 18        '�R�@���
Private Const pcmbD_SYUBETSU20% = 19        '�S�@���
Private Const pcmbD_SYUBETSU21% = 20        '21�@���
Private Const pcmbD_SYUBETSU22% = 21        '22�@���
Private Const pcmbD_SYUBETSU23% = 22        '23�@���
Private Const pcmbD_SYUBETSU24% = 23        '24�@���
Private Const pcmbD_SYUBETSU25% = 24        '25�@���

'�R�}���h�{�^���ŗL����
Private Const cmd1_25% = 5                  '1-25�s�\��
Private Const cmd26_50% = 6                 '26-50�s�\��

'�O��ʏ��
Private Const pcmbSHIMUKE% = 0              '�O��ʎd������
Private Const ptxSHIJI_QTY% = 8             '����




Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    
    
        
    Select Case KeyCode
        Case vbKeyReturn
        
            Select Case Index
                                        '�����^�\���@���
                Case pcmbD_SYUBETSU01 To pcmbD_SYUBETSU25
            
                    D_Item_Tbl(Doukon_Start - 1 + Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)
            
            End Select
            Call Tab_Ctrl(Shift)        '�ړ�
        
        
                
        
        
        
        Case vbKeyInsert                '�s�}��
        
            Call Gyo_Ins_Proc(Index)
    
            If Item_Disp_Proc() Then
                G_SCREEN_FLG = SYS_ERR
                PI000102.Visible = False
            End If
        
        
        
        Case vbKeyDelete                '�s�ǉ�
        
                        
            Call Gyo_Del_Proc(Index)
    

            If Item_Disp_Proc() Then
                G_SCREEN_FLG = SYS_ERR
                PI000102.Visible = False
            End If
    
    
    End Select

End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    
    Select Case Index
                                '�����^�\���@���
        Case pcmbD_SYUBETSU01 To pcmbD_SYUBETSU25
    
            D_Item_Tbl(Doukon_Start - 1 + Index - pcmbD_SYUBETSU01).SYUBETSU = Right(Combo1(Index).text, 2)
    
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer


    For i = ptxD_HIN_GAI01 To ptxD_BIKOU25
    
        Select Case i
            Case ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, _
                    ptxD_HIN_GAI06, ptxD_HIN_GAI07, ptxD_HIN_GAI08, ptxD_HIN_GAI09, ptxD_HIN_GAI10, _
                    ptxD_HIN_GAI11, ptxD_HIN_GAI12, ptxD_HIN_GAI13, ptxD_HIN_GAI14, ptxD_HIN_GAI15, _
                    ptxD_HIN_GAI16, ptxD_HIN_GAI17, ptxD_HIN_GAI18, ptxD_HIN_GAI19, ptxD_HIN_GAI20, _
                    ptxD_HIN_GAI21, ptxD_HIN_GAI22, ptxD_HIN_GAI23, ptxD_HIN_GAI24, ptxD_HIN_GAI25
    
                Text1(i).text = RTrim(StrConv(Text1(i), vbUpperCase))

        End Select
    
    
    
    
    Next i



    Select Case Index
        Case P_CMD_Upd                      '�X�V
        Case P_CMD_DEL                      '�폜
        Case P_CMD_DSP                      '����/�\��
        Case cmd1_25                        '1-25�s�\��
            
            
            
            ans = MsgBox("��ʓ��e��ۑ����܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
            
                For i = ptxD_HIN_GAI01 To ptxD_BIKOU25
                
                
                
                
                    If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                        Exit Sub
                    End If
                
                Next i
            End If
            
            Doukon_Start = 1
            If Item_Disp_Proc() Then
                G_SCREEN_FLG = SYS_ERR
                PI000102.Visible = False
            End If
        Case cmd26_50                       '26-50�s�\��
            
            ans = MsgBox("��ʓ��e��ۑ����܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
            
                For i = ptxD_HIN_GAI01 To ptxD_BIKOU25
                
                
    
                
                
                
                
                    If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                        Exit Sub
                    End If
                
                Next i
            End If
            
            
            Doukon_Start = 26
            If Item_Disp_Proc() Then
                G_SCREEN_FLG = SYS_ERR
                PI000102.Visible = False
            End If
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        
        
        Case P_CMD_End                      '�I��
            
            ans = MsgBox("��ʓ��e��ۑ����܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
            
                For i = ptxD_HIN_GAI01 To ptxD_BIKOU25
                
                    
                    
                    
                    If Error_Check_Proc(i) Then     '�G���[�`�F�b�N
                        Exit Sub
                    End If
                
                Next i
            End If
            
            
            Init_Flg = True
            PI000102.Visible = False
    
    
    End Select

End Sub

Private Sub Form_Activate()

    If Not Init_Flg Then
        Exit Sub
    End If
            
    Init_Flg = False


    If Item_Disp_Proc() Then
        G_SCREEN_FLG = SYS_ERR
        PI000102.Visible = False
    End If


End Sub

Private Sub Form_Load()
Dim i   As Integer

    Init_Flg = True

    '��ʂ̃Z�b�g
    For i = pcmbD_SYUBETSU01 To pcmbD_SYUBETSU25
        If Code_Set_Proc(i, P_KBN06_CD, 1) Then
            Unload Me
        End If
    Next i
    

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


Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts         As Integer

Dim SEQNO       As Integer
Dim Gyo         As Integer
Dim text        As Integer
Dim i           As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
    
    Item_Disp_Proc = True
    
    SEQNO = Doukon_Start
    Gyo = 0
    text = 0

    Do
        If SEQNO > Doukon_Start + 24 Then
            Exit Do
        End If
        
        '�s�ԍ�
        
        lblNumber(Gyo).Caption = Format(SEQNO, "00")
        
        '���
        Combo1(Gyo).ListIndex = -1
        For i = 0 To Combo1(Gyo).ListCount - 1
        
            If D_Item_Tbl(SEQNO - 1).SYUBETSU = Right(Combo1(Gyo).List(i), 2) Then
                Combo1(Gyo).ListIndex = i
                Exit For
            End If
        
        Next i
        '�i��
        Text1(text).text = Trim(D_Item_Tbl(SEQNO - 1).HIN_GAI)
        If Text1(text).text <> "" Then
            '�i��
            Call UniCode_Conv(K0_ITEM.JGYOBU, D_Item_Tbl(SEQNO - 1).JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, D_Item_Tbl(SEQNO - 1).NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, D_Item_Tbl(SEQNO - 1).HIN_GAI)
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Text1(text + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    
                    
                    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                        Text1(text + 4).text = ""
                    Else
                        Text1(text + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)
                    End If
                
                
                    If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                            StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                            StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                        Exit Function
                    
                    End If
        
                    Text1(text + 5).text = Format(Sumi_Qty + Mi_Qty, "#0")
                
                
                
                Case BtErrKeyNotFound
                    Text1(text + 1).text = "���o�^�i��"
                    Text1(text + 4).text = ""
                    Text1(text + 5).text = ""
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            
            End Select
                
            '����
            Text1(text + 2).text = Format(D_Item_Tbl(SEQNO - 1).QTY, "#0.00")
            '����
            Text1(text + 3).text = Format(D_Item_Tbl(SEQNO - 1).SHIJI_QTY, "#0.00")
            '���l
            Text1(text + 6).text = D_Item_Tbl(SEQNO - 1).BIKOU
            
        Else
            Text1(text + 1).text = ""
            Text1(text + 2).text = ""
            Text1(text + 3).text = ""
            Text1(text + 4).text = ""
            Text1(text + 5).text = ""
            Text1(text + 6).text = ""
        End If
            
            
        SEQNO = SEQNO + 1
        Gyo = Gyo + 1
        text = text + 7
    Loop

    Item_Disp_Proc = False

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


Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
    
    
Dim Mi_Qty      As Long
Dim Sumi_Qty    As Long
    
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
                                '�����^�\���@�i��
        Case ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, ptxD_HIN_GAI06, ptxD_HIN_GAI07, ptxD_HIN_GAI08, ptxD_HIN_GAI09, ptxD_HIN_GAI10, _
                ptxD_HIN_GAI11, ptxD_HIN_GAI12, ptxD_HIN_GAI13, ptxD_HIN_GAI14, ptxD_HIN_GAI15, ptxD_HIN_GAI16, ptxD_HIN_GAI17, ptxD_HIN_GAI18, ptxD_HIN_GAI19, ptxD_HIN_GAI20, _
                ptxD_HIN_GAI21, ptxD_HIN_GAI22, ptxD_HIN_GAI23, ptxD_HIN_GAI24, ptxD_HIN_GAI25
            
            If Trim(Text1(Mode).text) = "" Then
                Text1(Mode + 1).text = ""
                Text1(Mode + 2).text = ""
                Text1(Mode + 3).text = ""
                Text1(Mode + 4).text = ""
                Text1(Mode + 5).text = ""
                Text1(Mode + 6).text = ""
            
            
                i = 0
                j = Mode - ptxD_HIN_GAI01 + ((Doukon_Start - 1) * 7)
                Do
                    j = j - 7
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                D_Item_Tbl(i).JGYOBU = ""
                D_Item_Tbl(i).NAIGAI = ""
                D_Item_Tbl(i).HIN_GAI = ""
                D_Item_Tbl(i).QTY = 0
                D_Item_Tbl(i).SHIJI_QTY = 0
                D_Item_Tbl(i).BIKOU = ""
            
            
            Else
                Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(PI000101.Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(PI000101.Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                        
                        
                        
                        
                        '�i�ԁi���j�œǂݑւ�
                        Call UniCode_Conv(K2_ITEM.JGYOBU, Mid(Right(PI000101.Combo1(pcmbSHIMUKE).text, 4), 3, 1))
                        Call UniCode_Conv(K2_ITEM.NAIGAI, Mid(Right(PI000101.Combo1(pcmbSHIMUKE).text, 4), 4, 1))
                        Call UniCode_Conv(K2_ITEM.HIN_NAI, Text1(Mode).text)
                        
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K2_ITEM, Len(K2_ITEM), 2)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                        
                        
                        
                        
                        
                        
                        
                                '���ޕi�œǂݑւ�
                                                
                                Call UniCode_Conv(K0_ITEM.JGYOBU, SHIZAI)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, NAIGAI_NAI)
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(Mode).text)
                                
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                        
                                        If HIN_INV Then
                                            '���o�^�i�ԁ@�@���ނƂ��Ă���
                                            Call UniCode_Conv(ITEMREC.JGYOBU, SHIZAI)
                                            Call UniCode_Conv(ITEMREC.NAIGAI, NAIGAI_NAI)
                                            Call UniCode_Conv(ITEMREC.HIN_GAI, Text1(Mode).text)
                                            Call UniCode_Conv(ITEMREC.HIN_NAME, "���o�^�i��")
                                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                        
                                        Else
                                        
                                            MsgBox "���͂������ڂ̓G���[�ł��B"
                                            Text1(Mode).SetFocus
                                            Exit Function
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                        Exit Function
                                
                                End Select
                
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                                Exit Function
                       End Select
                        
                
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                        Exit Function
                
                End Select
    
                '�i��
                Text1(Mode + 1).text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                '�W���I��
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Else
                    Text1(Mode + 4).text = StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
                '�݌ɐ�
                If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                        StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                    Exit Function
                
                End If
            
                Text1(Mode + 5).text = Format(Sumi_Qty + Mi_Qty, "#0")
            
            
            
                i = 0
                j = Mode - ptxD_HIN_GAI01 + ((Doukon_Start - 1) * 7)
                Do
                    j = j - 7
                    If j < 0 Then
                        Exit Do
                    End If
                    i = i + 1
                Loop
            
                D_Item_Tbl(i).JGYOBU = StrConv(ITEMREC.JGYOBU, vbUnicode)
                D_Item_Tbl(i).NAIGAI = StrConv(ITEMREC.NAIGAI, vbUnicode)
                D_Item_Tbl(i).HIN_GAI = StrConv(ITEMREC.HIN_GAI, vbUnicode)
            
            
            End If
                                '�����^�\���@����
        Case ptxD_QTY01, ptxD_QTY02, ptxD_QTY03, ptxD_QTY04, ptxD_QTY05, ptxD_QTY06, ptxD_QTY07, ptxD_QTY08, ptxD_QTY09, ptxD_QTY10, _
                ptxD_QTY11, ptxD_QTY12, ptxD_QTY13, ptxD_QTY14, ptxD_QTY15, ptxD_QTY16, ptxD_QTY17, ptxD_QTY18, ptxD_QTY19, ptxD_QTY20, _
                ptxD_QTY21, ptxD_QTY22, ptxD_QTY23, ptxD_QTY24, ptxD_QTY25
            
            If Trim(Text1(Mode).text) = "" Then
                If Trim(Text1(Mode - 2).text) <> "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            Else
                If Trim(Text1(Mode - 2).text) = "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsNumeric(Text1(Mode).text) Then
                        MsgBox "���͂������ڂ̓G���[�ł��B"
                        Text1(Mode).SetFocus
                        Exit Function
                    Else
                        
                        Text1(Mode).text = Format(CDbl(Text1(Mode).text), "#0.00")
                        
                        
                        i = 0
                        j = Mode - ptxD_HIN_GAI01 + ((Doukon_Start - 1) * 7)
                        Do
                            j = j - 7
                            If j < 0 Then
                                Exit Do
                            End If
                            i = i + 1
                        Loop
                    
                        D_Item_Tbl(i).QTY = CDbl(Text1(Mode).text)
                        
                        
                        '����
                        If IsNumeric(PI000101.Text1(ptxSHIJI_QTY).text) Then
                            Text1(Mode + 1).text = Format(CDbl(CLng(PI000101.Text1(ptxSHIJI_QTY).text) * CDbl(Text1(Mode).text)), "#0.00")
                            D_Item_Tbl(i).SHIJI_QTY = CDbl(Text1(Mode + 1).text)
                        Else
                            Text1(Mode + 1).text = ""
                            D_Item_Tbl(i).SHIJI_QTY = 0
                        End If
                    
                    End If
                End If
            End If
                                '�����^�\���@���l
        Case ptxD_BIKOU01, ptxD_BIKOU02, ptxD_BIKOU03, ptxD_BIKOU04, ptxD_BIKOU05, ptxD_BIKOU06, ptxD_BIKOU07, ptxD_BIKOU08, ptxD_BIKOU09, ptxD_BIKOU10, _
                ptxD_BIKOU11, ptxD_BIKOU12, ptxD_BIKOU13, ptxD_BIKOU14, ptxD_BIKOU15, ptxD_BIKOU16, ptxD_BIKOU17, ptxD_BIKOU18, ptxD_BIKOU19, ptxD_BIKOU20, _
                ptxD_BIKOU21, ptxD_BIKOU22, ptxD_BIKOU23, ptxD_BIKOU24, ptxD_BIKOU25
            If Trim(Text1(Mode).text) <> "" Then
                If Trim(Text1(Mode - 6).text) = "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
            i = 0
            j = Mode - ptxD_HIN_GAI01 + ((Doukon_Start - 1) * 7)
            Do
                j = j - 7
                If j < 0 Then
                    Exit Do
                End If
                i = i + 1
            Loop
    
            D_Item_Tbl(i).BIKOU = Text1(Mode).text
            
    End Select
        
        
    Error_Check_Proc = False
    

End Function


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    Select Case Index
        Case ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, _
                ptxD_HIN_GAI06, ptxD_HIN_GAI07, ptxD_HIN_GAI08, ptxD_HIN_GAI09, ptxD_HIN_GAI10, _
                ptxD_HIN_GAI11, ptxD_HIN_GAI12, ptxD_HIN_GAI13, ptxD_HIN_GAI14, ptxD_HIN_GAI15, _
                ptxD_HIN_GAI16, ptxD_HIN_GAI17, ptxD_HIN_GAI18, ptxD_HIN_GAI19, ptxD_HIN_GAI20, _
                ptxD_HIN_GAI21, ptxD_HIN_GAI22, ptxD_HIN_GAI23, ptxD_HIN_GAI24, ptxD_HIN_GAI25

                Text1(Index).text = RTrim(StrConv(Text1(Index), vbUpperCase))
    
    End Select
        
        
        
        
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Gyo_Ins_Proc(Index As Integer)
'----------------------------------------------------------------------------
'                   �s�}������
'----------------------------------------------------------------------------
Dim i   As Integer

    If Index + (Doukon_Start - 1) >= UBound(D_Item_Tbl) Then
        Exit Sub
    End If


    For i = UBound(D_Item_Tbl) - 1 To Index + (Doukon_Start - 1) Step -1
    
    
                                                    
        D_Item_Tbl(i + 1).SYUBETSU = D_Item_Tbl(i).SYUBETSU         '���
        D_Item_Tbl(i + 1).JGYOBU = D_Item_Tbl(i).JGYOBU             '���ƕ�
        D_Item_Tbl(i + 1).NAIGAI = D_Item_Tbl(i).NAIGAI             '�����O
        D_Item_Tbl(i + 1).HIN_GAI = D_Item_Tbl(i).HIN_GAI           '�i��
        D_Item_Tbl(i + 1).QTY = D_Item_Tbl(i).QTY                   '����
        D_Item_Tbl(i + 1).SHIJI_QTY = D_Item_Tbl(i).SHIJI_QTY       '���ʁi�w�����j
        D_Item_Tbl(i + 1).BIKOU = D_Item_Tbl(i).BIKOU               '���l�i���͒l�j
        D_Item_Tbl(i + 1).ID_NO = D_Item_Tbl(i).ID_NO               'ID_No(�o�ח\��ID_No)
    
    
    
    Next i


    Combo1(Index).ListIndex = -1

    For i = Index * 7 To (Index * 7) + 6
        Text1(Index).text = ""
    Next i


    D_Item_Tbl(Index + (Doukon_Start - 1)).SYUBETSU = ""        '���
    D_Item_Tbl(Index + (Doukon_Start - 1)).JGYOBU = ""          '���ƕ�
    D_Item_Tbl(Index + (Doukon_Start - 1)).NAIGAI = ""          '�����O
    D_Item_Tbl(Index + (Doukon_Start - 1)).HIN_GAI = ""         '�i��
    D_Item_Tbl(Index + (Doukon_Start - 1)).QTY = 0              '����
    D_Item_Tbl(Index + (Doukon_Start - 1)).SHIJI_QTY = 0        '���ʁi�w�����j
    D_Item_Tbl(Index + (Doukon_Start - 1)).BIKOU = ""           '���l�i���͒l�j
    D_Item_Tbl(Index + (Doukon_Start - 1)).ID_NO = ""           'ID_No(�o�ח\��ID_No)


End Sub

Private Sub Gyo_Del_Proc(Index As Integer)
'----------------------------------------------------------------------------
'                   �s�폜����
'----------------------------------------------------------------------------
Dim i   As Integer

    For i = Index + (Doukon_Start - 1) To UBound(D_Item_Tbl) - 1
    
    
                                                    
        D_Item_Tbl(i).SYUBETSU = D_Item_Tbl(i + 1).SYUBETSU     '���
        D_Item_Tbl(i).JGYOBU = D_Item_Tbl(i + 1).JGYOBU         '���ƕ�
        D_Item_Tbl(i).NAIGAI = D_Item_Tbl(i + 1).NAIGAI         '�����O
        D_Item_Tbl(i).HIN_GAI = D_Item_Tbl(i + 1).HIN_GAI       '�i��
        D_Item_Tbl(i).QTY = D_Item_Tbl(i + 1).QTY               '����
        D_Item_Tbl(i).SHIJI_QTY = D_Item_Tbl(i + 1).SHIJI_QTY   '���ʁi�w�����j
        D_Item_Tbl(i).BIKOU = D_Item_Tbl(i + 1).BIKOU           '���l�i���͒l�j
        D_Item_Tbl(i).ID_NO = D_Item_Tbl(i + 1).ID_NO           'ID_No(�o�ח\��ID_No)
    
    
    
    Next i


End Sub

Private Sub Text1_LostFocus(Index As Integer)

    

    Select Case Index
        Case ptxD_HIN_GAI01, ptxD_HIN_GAI02, ptxD_HIN_GAI03, ptxD_HIN_GAI04, ptxD_HIN_GAI05, _
                ptxD_HIN_GAI06, ptxD_HIN_GAI07, ptxD_HIN_GAI08, ptxD_HIN_GAI09, ptxD_HIN_GAI10, _
                ptxD_HIN_GAI11, ptxD_HIN_GAI12, ptxD_HIN_GAI13, ptxD_HIN_GAI14, ptxD_HIN_GAI15, _
                ptxD_HIN_GAI16, ptxD_HIN_GAI17, ptxD_HIN_GAI18, ptxD_HIN_GAI19, ptxD_HIN_GAI20, _
                ptxD_HIN_GAI21, ptxD_HIN_GAI22, ptxD_HIN_GAI23, ptxD_HIN_GAI24, ptxD_HIN_GAI25

                Text1(Index).text = RTrim(StrConv(Text1(Index), vbUpperCase))
    
    End Select

End Sub
