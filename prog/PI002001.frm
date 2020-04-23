VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PI002001 
   Caption         =   "構成マスター品番一括変更"
   ClientHeight    =   13830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20190
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   14.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13830
   ScaleWidth      =   20190
   StartUpPosition =   2  '画面の中央
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10140
      Left            =   240
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   63
      Top             =   1320
      Width           =   4215
   End
   Begin VB.TextBox txtTANTO_NAME 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   62
      Top             =   240
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   4
      Left            =   4680
      TabIndex        =   52
      Top             =   10920
      Width           =   13335
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   240
         Width           =   3375
      End
      Begin RichTextLib.RichTextBox txtBEF_BIKOU 
         Height          =   1455
         Index           =   4
         Left            =   1200
         TabIndex        =   55
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"PI002001.frx":0000
      End
      Begin RichTextLib.RichTextBox txtAFT_BIKOU 
         Height          =   1455
         Index           =   4
         Left            =   7680
         TabIndex        =   56
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"PI002001.frx":00DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   6840
         TabIndex        =   57
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   360
         TabIndex        =   58
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "品　番"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更前"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   360
         TabIndex        =   60
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更後"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   6840
         TabIndex        =   59
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   3
      Left            =   4680
      TabIndex        =   42
      Top             =   8520
      Width           =   13335
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   240
         Width           =   2415
      End
      Begin RichTextLib.RichTextBox txtBEF_BIKOU 
         Height          =   1455
         Index           =   3
         Left            =   1200
         TabIndex        =   45
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"PI002001.frx":01BC
      End
      Begin RichTextLib.RichTextBox txtAFT_BIKOU 
         Height          =   1455
         Index           =   3
         Left            =   7680
         TabIndex        =   46
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"PI002001.frx":029A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   6840
         TabIndex        =   47
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   360
         TabIndex        =   48
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更後"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   6840
         TabIndex        =   51
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更前"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   360
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "品　番"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   2
      Left            =   4680
      TabIndex        =   32
      Top             =   6000
      Width           =   13335
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   3375
      End
      Begin RichTextLib.RichTextBox txtBEF_BIKOU 
         Height          =   1455
         Index           =   2
         Left            =   1200
         TabIndex        =   35
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"PI002001.frx":0378
      End
      Begin RichTextLib.RichTextBox txtAFT_BIKOU 
         Height          =   1455
         Index           =   2
         Left            =   7680
         TabIndex        =   36
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"PI002001.frx":0456
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   6840
         TabIndex        =   37
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   360
         TabIndex        =   39
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "品　番"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更前"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   360
         TabIndex        =   40
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更後"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   6840
         TabIndex        =   38
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   1
      Left            =   4680
      TabIndex        =   22
      Top             =   3600
      Width           =   13335
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   3375
      End
      Begin RichTextLib.RichTextBox txtBEF_BIKOU 
         Height          =   1455
         Index           =   1
         Left            =   1200
         TabIndex        =   25
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"PI002001.frx":0534
      End
      Begin RichTextLib.RichTextBox txtAFT_BIKOU 
         Height          =   1455
         Index           =   1
         Left            =   7680
         TabIndex        =   26
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"PI002001.frx":0612
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   6840
         TabIndex        =   27
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更後"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6840
         TabIndex        =   28
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   29
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "品　番"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
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
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更前"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "次　頁"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   11880
      TabIndex        =   21
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "前　頁"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   10200
      TabIndex        =   20
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtKEN_SU 
      Alignment       =   1  '右揃え
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   11640
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   0
      Left            =   4680
      TabIndex        =   8
      Top             =   1200
      Width           =   13335
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
      Begin RichTextLib.RichTextBox txtBEF_BIKOU 
         Height          =   1455
         Index           =   0
         Left            =   1200
         TabIndex        =   14
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"PI002001.frx":06F0
      End
      Begin RichTextLib.RichTextBox txtAFT_BIKOU 
         Height          =   1455
         Index           =   0
         Left            =   7680
         TabIndex        =   17
         Top             =   840
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"PI002001.frx":07CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   6840
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更後"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6840
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "　備考"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "変更前"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "品　番"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
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
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   4
      Top             =   600
      Width           =   2100
   End
   Begin VB.TextBox txtTANTO_CODE 
      Appearance      =   0  'ﾌﾗｯﾄ
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   3
      Top             =   240
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "クリア"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "終　　了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   13560
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "更　新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "status"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   66
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label lblEXCEL_FILE 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4680
      TabIndex        =   65
      Top             =   840
      Width           =   120
   End
   Begin VB.Label lblPage_Su 
      Alignment       =   1  '右揃え
      Height          =   255
      Left            =   17040
      TabIndex        =   64
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "件　数"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   12
      Top             =   11640
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "仕向先"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   110
      Left            =   360
      TabIndex        =   7
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "担当者"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   111
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "親品番"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   119
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   960
   End
End
Attribute VB_Name = "PI002001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'テキスト用添字




'コンボ用添字
Private Const pcmbSHIMUKE% = 0              '仕向け先

'チェック用添字

'ｵﾌﾟｼｮﾝﾎﾞﾀﾝ用添字


'リッチテキスト用添字



'コマンドボタン固有操作
Private Const cmdClear% = 0                 'クリアー
Private Const cmdUpdate% = 1                '更新
Private Const cmdBEFO_PAGE% = 2             '前ページ
Private Const cmdEND% = 3                   '終了
Private Const cmdNEXT_PAGE% = 4             '次ページ


'EXCEL 列
Private Const exeHIN_GAI% = 1


Dim SHIMUKE_CODE    As String * 2


Private Type Item_tbl_tag
    Err_Mark    As String * 1
    Item_code   As String * 20
    BEF_BIKOU   As String * 120
    AFT_BIKOU   As String * 120
End Type

Private Item_tbl(0 To 99) As Item_tbl_tag


Dim P_Cnt       As Integer
Dim Max_P_cnt   As Integer

'Private Const Last_Update_day$ = "構成マスター備考一括変更 (PI00200 2019.04.19 11:15)"
'Private Const Last_Update_day$ = "構成マスター備考一括変更 (PI00200 2019.09.12 19:25)"
Private Const Last_Update_day$ = "構成マスター備考一括変更 (PI00200 2019.11.25 17:30) 備考欄余白時のエラー修正"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    PI002001.MousePointer = vbHourglass

    Call Ctrl_Lock(PI002001)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI002001)


    PI002001.MousePointer = vbDefault

End Sub


Private Function Item_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   画面表示
'----------------------------------------------------------------------------
Dim i       As Integer

Dim j       As Integer


Dim k       As Integer

Dim sts     As Integer
    
    
    Item_Disp_Proc = True
    
    j = (P_Cnt - 1) * 5


    k = -1


    For i = j To j + 4


        k = k + 1
        If Trim(Item_tbl(i).Item_code) = "" Then
            txtAFT_BIKOU(k).Locked = True
            txtAFT_BIKOU(k).BackColor = &H8000000F
            
            '2019.06.12 下記４行追加（クリアする事！"
            txtBEF_BIKOU(k).Text = ""
            txtAFT_BIKOU(k).Text = ""
            Text1(k).Text = ""
            Text3(k).Text = ""
            
        Else
        
            txtAFT_BIKOU(k).Locked = False
            txtAFT_BIKOU(k).BackColor = &H80000005
        
        
            Text1(k).Text = Trim(Item_tbl(i).Item_code)
            txtBEF_BIKOU(k).Text = Trim(Item_tbl(i).BEF_BIKOU)
            
            
            txtAFT_BIKOU(k).Text = Trim(Item_tbl(i).AFT_BIKOU)
        
        
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
            Call UniCode_Conv(K0_ITEM.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Item_tbl(i).Item_code)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Text3(k).Text = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text3(k).Text = ""
                Case Else
                
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Function
                
            End Select
        
        End If
    Next i



    lblPage_Su.Caption = P_Cnt & "/" & Max_P_cnt




    Item_Disp_Proc = False


End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   構成マスタ＆商品化指示ﾃﾞｰﾀ出力
'----------------------------------------------------------------------------
Dim sts         As Integer

Dim i           As Integer

    Update_Proc = True

    Call Input_Lock

                                        
                                        
                                        'トランザクション開始
    sts = BTRV(BtOpBeginConcurrentTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
        Exit Function
    End If


    For i = 0 To 99
    
        If Trim(Item_tbl(i).Item_code) = "" Then
            Exit For
        End If
    
    
    
    
        Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
        Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
        Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
        Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Item_tbl(i).Item_code)
            
        Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
        Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
            
        sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                Call UniCode_Conv(P_COMPO_O_REC.BIKOU, Item_tbl(i).AFT_BIKOU)
                sts = BTRV(BtOpUpdate, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
            
                If sts Then
                    Call File_Error(sts, BtOpUpdate, "構成マスタ")
                    GoTo Abort_Tran
                End If
            
            Case BtErrKeyNotFound
            Case Else
                    
                Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                GoTo Abort_Tran
                
                
        End Select
    
    
    
    
    
    
    Next i
    
















                                        'トランザクション終了
    sts = BTRV(BtOpEndTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If

    Call Input_UnLock

    Update_Proc = False
    

    Exit Function

Abort_Tran:

    sts = BTRV(BtOpAbortTransaction, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock

End Function



Private Sub Combo1_GotFocus(Index As Integer)

Dim sts         As Integer

Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long



    Select Case Index
        Case pcmbSHIMUKE        '仕向け先



            SHIMUKE_CODE = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2)   '2013.08.29









    End Select
End Sub






Private Sub Command2_Click(Index As Integer)

Dim sts     As Integer
Dim yn      As Integer

    Select Case Index
        Case cmdClear                       'クリアー

            Init_Proc

        Case cmdUpdate                      '更新
        
        
            If Err_Check_Proc Then
                Exit Sub
            End If
        
        
        
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
        
        
        
        
        Case cmdBEFO_PAGE                   '前ページ
        
            
            If Err_Check_Proc Then
                Exit Sub
            End If
            
            
            
            If P_Cnt = 1 Then
                MsgBox "先頭頁を表示しています。"
                Exit Sub
            End If
            
        
            P_Cnt = P_Cnt - 1
        
        
            sts = Item_Disp_Proc()
        
        
        Case cmdNEXT_PAGE                   '次ページ

            If Err_Check_Proc Then
                Exit Sub
            End If


            If P_Cnt = Max_P_cnt Then
                MsgBox "最終頁を表示しています。"
                Exit Sub
            End If
            
        
            P_Cnt = P_Cnt + 1
        
        
            sts = Item_Disp_Proc()



        Case cmdEND                         '終了

            Unload Me


    End Select
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer

Dim MUKE_CODE   As Variant



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "構成マスター備考一括変更　「起動処理中」", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    Me.Enabled = False
                                
                                
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止します。"
        End
    End If
    LOG_F = RTrim(c)



Show    '2015.03.26


    PI002001.Caption = Last_Update_day      '2016.02.10


    If File_Open_Proc() Then
        Unload Me
    End If



    'ｺｰﾄﾞﾏｽﾀ定義
    Call P_CODE_TBL_Proc


    '仕向け先のセット
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If


    Combo1(pcmbSHIMUKE).ListIndex = 0



    '画面初期設定
    Call Init_Proc


    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "構成マスター備考一括変更　「準備完了」", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    Me.Enabled = True
    txtTANTO_CODE.SetFocus
    

End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer





    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "", 0)
    End If
    Set PI002001 = Nothing

    End
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sts As Integer
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, txtTANTO_CODE.Text)

    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    
    Select Case sts
        Case BtNoErr
            txtTANTO_NAME.Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        Case BtErrKeyNotFound
            txtTANTO_NAME.Text = ""

            MsgBox "入力した項目はエラーです。(担当者)"
            txtTANTO_CODE.SetFocus
            Exit Sub
        
        Case Else
            
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Unload Me

    End Select
    
    
    
    
    
    
    
    
    
    lblEXCEL_FILE.Caption = Trim(Data.Files(1))


    If List_Disp_Proc() Then
        Unload Me
    End If

End Sub



Private Sub Init_Proc()
'----------------------------------------------------------------------------
'                   入力画面の初期設定
'----------------------------------------------------------------------------
Dim i           As Integer
Dim sts         As Integer


    
    List1.Clear

    For i = 0 To 4
    
        Text1(i).Text = ""
        Text3(i).Text = ""
        
    
    Next i
    
    
    For i = 0 To 4
    
        txtBEF_BIKOU(i).Text = ""
        txtAFT_BIKOU(i).Text = ""
    
    Next i
    
    

End Sub









Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   コードマスタをコンボにセットする。
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer

Start_Proc0:        '2015.03.26ok

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
                
                
                Call File_Error(sts, com, "コードマスタ")
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

            MsgBox "入力した項目はエラーです。(担当者)"
            txtTANTO_CODE.SetFocus
        Case Else
            
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Unload Me

    End Select




    Call Tab_Ctrl(Shift)        '移動

End Sub

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「EXCEL」読込み処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long



Dim wkHIN_GAI       As String * 20

Dim i               As Long
Dim j               As Long
Dim k               As Long

Dim TEXT_BEF        As String


Dim Err_Mark        As String * 4


    List_Disp_Proc = True

    Call Input_Lock


hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "取込みデータ　表示処理開始！！", Me.hwnd, 0)





    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    
    
    On Error GoTo Error_Proc

    xlApp.Workbooks.Open (lblEXCEL_FILE.Caption), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0
    
    
    Row = 0
    
    
    List1.Clear
    
    
    For i = 0 To 99
    
        Item_tbl(i).Err_Mark = ""
    
        Item_tbl(i).Item_code = ""
        Item_tbl(i).BEF_BIKOU = ""
        Item_tbl(i).AFT_BIKOU = ""
        
    
    
    
    
    Next i
    
    k = -1
    
    For j = 1 To xlApp.Worksheets.Count
    
        Set xlSheet = xlApp.Worksheets(j)
        xlSheet.Activate
    
    
        For i = 1 To 1048576
            
            
            If Trim(xlSheet.Application.Cells(i, exeHIN_GAI)) = "" Then
                Exit For
            End If
            
            
            If i > 1 Then
            
            
            
            
                Row = Row + 1
            
                If Row > 99 Then
                    MsgBox "最大作業件数は１００件です。データの分割を行って下さい。"
                    Exit For
                End If
            
            
            
            '品番
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, exeHIN_GAI))
        
            
            
            
            
            
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2))
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1))
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1))
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, wkHIN_GAI)
            
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
            
                sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        
                        Err_Mark = ""
                    
                        
                        k = k + 1
                        Item_tbl(k).Item_code = wkHIN_GAI
                        
                        Item_tbl(k).BEF_BIKOU = StrConv(P_COMPO_O_REC.BIKOU, vbUnicode)
                        Item_tbl(k).AFT_BIKOU = StrConv(P_COMPO_O_REC.BIKOU, vbUnicode)
                    
                    
                    
                    Case BtErrKeyNotFound
                    
                        Err_Mark = "Err "
                            
            
                    Case Else
                            
                        Call File_Error(sts, BtOpGetEqual, "構成マスタ")
                        Exit For
                
                
                End Select
            
            
                List1.AddItem wkHIN_GAI & Err_Mark
            
            
            
            
            
                txtKEN_SU.Text = Format(Row, "#0") & "件"
                DoEvents
            
            
        
        
            End If
        Next i
    
    
    
    
    Next j


    


    Max_P_cnt = CInt(ToRoundUp(CCur(Row) / 5, 0))




    xlApp.DisplayAlerts = False

    xlBook.Close False
    xlApp.Quit 'EXCELを閉じる
    Set xlApp = Nothing



    P_Cnt = 1


    sts = Item_Disp_Proc()
        
    




hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "取込みデータ　表示処理終了！！", Me.hwnd, 0)



    Call Input_UnLock


    List_Disp_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 ファイル名または番号が不正です。
        '53 ファイルが見つかりません。
        '54 ファイル モードが不正です。
        '55 ファイルは既に開かれています。
        '57 デバイス I/O エラーです。
        '59 レコード長が一致しません。
        '61 ディスクの空き容量が不足しています。
        '62 ファイルにこれ以上データがありません。
        '63 レコード番号が不正です。
        '68 デバイスが準備されていません。
        '70 書き込みできません。
        '71 ディスクが準備されていません。
        '75 パス名が無効です。
        '76 パスが見つかりません。
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            
            List_Disp_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "読込み対象のEXCELデータに異常が有ります。内容を確認後、再実行してください。"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCELを閉じる
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_Proc = False      '



        Case Else
    End Select
    Call Input_UnLock

End Function


' ------------------------------------------------------------------------
'       指定した精度の数値に切り上げします。
'
' @Param    dValue      丸め対象の倍精度浮動小数点数。
' @Param    iDigits     戻り値の有効桁数の精度。
' @Return               iDigits に等しい精度の数値に切り上げられた数値。
'
'       2012.03.25  frm より　移管
'
' ------------------------------------------------------------------------
Public Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function


Private Function Err_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------

Dim GYO_SU      As Long
Dim i           As Integer
Dim j           As Integer

    Err_Check_Proc = True
        
    
    j = (P_Cnt - 1) * 5
    
    
    For i = 0 To 4
        
        
        If LenB(StrConv(RTrim(txtAFT_BIKOU(i).Text), vbFromUnicode)) > 120 Then
            txtAFT_BIKOU(i).SetFocus  '2019/11/25 位置変更
            MsgBox ("備考 桁数オーバー " & LenB(StrConv(RTrim(txtAFT_BIKOU(i).Text), vbFromUnicode)) & "(最大120文字)") '2019/11/25 メッセージ変更
            'MsgBox ("備考が桁数オーバーしています。 (最大120文字) 内容を確認して下さい。")  2019/11/25 コメントアウト
            'Exit Function 2019/11/25 コメントアウト
        End If
        
        
        GYO_SU = SendMessage(txtAFT_BIKOU(i).hwnd, EM_GETLINECOUNT, 0&, 0&)
        If GYO_SU > 5 Then
            MsgBox "備考最大印字行数は５行です。内容を確認して下さい。"
            txtAFT_BIKOU(i).SetFocus
            Exit Function
        End If


        Item_tbl(j + i).AFT_BIKOU = RTrim(txtAFT_BIKOU(i).Text)

    Next i


    Err_Check_Proc = False


End Function
