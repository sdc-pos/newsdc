VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form F1201001 
   Caption         =   "[美的品番管理]ユニット指図票発行処理([F120100] "
   ClientHeight    =   12345
   ClientLeft      =   2025
   ClientTop       =   2325
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12345
   ScaleWidth      =   14910
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   10440
      MaxLength       =   20
      TabIndex        =   16
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   5640
      TabIndex        =   4
      Top             =   2280
      Width           =   8535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   5640
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "検  索"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3840
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'なし
      Height          =   375
      Index           =   1
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   165
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終　了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印 刷"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   8775
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   15478
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "印刷"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "事"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "国"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "対外品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品　　名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "美的品番"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AllowColMove=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1138"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1005"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=529"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=397"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(10)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=582"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=450"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=8196"
      Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3149"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=8192"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=6059"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=5927"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(29)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=13176"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=13044"
      Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=8192"
      Splits(0)._ColumnProps(35)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   0
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFF&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=43,.parent=1,.bgcolor=&HFFFF00&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(25)  =   ":id=43,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(26)  =   ":id=43,.fontname=ＭＳ ゴシック"
      _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
      _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=45,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=47,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=50,.parent=9,.bgcolor=&HC0C0C0&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=51,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=43,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=44"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=45"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=47"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=16,.parent=43,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=44"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=45"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=47"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=20,.parent=43,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=17,.parent=44"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=18,.parent=45"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=19,.parent=47"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=70,.parent=43,.alignment=0,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=44"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=45"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=47"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=24,.parent=43,.locked=-1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=21,.parent=44"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=22,.parent=45"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=23,.parent=47"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=66,.parent=43,.alignment=0,.locked=-1,.bold=0"
      _StyleDefs(59)  =   ":id=66,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(60)  =   ":id=66,.fontname=ＭＳ ゴシック"
      _StyleDefs(61)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=44"
      _StyleDefs(62)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=45"
      _StyleDefs(63)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=47"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
      _StyleDefs(69)  =   "Named:id=35:Footing"
      _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(71)  =   "Named:id=36:Selected"
      _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=37:Caption"
      _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(75)  =   "Named:id=38:HighlightRow"
      _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=39:EvenRow"
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33,.bgcolor=&HFFFFFF&"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      Caption         =   "印刷枚数"
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "美的品番"
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "対外品番"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "事業部"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "担当者"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   9360
      Width           =   180
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "編集確定"
         Index           =   0
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "検索"
         Index           =   2
      End
      Begin VB.Menu SHORI 
         Caption         =   "画面印刷"
         Index           =   3
      End
   End
End
Attribute VB_Name = "F1201001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxTanto_Code% = 0        '担当者ｺｰﾄﾞ
Private Const ptxTanto_Name% = 1        '担当者名称
Private Const ptxHin_Gai% = 2           '品目ｺｰﾄﾞ
Private Const ptxB_Hin_Code% = 3        '美的品番

Private Const ptxPrint_Maisu% = 4       '印刷枚数

Private Const pcmbJGYOBU% = 0           '事業部


Dim B_ITEM      As New XArrayDB

Private Const Min_Row% = 1              '最小行数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 9              '最大列数

Private Const ColSHORI% = 0             '印刷有無
Private Const ColJGYOBU% = 1            '事業部
Private Const ColNAIGAI% = 2            '国内外
Private Const ColHIN_GAI% = 3           '対外品番
Private Const ColHIN_NAME% = 4          '品名
Private Const ColB_HIN_CODE% = 5        '美的品番



Private DEF_NAIGAI  As String * 1       'ﾃﾞﾌｫﾙﾄ国内外





Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
    Command1(2).Value = True
    

End Sub

Private Sub Command1_Click(Index As Integer)

Dim yn  As Integer
Dim i   As Integer


    Select Case Index
    
        Case 0
    
    
            
            
            For i = ptxTanto_Code To ptxTanto_Name
            
                If Error_Check_Proc(i) Then     'エラーチェック
                    Exit Sub
                End If
            
            Next i
    
            If Grid_Error_Check_Proc() Then
                Exit Sub
            End If
    
    
            yn = MsgBox("ユニット指図票発行しますか？", vbYesNo, "確認入力")
    
            If yn = vbYes Then
        
                If Update_Proc() Then
                    Unload Me
                End If
            
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
                DoEvents
                
                MsgBox "ユニット指図票発行処理が終了しました。"
            
            
            End If
        
        
        
        
        Case 1
    
            Unload Me
    
    
        Case 2
            
            
                        
            
            
            
            If List_Disp_Proc() Then
                Unload Me
            End If
    
    End Select





End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer




    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If


    'コモンコントロールを初期化する
'    cc.dwSize = Len(cc)
'    cc.dwICC = ICC_BAR_CLASSES
    
    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[美的品番管理]ユニット指図票発行処理", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
        
                                
    Combo1(pcmbJGYOBU).Clear
    For i = 0 To UBound(JGYOBU_T)
        Combo1(pcmbJGYOBU).AddItem JGYOBU_T(i).NAME & "                 " & JGYOBU_T(i).CODE
    Next i
                                
                                
                                    
                                'デフォルト事業部取り込み
    If GetIni(App.EXEName, "JGYOBU", App.EXEName, c) Then
    Else
    
        For i = 0 To Combo1(pcmbJGYOBU).ListCount - 1
            If Trim(c) = Right(Combo1(pcmbJGYOBU).List(i), 1) Then
                Combo1(pcmbJGYOBU).ListIndex = i
                Exit For
            End If
        Next i
    
    
    End If
                                
                                'デフォルト国内外取り込み
    If GetIni(App.EXEName, "NAIGAI", App.EXEName, c) Then
        DEF_NAIGAI = NAIGAI_NAI
    Else
        DEF_NAIGAI = Trim(c)
    End If
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '美的品番管理ﾃﾞｰﾀＯＰＥＮ
    If B_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                



    Text1(ptxTanto_Code).SetFocus


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
                                            
                                            
    yn = MsgBox("終了しますか？", vbYesNo, "確認入力")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
                                            
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "美的品番管理ﾃﾞｰﾀ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ内容の表示
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "検索処理　開始")
                        
                        'テーブルリセット
    Set B_ITEM = Nothing
    Row = Min_Row - 1
        
    Last_JGYOBU = Right(Combo1(pcmbJGYOBU).Text, 1)
                        '美的品番管理ﾃﾞｰﾀ読み込み開始
    Call UniCode_Conv(K0_B_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_B_ITEM.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_B_ITEM.HIN_GAI, Text1(ptxHin_Gai).Text)
    
    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
        sts = BTRV(com, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If Last_JGYOBU <> StrConv(B_ITEMREC.JGYOBU, vbUnicode) Then
                    Exit Do
                End If
            
                If Trim(Text1(ptxHin_Gai).Text) <> "" Then
                    If Trim(Text1(ptxHin_Gai).Text) <> Mid(StrConv(B_ITEMREC.HIN_GAI, vbUnicode), 1, Len(Text1(ptxHin_Gai).Text)) Then
                        Exit Do
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "美的品番管理ﾃﾞｰﾀ")
                List_Disp_Proc = SYS_ERR
                Exit Function
        End Select
        
        If Trim(Text1(ptxB_Hin_Code).Text) <> "" And _
            Trim(Text1(ptxB_Hin_Code).Text) <> Mid(StrConv(B_ITEMREC.B_HIN_CODE, vbUnicode), 1, Len(Text1(ptxB_Hin_Code).Text)) Then
        Else
            Row = Row + 1
            If Grid_Set_Proc(Row) Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
        
    Loop
                                'DBテーブルリンク
    Set TDBGrid1.Array = B_ITEM
    
    
    TDBGrid1.Bookmark = Null
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
    
    If B_ITEM.Count(1) > 0 Then
        
        TDBGrid1.MoveFirst
    
        TDBGrid1.Bookmark = 1
        TDBGrid1.Col = ColB_HIN_CODE
    
    End If
    
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "検索処理　終了")
    
    
    Call Input_UnLock
    
    
    List_Disp_Proc = False

    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------

    F1201001.MousePointer = vbHourglass

    Call Ctrl_Lock(F1201001)

    TDBGrid1.Enabled = False

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1201001)

    TDBGrid1.Enabled = True

    F1201001.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    B_ITEM.ReDim Min_Row, Row, Min_Col, Max_Col
    
    
    '削除
    B_ITEM(Row, ColSHORI) = False
    '事業部
    B_ITEM(Row, ColJGYOBU) = StrConv(B_ITEMREC.JGYOBU, vbUnicode)
    '国内外
    B_ITEM(Row, ColNAIGAI) = StrConv(B_ITEMREC.NAIGAI, vbUnicode)
    '対外品番
    B_ITEM(Row, ColHIN_GAI) = RTrim(StrConv(B_ITEMREC.HIN_GAI, vbUnicode))
    '品名
    Call UniCode_Conv(K0_ITEM.JGYOBU, B_ITEM(Row, ColJGYOBU))
    Call UniCode_Conv(K0_ITEM.NAIGAI, B_ITEM(Row, ColNAIGAI))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, B_ITEM(Row, ColHIN_GAI))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            B_ITEM(Row, ColHIN_NAME) = RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
        Case BtErrKeyNotFound
            B_ITEM(Row, ColHIN_NAME) = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Function
    End Select
    '美的品番
    B_ITEM(Row, ColB_HIN_CODE) = RTrim(StrConv(B_ITEMREC.B_HIN_CODE, vbUnicode))
    
    '追加担当者
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(B_ITEMREC.INS_TANTO, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            B_ITEM(Row, ColINS_TANTO) = RTrim(StrConv(B_ITEMREC.INS_TANTO, vbUnicode)) & " " & Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
        Case BtErrKeyNotFound
            B_ITEM(Row, ColINS_TANTO) = StrConv(B_ITEMREC.INS_TANTO, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    '追加日時
    B_ITEM(Row, ColINS_DateTime) = Trim(StrConv(B_ITEMREC.Ins_DateTime, vbUnicode))
    
    
    '更新担当者
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(B_ITEMREC.UPD_TANTO, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
            B_ITEM(Row, ColUPD_TANTO) = RTrim(StrConv(B_ITEMREC.UPD_TANTO, vbUnicode)) & " " & Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
        Case BtErrKeyNotFound
            B_ITEM(Row, ColUPD_TANTO) = StrConv(B_ITEMREC.UPD_TANTO, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Function
    End Select
    '更新日時
    B_ITEM(Row, ColUPD_DateTime) = Trim(StrConv(B_ITEMREC.UPD_DATETIME, vbUnicode))
    
    
    
    Grid_Set_Proc = False
End Function



Private Sub SHORI_Click(Index As Integer)
    Select Case Index
    
        
        Case 0      '更新
        
        
            Command1(Index).Value = True
        
        
        Case 1      '終了
        
        
            Command1(Index).Value = True
        
        Case 2      '検索
        
        
            Command1(Index).Value = True
        
        
                    
    
    End Select

End Sub




Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   データ更新
'----------------------------------------------------------------------------
Dim sts         As Integer
    
Dim i           As Integer
    
Dim com         As Integer
    
Dim Upd_Now     As String
    
    
    Update_Proc = True
                                     
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "編集書き込み処理　開始")
                                     
                                     
    Set TDBGrid1.Array = B_ITEM
    TDBGrid1.Refresh
    
    TDBGrid1.Update
                                     
    If B_ITEM.Count(1) < 1 Then
        Update_Proc = False
        Exit Function
    End If
                                     
    Call Input_Lock
                                    
                                    
    Upd_Now = Format(Now, "YYYYMMDDHHMMSS")
                                    
    For i = 1 To B_ITEM.Count(1)
                                    
        Call UniCode_Conv(K0_B_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_B_ITEM.NAIGAI, DEF_NAIGAI)
        Call UniCode_Conv(K0_B_ITEM.HIN_GAI, B_ITEM(i, ColHIN_GAI))
                                
        sts = BTRV(BtOpGetEqual, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
        Select Case sts
            Case BtNoErr
                
                com = BtOpUpdate
            
            Case BtErrKeyNotFound
                
                com = BtOpInsert
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "美的品番ﾃﾞｰﾀ")
                Exit Function
        End Select
                                        
                                
        If B_ITEM(i, ColSHORI) Then
            If com = BtOpUpdate Then
                sts = BTRV(BtOpDelete, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        
                    
                    Case BtErrKeyNotFound
                        
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "美的品番ﾃﾞｰﾀ")
                        Exit Function
                End Select
            End If
                                
        Else
            
            If com = BtOpInsert Then
                                                                
                Call UniCode_Conv(B_ITEMREC.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(B_ITEMREC.NAIGAI, DEF_NAIGAI)
                Call UniCode_Conv(B_ITEMREC.HIN_GAI, B_ITEM(i, ColHIN_GAI))
            
                
                Call UniCode_Conv(B_ITEMREC.INS_TANTO, Text1(ptxTanto_Code))
                Call UniCode_Conv(B_ITEMREC.Ins_DateTime, Upd_Now)
            
            
                Call UniCode_Conv(B_ITEMREC.UPD_TANTO, "")
                Call UniCode_Conv(B_ITEMREC.UPD_DATETIME, "")
            
            
            Else
            
            
                If RTrim(StrConv(B_ITEMREC.B_HIN_CODE, vbUnicode)) <> RTrim(B_ITEM(i, ColB_HIN_CODE)) Then
            
                    Call UniCode_Conv(B_ITEMREC.UPD_TANTO, Text1(ptxTanto_Code))
                    Call UniCode_Conv(B_ITEMREC.UPD_DATETIME, Upd_Now)
            
            
                End If
            End If



            Call UniCode_Conv(B_ITEMREC.B_HIN_CODE, RTrim(B_ITEM(i, ColB_HIN_CODE)))

               
 

            sts = BTRV(com, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
                Case Else
                    Call File_Error(sts, com, "美的品番ﾃﾞｰﾀ")
                    Exit Function
            End Select
    
    
        End If
    Next i
                                    
                                    
                                    
    Call Input_UnLock
                                        
Call SendMessageStr(hStatusWnd, SB_SETTEXT, 0, "編集書き込み処理　終了")
                                        
                                        
    
    
    Update_Proc = False
    


End Function



Private Sub TDBGrid1_AfterColEdit(ByVal ColIndex As Integer)
    
    
Dim sts     As Integer
    
    If TDBGrid1.Bookmark = Null Then
        Exit Sub
    End If
    
    If TDBGrid1.Bookmark < 1 Then
        Exit Sub
    End If
    
    
    
    
    
    Set TDBGrid1.Array = B_ITEM

    TDBGrid1.Refresh
    
    TDBGrid1.Update

    Select Case ColIndex
    
        
        Case ColHIN_GAI               '対外品番
        
            B_ITEM(TDBGrid1.Bookmark, ColJGYOBU) = Last_JGYOBU
            B_ITEM(TDBGrid1.Bookmark, ColNAIGAI) = DEF_NAIGAI
            
                            
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, DEF_NAIGAI)
            Call UniCode_Conv(K0_ITEM.HIN_GAI, B_ITEM(TDBGrid1.Bookmark, ColHIN_GAI))
                            
                            
                            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    B_ITEM(TDBGrid1.Bookmark, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                Case BtErrKeyNotFound
'                    MsgBox "入力した項目は、エラーです。（品目未登録）" & TDBGrid1.Bookmark & "行目"
                    
'                    TDBGrid1.Col = ColHIN_GAI
'                    TDBGrid1.SetFocus
                

'                    Exit Sub
                
                    B_ITEM(TDBGrid1.Bookmark, ColHIN_NAME) = "（品目未登録）"
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Unload Me
            End Select
    End Select


    Set TDBGrid1.Array = B_ITEM
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.ScrollBars = dbgAutomatic
    TDBGrid1.SetFocus




End Sub

Private Sub TDBGrid1_BeforeInsert(Cancel As Integer)
    
    B_ITEM.ReDim Min_Row, B_ITEM.Count(1), Min_Col, Max_Col

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
        
        
    If Error_Check_Proc(Index) Then     'エラーチェック
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '移動

End Sub
Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   入力項目のエラーチェック
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
    
        Case ptxTanto_Code     '担当者ｺｰﾄﾞ
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text1(ptxTanto_Code).Text)
            
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Text1(ptxTanto_Name).Text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    Text1(ptxTanto_Name).Text = ""
                    MsgBox "入力した項目はエラーです。(担当者)"
                    Text1(ptxTanto_Code).SetFocus
                    Exit Function
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Exit Function
            End Select
            
            
    
    
    End Select
        
        
    Error_Check_Proc = False
    

End Function

Private Function Grid_Error_Check_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド入力項目のエラーチェック
'----------------------------------------------------------------------------
Dim i   As Integer
Dim sts As Integer
    
    Grid_Error_Check_Proc = True
    
    
    
    
    Set TDBGrid1.Array = B_ITEM
    
    
    TDBGrid1.Update
    
    If B_ITEM.Count(1) < 1 Then
        Grid_Error_Check_Proc = False
        Exit Function
    End If
    
    
    
    
    
    
    For i = 1 To B_ITEM.Count(1)
        
        
        
                
        
        
        
        
        If B_ITEM(i, ColSHORI) Then
        Else
        
            If Trim(B_ITEM(i, ColHIN_GAI)) = "" Then
            Else
                
                B_ITEM(i, ColJGYOBU) = Last_JGYOBU
                B_ITEM(i, ColNAIGAI) = DEF_NAIGAI
                
                                
                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_ITEM.NAIGAI, DEF_NAIGAI)
                Call UniCode_Conv(K0_ITEM.HIN_GAI, B_ITEM(i, ColHIN_GAI))
                                
                                
                                
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        B_ITEM(i, ColHIN_NAME) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                    Case BtErrKeyNotFound
                        MsgBox "入力した項目は、エラーです。（品目未登録）" & i & "行目"
                        
                        TDBGrid1.Bookmark = i
                        TDBGrid1.Col = ColHIN_GAI
                        TDBGrid1.SetFocus
                    
                        
                        Exit Function
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                        Exit Function
                End Select
                                
                                
                                
                
                
            
            End If
            
            
        
        End If
    
        
        
        
        
    Next i


    Grid_Error_Check_Proc = False

End Function
