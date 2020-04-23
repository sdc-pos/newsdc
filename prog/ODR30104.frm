VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ODR30104 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "親部品注文情報"
   ClientHeight    =   4755
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   10350
   StartUpPosition =   2  '画面の中央
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   3660
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "キャンセル"
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
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   1800
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   3435
      Left            =   315
      TabIndex        =   1
      Top             =   960
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   6059
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "親品番"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "受注総数"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ＫＥＹ項目"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "必要数"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "不足数"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "注文納期"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "使用月"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "可能日"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   688
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4075"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3942"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2117"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1984"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8194"
      Splits(0)._ColumnProps(10)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2778"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2117"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1984"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2117"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1984"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2302"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2170"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1640"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=1"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=2461"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=2328"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=1"
      Splits(0)._ColumnProps(41)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ Ｐゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "親部品　注文情報"
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      InsertMode      =   0   'False
      DeadAreaBackColor=   -2147483643
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ Ｐゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFF00&,.bold=0,.fontsize=1125"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=1125,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&H80FF00&"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40,.bgcolor=&HFF80&"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=87,.parent=1,.bgcolor=&H80FF80&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=96,.parent=4,.bgcolor=&H80FF00&"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=89,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=91,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=94,.parent=9,.bgcolor=&H80FF80&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=95,.parent=10,.bgcolor=&H80FF00&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=106,.parent=87,.alignment=3,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=103,.parent=88"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=104,.parent=89"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=105,.parent=91"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=16,.parent=87,.alignment=1,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=88"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=89"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=91"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=138,.parent=87"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=135,.parent=88"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=136,.parent=89"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=137,.parent=91"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=20,.parent=87,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=88"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=89"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=91"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=87,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=88"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=89"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=91"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=87,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=88"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=89"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=91"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=87,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=88"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=89"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=91"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=24,.parent=87,.alignment=2"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=88"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=89"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=91"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&H80FF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "ODR30104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'コンボ用添字
'Private Const pcmbHBUN = 0



'グリッド用定義
Private ORDR_GRID   As New XArrayDB


Private Const Min_Row% = 1              '最小行数
Private Const Max_Row = 9999            '最大行数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 7              '最大列数

Private Const Col_ITEM% = 0             '部品コード
Private Const Col_QTY% = 1              '使用数量
Private Const Col_Req% = 3              '必要数量
Private Const Col_MAI% = 4              '不足数         '2010/03/04追加（デバッグ用）
Private Const Col_NOUKI% = 5            '納　期         ' 〃
Private Const Col_USE% = 6              '使用月         ' 〃
Private Const Col_OK_DT% = 7            '可能日

Dim row         As Long                 '対象　行

Dim Cor_Row     As Long                 'カレント行

Dim Init_F      As Integer

Dim svJGYOBU    As String '* 1
Dim svNAIGAI    As String '* 1
Dim svHin_gai   As String '* 20
Dim svUSE       As String
Dim svNOUKI     As String

Dim sumQty      As Double
Dim sumOdr      As Double
Dim sumMAI      As Double

'Dim sumReq      As Double

Dim W_OK_DT     As String

Private Function Data_Disp() As Integer
Dim com         As Integer
Dim sts         As Integer
Dim yn          As Integer

Dim W_KEY_USE_YM    As String
Dim W_Key           As String
Dim X_i         As Integer


    Data_Disp = True
    
    row = Min_Row - 1
    Call Input_Lock                             '画面項目ロック
                                '所要量Ｆ
    If ODR_REQUIRE_Open(BtOpenNomal) Then
        Exit Function
    End If
    
    DoEvents
    
    Set ORDR_GRID = Nothing
    
    
    row = 0
    
    W_KEY_USE_YM = Left(Key_USE_YM, 4) & "/" & Right(Key_USE_YM, 2) & "/01"
    
    For X_i = 0 To 1
    
    W_Key = Format(DateAdd("m", X_i, W_KEY_USE_YM), "yyyy/mm/dd")
    W_Key = Left(W_Key, 4) & Mid(W_Key, 6, 2)
    
    Call UniCode_Conv(K3_ODR_REQ.USE_YM, W_Key)
    
    Call UniCode_Conv(K3_ODR_REQ.KO_JGYOBU, Key_JIGYOBU)
    Call UniCode_Conv(K3_ODR_REQ.KO_NAIGAI, Key_NAIGAI)
    Call UniCode_Conv(K3_ODR_REQ.KO_HIN_GAI, Key_HinGai)
    
'    Call UniCode_Conv(K3_ODR_REQ.SHIMUKE, GW_SIMUKE)
    Call UniCode_Conv(K3_ODR_REQ.SHIMUKE, "")
    
    Call UniCode_Conv(K3_ODR_REQ.JGYOBU, "")
    Call UniCode_Conv(K3_ODR_REQ.NAIGAI, "")
    Call UniCode_Conv(K3_ODR_REQ.HIN_GAI, "")
    Call UniCode_Conv(K3_ODR_REQ.ORDER_NO, "")
    Call UniCode_Conv(K3_ODR_REQ.INS_NO, "")
    Call UniCode_Conv(K3_ODR_REQ.BUN_NO, "")
    
    svJGYOBU = ""
    svNAIGAI = ""
    svHin_gai = ""
    svUSE = ""
    svNOUKI = ""
    sumQty = 0
    sumOdr = 0
    sumMAI = 0
    'sumReq = 0
    
    
    W_OK_DT = ""
    
    com = BtOpGetGreaterEqual
        
    Do
        sts = BTRV(com, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K3_ODR_REQ, Len(K3_ODR_REQ), 3)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                Exit Do
                        
            Case Else
                Call File_Error(sts, com, "ODR_REQ")
                GoTo Err_exit
        End Select
        
        If StrConv(ODR_REQ_R.USE_YM, vbUnicode) <> W_Key Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_REQ_R.KO_JGYOBU, vbUnicode)) <> Trim(Key_JIGYOBU) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_REQ_R.KO_NAIGAI, vbUnicode)) <> Trim(Key_NAIGAI) Then
            Exit Do
        End If
        
        If Trim(StrConv(ODR_REQ_R.KO_HIN_GAI, vbUnicode)) <> Trim(Key_HinGai) Then
            Exit Do
        End If
               
''        If Trim(StrConv(ODR_REQ_R.SHIMUKE, vbUnicode)) <> GW_SIMUKE Then
''            Exit Do
''        End If
            
            
'        If CInt(StrConv(ODR_REQ_R.BUN_NO, vbUnicode)) <> 0 Then
            
            
            If Trim(svHin_gai) = "" Then
                
                svJGYOBU = Trim(StrConv(ODR_REQ_R.JGYOBU, vbUnicode))
                svNAIGAI = Trim(StrConv(ODR_REQ_R.NAIGAI, vbUnicode))
                svHin_gai = Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode))
                W_OK_DT = Trim(StrConv(ODR_REQ_R.OK_DT, vbUnicode))
                
            End If
                
                
            If Trim(svJGYOBU) <> Trim(StrConv(ODR_REQ_R.JGYOBU, vbUnicode)) Or _
                Trim(svNAIGAI) <> Trim(StrConv(ODR_REQ_R.NAIGAI, vbUnicode)) Or _
                Trim(svHin_gai) <> Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode)) Then
                
                
                
                '編集
                If sumOdr <> 0 Then
                    row = row + 1
                    If Grid_Set_Proc() Then
                        GoTo Err_exit
                    End If
                End If
                
                svJGYOBU = Trim(StrConv(ODR_REQ_R.JGYOBU, vbUnicode))
                svNAIGAI = Trim(StrConv(ODR_REQ_R.NAIGAI, vbUnicode))
                svHin_gai = Trim(StrConv(ODR_REQ_R.HIN_GAI, vbUnicode))
                
                W_OK_DT = Trim(StrConv(ODR_REQ_R.OK_DT, vbUnicode))
                
                svUSE = Trim(StrConv(ODR_REQ_R.USE_YM, vbUnicode))
                svNOUKI = Trim(StrConv(ODR_REQ_R.CYUMON_DT, vbUnicode))
                sumQty = 0
                sumOdr = 0
                sumMAI = 0
            End If
                        
                        
            Call UniCode_Conv(K0_ODR_ORDER.SHIMUKE, StrConv(ODR_REQ_R.SHIMUKE, vbUnicode))
            Call UniCode_Conv(K0_ODR_ORDER.JGYOBU, StrConv(ODR_REQ_R.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ODR_ORDER.NAIGAI, StrConv(ODR_REQ_R.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ODR_ORDER.HIN_GAI, StrConv(ODR_REQ_R.HIN_GAI, vbUnicode))
            Call UniCode_Conv(K0_ODR_ORDER.ORDER_NO, StrConv(ODR_REQ_R.ORDER_NO, vbUnicode))
            Call UniCode_Conv(K0_ODR_ORDER.INS_NO, StrConv(ODR_REQ_R.INS_NO, vbUnicode))
            Call UniCode_Conv(K0_ODR_ORDER.BUN_NO, StrConv(ODR_REQ_R.BUN_NO, vbUnicode))
            
            
            'sts = BTRV(com, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            sts = BTRV(BtOpGetEqual, ODR_ORDER_POS, ODR_ORDER_REC, Len(ODR_ORDER_REC), K0_ODR_ORDER, Len(K0_ODR_ORDER), 0)
            Select Case sts
                Case BtNoErr
                    
                    If CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode)) > 0 Then
                    sumQty = sumQty + CDbl(StrConv(ODR_ORDER_REC.ODR_QTY, vbUnicode))
                    End If
                    
                    
                    
                    '           正規の算定
                    If CDbl(StrConv(ODR_REQ_R.ODR_QTY, vbUnicode)) > 0 Then
                        sumOdr = sumOdr + CDbl(StrConv(ODR_REQ_R.ODR_QTY, vbUnicode))
                    End If
            
                    '                   デバッグ時は下記に！
                    '           展開数
                    'If CDbl(StrConv(ODR_REQ_R.REQ_QTY, vbUnicode)) > 0 Then
                    '    sumOdr = sumOdr + CDbl(StrConv(ODR_REQ_R.REQ_QTY, vbUnicode))
                    'End If
                    
                    
                    '   不足数  '2010/03/04
                    If CDbl(StrConv(ODR_REQ_R.FUSOKU_QTY, vbUnicode)) > 0 Then
                        sumMAI = sumMAI + CDbl(StrConv(ODR_REQ_R.FUSOKU_QTY, vbUnicode))
                    End If
                    
                    
                    'If CDbl(StrConv(ODR_REQ_R.REQ_QTY, vbUnicode)) > 0 Then
                    '    sumReq = sumReq + CDbl(StrConv(ODR_REQ_R.REQ_QTY, vbUnicode))
                    'End If
                    
                    If W_OK_DT < StrConv(ODR_REQ_R.OK_DT, vbUnicode) Then
                        W_OK_DT = StrConv(ODR_REQ_R.OK_DT, vbUnicode)
                    End If
                    
                    svUSE = Trim(StrConv(ODR_REQ_R.USE_YM, vbUnicode))
                    svNOUKI = Trim(StrConv(ODR_REQ_R.CYUMON_DT, vbUnicode))
                Case BtErrKeyNotFound, BtErrEOF      'レコード無し
                            
                Case Else
                    Call File_Error(sts, com, "ODR_REQ")
                    GoTo Err_exit
            End Select
            
            
            
            
        
'        End If
        
        com = BtOpGetNext
    Loop
    
    
    If Trim(svJGYOBU) <> "" Then
        
        '編集
        If sumOdr <> 0 Then
        
            row = row + 1
            If Grid_Set_Proc() Then
                GoTo Err_exit
            End If
        End If
    End If
    
    Next X_i
    
    
    Set TDBGrid1.Array = ORDR_GRID
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    TDBGrid1.MoveFirst
    TDBGrid1.ScrollBars = dbgAutomatic
    
    DoEvents
    
    Call Input_UnLock                             '画面項目ロック
    
    Data_Disp = False
    
Err_exit:
    
    sts = BTRV(BtOpClose, ODR_REQ_POS, ODR_REQ_R, Len(ODR_REQ_R), K0_ODR_REQ, Len(K0_ODR_REQ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "ODR_REQ")
        End If
    End If
    
    
End Function
Private Function Grid_Set_Proc() As Integer
'----------------------------------------------------------------------------
'                   グリッド表示（移動歴データ内容）
'               Row   行数
'               mode　FALSE:ﾁｪｯｸOFF  TRUE:ﾁｪｯｸON
'----------------------------------------------------------------------------
Dim W_STR       As String



    Grid_Set_Proc = True

    ORDR_GRID.ReDim Min_Row, row, Min_Col, Max_Col
    '品番
    ORDR_GRID(row, Col_ITEM) = svHin_gai
    '数量
    ORDR_GRID(row, Col_QTY) = Format(sumQty, "###,##0.00")
    
    '必要数
    ORDR_GRID(row, Col_Req) = Format(sumOdr, "###,##0.00")
    
    '不足数
    If sumMAI = 0 Then
        ORDR_GRID(row, Col_MAI) = ""
    Else
        ORDR_GRID(row, Col_MAI) = Format(sumMAI, "###,##0.00")
    End If
    
    '使用月
    If Trim(svUSE) <> "" Then
        ORDR_GRID(row, Col_USE) = Left(svUSE, 4) & "/" & Right(svUSE, 2)
    Else
        ORDR_GRID(row, Col_USE) = ""
    End If
    
    '注文納期
    If Trim(svNOUKI) <> "" Then
        ORDR_GRID(row, Col_NOUKI) = Left(svNOUKI, 4) & "/" & Mid(svNOUKI, 5, 2) & "/" & Right(svNOUKI, 2)
    Else
        ORDR_GRID(row, Col_NOUKI) = ""
    End If
    
    '可能日
    If Trim(W_OK_DT) = "" Then
        W_STR = ""
    Else
        W_STR = Mid(W_OK_DT, 3, 2) & "/" & Mid(W_OK_DT, 5, 2) & "/" & Right(W_OK_DT, 2)
    End If
    ORDR_GRID(row, Col_OK_DT) = W_STR
    
    Grid_Set_Proc = False

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i As Integer

    ODR30104.MousePointer = vbHourglass

    Call Ctrl_Lock(ODR30104)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(ODR30104)


    ODR30104.MousePointer = vbDefault

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
    
            
        Case 0
            Init_F = 0
            Set ORDR_GRID = Nothing
            Set TDBGrid1.Array = ORDR_GRID
            TDBGrid1.ReBind
            TDBGrid1.Update
            DoEvents
            
            ODR30104_Return = True                '確認画面ｷｬﾝｾﾙ終了
            Me.Visible = False
    
    End Select

End Sub

Private Sub Form_Activate()
Dim X_i As Integer

    If Init_F <> 0 Then Exit Sub
    
    ODR30104.Top = ODR30101.Top + (ODR30101.Height - ODR30104.Height)
    ODR30104.Left = ODR30101.Left + (ODR30101.Width - ODR30104.Width) / 2
    
    
    
    If Data_Disp Then
        Call Input_UnLock                             '画面項目ロック
        ODR30104_Return = True                '確認画面ｷｬﾝｾﾙ終了
        Me.Visible = False
    End If
    
    ODR30104_Return = True
    TDBGrid1.SetFocus
    
    
    Init_F = 1
    
End Sub

Private Sub Form_Load()

    Init_F = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'If UnloadMode <> 0 Then Exit Sub
    Me.Visible = False
    
End Sub


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)

'    TDBGrid1.Bookmark = -1         '2016.02.15
    
End Sub

