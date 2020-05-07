VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PI00010F2 
   ClientHeight    =   12240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   24765
   _ExtentY        =   21590
   SectionData     =   "PI00010F2.dsx":0000
End
Attribute VB_Name = "PI00010F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Doukon_com      As Integer      '構成／同梱のBtrieve Operation
Private Doukon_eof      As Integer      '構成／同梱 Eof

Private Doukon_cnt      As Integer      '構成／同梱のLINE COUNT


Private SHIJI_QTY       As Double       '今回指示数


Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "KO_NO"               'No
    Me.Fields.Add "KO_HIN_GAI"          '品番
    Me.Fields.Add "KO_SYUBETSU"         '種別
    Me.Fields.Add "KO_QTY"              '員数
    Me.Fields.Add "KO_SHIJI_QTY"        '数量

    Me.Fields.Add "KO_ST_LOCATION"      '棚番
    Me.Fields.Add "KO_ZAIKO_QTY"        '理論在庫
    Me.Fields.Add "KO_ID_NO"            'ID_NO
    Me.Fields.Add "KO_ID_BCR"           'ID_NOﾊﾞｰｺｰﾄﾞ
    Me.Fields.Add "KO_BIKOU"            '備考

    Doukon_com = BtOpGetGreater


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
    
Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
    
Dim SURYO       As String

Dim ST_SOKO     As String
Dim c           As String * 128
    
    If Doukon_cnt > 19 Then
        Exit Sub
    End If
    
    Me.Fields("ko_no").Value = Doukon_Tbl_No(Doukon_cnt)
    
    If Doukon_eof Then
        Me.Fields("KO_HIN_GAI") = " "       '品番
        Me.Fields("KO_SYUBETSU") = " "      '種別
        Me.Fields("KO_QTY") = " "           '員数
        Me.Fields("KO_SHIJI_QTY") = " "     '数量
        Me.Fields("KO_ST_LOCATION") = " "   '棚番
        Me.Fields("KO_ZAIKO_QTY") = " "     '理論在庫
        Me.Fields("KO_ID_NO") = " "         'ID_NO
        Me.Fields("KO_ID_BCR") = " "        'ID_NOﾊﾞｰｺｰﾄﾞ
        Me.Fields("KO_BIKOU") = " "         '備考
    
        
    Else
        sts = BTRV(Doukon_com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Taget_SHIMUKE_CODE_KEY Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Taget_JGYOBU_key Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Taget_NAIGAI_key Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Taget_Hin_key)) Or _
                    StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                    Doukon_eof = True
                End If
            Case BtErrEOF
                
                Doukon_eof = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "構成マスタ（親）")
                Exit Sub
        
        End Select
                                            
        If Doukon_eof Then
            Me.Fields("KO_HIN_GAI") = " "        '品番
            Me.Fields("KO_SYUBETSU") = " "       '種別
            Me.Fields("KO_QTY") = " "            '員数
            Me.Fields("KO_SHIJI_QTY") = " "      '数量
            Me.Fields("KO_ST_LOCATION") = " "    '棚番
            Me.Fields("KO_ZAIKO_QTY") = " "      '理論在庫
            Me.Fields("KO_ID_NO") = " "          'ID_NO
            Me.Fields("KO_ID_BCR") = " "         'ID_NOﾊﾞｰｺｰﾄﾞ
            Me.Fields("KO_BIKOU") = " "          '備考
                                            
                                            
                                            
        Else
                                                '品番
            Me.Fields("KO_HIN_GAI") = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                                                '種別
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                    Exit Sub
            
            End Select
            Me.Fields("KO_SYUBETSU") = StrConv(P_CODEREC.C_RNAME, vbUnicode)
                                                '員数
            If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
            Else
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
            End If
        
        
        
            '品目マスタ読み込み
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    
                    
                    Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Sub
    
            End Select
        
        
        
            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Me.Fields("KO_ST_LOCATION") = ""
            Else
                '標準棚番
                
                ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                'P_SYS.INI--> PI00010.INI 2011.08.04
                If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                Else
                    ST_SOKO = RTrim(c)
                End If
                
                
                
                Me.Fields("KO_ST_LOCATION") = Trim(ST_SOKO) & "-" & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If
        
        
            '在庫数
            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, StrConv(ITEMREC.JGYOBU, vbUnicode), _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Sub
            
            End If
            Me.Fields("KO_ZAIKO_QTY") = Format(Sumi_Qty + Mi_Qty, "#0")
            '備考OR出荷ﾊﾞｰｺｰド
            If PRI_BIKOU_BCR Then
                                                                                        'ID_NO
                Me.Fields("KO_ID_NO") = ""
                                                                                    'ID_NOﾊﾞｰｺｰﾄﾞ
                Me.Fields("KO_ID_BCR") = ""
            Else
                Me.Fields("KO_BIKOU") = StrConv(P_COMPO_K_REC.KO_BIKOU, vbUnicode)     '備考
        
            End If
        End If
            
    
    
    
    
    
    
        Doukon_com = BtOpGetNext
    End If
    
    
    
    Doukon_cnt = Doukon_cnt + 1
    
            
    eof = False
    
    
    


End Sub

Private Sub ActiveReport_Initialize()

Dim sts             As Integer

Dim cnt             As Integer
Dim com             As Integer


Dim i               As Integer
Dim Total_Times     As Double
Dim AVE             As Double


Dim SURYO           As String

Dim ST_SOKO         As String
Dim c               As String * 128

Dim Target          As Double

Dim wkValue         As String
Dim wkEDIT_NIN      As String
Dim wkEDIT_TIMES    As String
Dim wkAVE           As String




    '構成マスタ　ﾍｯﾀﾞｰ
    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")

    
    
    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "構成マスタ")
            Exit Sub
    
    End Select



    '仕向け先名
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, Taget_SHIMUKE_CODE_KEY)
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Sub
    
    End Select
       
    Field1.text = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))       '仕向け先名
    
    Field2.text = ""
    Field3.text = Format(Now, "YYYY/MM/DD HH:MM")                   '発行日時

    Field4.text = ""                                                '承認者
    
    '担当者
    Field5.text = ""                                                '担当者
    
    '収単／担当者
    lblS_Tanto1.Visible = False
    fldS_Tanto.Visible = False
    speS_tanto1.Visible = False
    l_S_Tanto1.Visible = False
    fldS_Tanto.text = ""                                        '収単／担当者
    
    
    lblSHIJI_F.Caption = ""
    
    
    
    
    
    
    Field7.text = Taget_Hin_key                                 '品番
                                                                '数量
    SHIJI_QTY = SHIJI_QTY
    Field8.text = Format(SHIJI_QTY, "#0")
    '品名／棚番
    Call UniCode_Conv(K0_ITEM.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Taget_Hin_key)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
            Call UniCode_Conv(ITEMREC.ST_RETU, "")
            Call UniCode_Conv(ITEMREC.ST_REN, "")
            Call UniCode_Conv(ITEMREC.ST_DAN, "")
        
            Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Sub
    
    End Select
    Field9.text = StrConv(ITEMREC.HIN_NAME, vbUnicode)                      '品名

    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
        Field10.text = ""                                                   '標準棚番
    Else
        ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        'P_SYS.INI--> PI00010.INI 2011.08.04
        If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
        Else
            ST_SOKO = RTrim(c)
        End If
        
        
        Field10.text = Trim(ST_SOKO) & "-" & _
                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_DAN, vbUnicode)
    End If

    Field11.text = Trim(StrConv(P_COMPO_O_REC.CLASS_CODE, vbUnicode))       '商品化ｸﾗｽ
    Field12.text = Trim(StrConv(P_COMPO_O_REC.F_CLASS_CODE, vbUnicode))    '付加ｸﾗｽ
    Field13.text = Trim(StrConv(P_COMPO_O_REC.N_CLASS_CODE, vbUnicode))    '内職ｸﾗｽ


    'ラベル貼付計上有無
    If StrConv(ITEMREC.G_LABEL_NON, vbUnicode) = P_G_LABEL_OFF Then
        lblLabel_NIN.Caption = "******"
        lblLabel_TIMES.Caption = "******"
    Else
        lblLabel_NIN.Caption = ""
        lblLabel_TIMES.Caption = ""
    End If


    Field14.text = ""                                                       '商品化手配先
    

    '個装資材のループ
    cnt = 0

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Taget_SHIMUKE_CODE_KEY Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Taget_JGYOBU_key Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Taget_NAIGAI_key Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Taget_Hin_key)) Or _
                    StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_KOSOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "構成マスタ（子）")
                Exit Sub
        
        End Select
        '品目マスタ読み込み
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                '個装資材№
                Field15.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field16.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field16.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
        
                Field17.text = ""
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field18.text = ""
                Else
                    
                    
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    
                    Field18.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                '個装資材№
                Field19.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field20.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field20.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field21.text = ""
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field22.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field22.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                '個装資材№
                Field23.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field24.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field24.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field25.text = ""
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field26.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field26.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
            Case 4
            
                '個装資材№
                Field27.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field28.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field28.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field29.text = ""
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field30.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field30.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 5
                '個装資材№
                Field31.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field32.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field32.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                Field33.text = ""
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field34.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field34.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
        
        End Select
        com = BtOpGetNext
    
    Loop


    '外装資材のループ
    cnt = 0

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
                If (StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> Taget_SHIMUKE_CODE_KEY Or _
                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> Taget_JGYOBU_key Or _
                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> Taget_NAIGAI_key Or _
                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(Taget_Hin_key)) Or _
                    StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "構成マスタ（子）")
                Exit Sub
        
        End Select
        '品目マスタ読み込み

'>>>>>>>>>>>>>>>>>>>>   2012.10.19
'        Call UniCode_Conv(K0_ITEM.JGYOBU, Taget_JGYOBU_key)
'        Call UniCode_Conv(K0_ITEM.NAIGAI, Taget_NAIGAI_key)
'        Call UniCode_Conv(K0_ITEM.HIN_GAI, Taget_Hin_key)
        
        
        
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))          '2012.10.19
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))          '2012.10.19
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))        '2012.10.19
'>>>>>>>>>>>>>>>>>>>>   2012.10.19
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                '外装資材№
                Field35.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '外装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field36.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field36.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                
                
                Field37.text = ""
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field38.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field38.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                '外装資材№
                Field39.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '外装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field40.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field40.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                
                Field41.text = ""
                
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field42.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field42.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                '外装資材№
                Field43.text = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                '外装資材　員数
                If CInt(Right(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field44.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field44.text = Format(CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                Field45.text = ""
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field46.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                    'P_SYS.INI--> PI00010.INI 2011.08.04
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field46.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
        
        End Select
    
        com = BtOpGetNext
    
    Loop

    Field47.text = Trim(StrConv(P_COMPO_O_REC.BIKOU, vbUnicode))               '備考
    
    
        
    '作業日／数量／担当
    ShaSagyo_Day.Visible = PRI_SAGYO_DAY
    LineSagyo_Day1.Visible = PRI_SAGYO_DAY
    LineSagyo_Day2.Visible = PRI_SAGYO_DAY
    LineSagyo_Day3.Visible = PRI_SAGYO_DAY
    lblSagyo_day1.Visible = PRI_SAGYO_DAY
    lblSagyo_day2.Visible = PRI_SAGYO_DAY
    lblSagyo_day3.Visible = PRI_SAGYO_DAY
        
    
    
    '見本作成の印字有無
    lblSample.Visible = False
    Shape10.Visible = False

    
    
    'ﾒｲﾝﾊﾞｰｺｰﾄﾞ
    fldMain_Bcr.text = ""

    
    '明細備考
    fldBIKOU.Visible = False

    
    fldSyuka_No.Visible = False
    fldSyuka_Bcr.Visible = False

    '同梱部品   2011.08.04
'    lblDOUKON.Visible = PRI_DOUKON
'    lblDOUKON_GOUHI.Visible = PRI_DOUKON

    Select Case PRI_DOUKON
        Case 0
            lblDOUKON.Visible = False
            lblKOKUIN.Visible = False
            lblDOUKON_GOUHI.Visible = False
        Case 1
            lblDOUKON.Visible = True
            lblKOKUIN.Visible = False
            lblDOUKON_GOUHI.Visible = True
        Case 2
            lblDOUKON.Visible = False
            lblKOKUIN.Visible = True
            lblDOUKON_GOUHI.Visible = False
    End Select
    '同梱部品   2011.08.04



    '構成／同梱
    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    '入庫完了印
    l_Nyuko_IN1.Visible = PRI_NYUKO_IN
    l_Nyuko_IN2.Visible = PRI_NYUKO_IN
    l_Nyuko_IN3.Visible = PRI_NYUKO_IN
    l_Nyuko_IN4.Visible = PRI_NYUKO_IN

    lblNyuko_In.Visible = PRI_NYUKO_IN

    '入力完了印
    l_Input_IN1.Visible = PRI_INPUT_IN
    l_Input_IN2.Visible = PRI_INPUT_IN
    l_Input_IN3.Visible = PRI_INPUT_IN
    l_Input_IN4.Visible = PRI_INPUT_IN

    lblInput_In.Visible = PRI_INPUT_IN


    If Not PRI_NYUKO_IN And Not PRI_NYUKO_IN Then
        l_IN_Center.Visible = False
    Else
        l_IN_Center.Visible = True
    End If

    lblBunnou.Visible = False

    '下部　品番／№／数量   2007.05.22
    ShaHINBAN_BIKOU.Visible = PRI_HINBAN_BIKOU
    
    LineHINBAN_BIKOU1.Visible = PRI_HINBAN_BIKOU
    LineHINBAN_BIKOU2.Visible = PRI_HINBAN_BIKOU
    LineHINBAN_BIKOU3.Visible = PRI_HINBAN_BIKOU
    LineHINBAN_BIKOU4.Visible = PRI_HINBAN_BIKOU

    lblHINBAN_BIKOU1.Visible = PRI_HINBAN_BIKOU
    lblHINBAN_BIKOU2.Visible = PRI_HINBAN_BIKOU
    lblHINBAN_BIKOU3.Visible = PRI_HINBAN_BIKOU

    Field60.Visible = PRI_HINBAN_BIKOU
    Field61.Visible = PRI_HINBAN_BIKOU
    Field62.Visible = PRI_HINBAN_BIKOU
    
    Field60.text = StrConv(P_COMPO_O_REC.HIN_GAI, vbUnicode)           '品番
'2011.08.04    Field61.text = "00000"                                              '№
                                                                        '数量
    Field62.text = Format(0, "#0")



    '自責タイトル
    lblBEF_JISSEKI.Caption = ""
    lblBEF_BEFORE1.Caption = ""
    lblBEF_BEFORE2.Caption = ""
    lblBEF_BEFORE3.Caption = ""
    lblBEF_BEFORE4.Caption = ""
    lblBEF_SAGYO1.Caption = ""
    lblBEF_SAGYO2.Caption = ""
    lblBEF_SAGYO3.Caption = ""
    lblBEF_AFTER1.Caption = ""
    lblBEF_AFTER2.Caption = ""
    lblBEF_JISEKI.Caption = ""
    lblBEF_TASEKI.Caption = ""
            

    lblBunnou.Visible = False


    lblTarget1.Visible = True
    lblTarget2.Visible = True
    

    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
    
    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")


    Doukon_com = BtOpGetGreater
    Doukon_eof = False

    Doukon_cnt = 0



' 2013.01.08 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    If GENSANKOKU_MSG_F Then                            '2013.02.19
        If Trim(chk_TORI_GENSANKOKU) <> "" Then
            GENSANKOKU_Alart.Visible = True
        End If
    End If                                              '2013.02.19
' 2013.01.08 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


End Sub

Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORPortrait
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 10
    Me.PageTopMargin = 10
    Me.PageLeftMargin = 20
    Me.PageRightMargin = 20

    Me.documentName = "商品化指図票："

    DoEvents

End Sub

