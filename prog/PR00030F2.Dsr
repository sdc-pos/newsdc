VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00030F2 
   ClientHeight    =   12405
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Μωθl
   _ExtentX        =   33655
   _ExtentY        =   21881
   SectionData     =   "PR00030F2.dsx":0000
End
Attribute VB_Name = "PR00030F2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      'ΎΧΜBtrieve Operation








Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "DET_HIN_GAI"         'iΤO
    Me.Fields.Add "DET_HIN_NAME"        'iΌ
    Me.Fields.Add "DET_G_SYUSHI"        'έΙ³
    Me.Fields.Add "DET_ZEN_ZAIKO_QTY"   'OέΙ
    Me.Fields.Add "DET_NYUKO_QTY"       'όΙ
    Me.Fields.Add "DET_SYUKO_QTY"       'oΙ
    Me.Fields.Add "DET_ZAIKO_QTY"       'έΙ
    Me.Fields.Add "DET_SHI_TANKA"       'dόPΏ
    Me.Fields.Add "DET_ZAIKO_KIN"       'έΙPΏ
    Me.Fields.Add "DET_SHI_CODE"        'dόζ
    Me.Fields.Add "DET_LAST_SYUKA_DT"   'ΕIoΧϊ
    Me.Fields.Add "DET_LAST_SYUKA_QTY"  'ΕIoΧ
    Me.Fields.Add "DET_MAEGARI_QTY"     'OΨ





End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

    
    
'    sts = BTRV(DET_com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K1_tmpP_STOCK, Len(K1_tmpP_STOCK), 1)
    sts = BTRV(DET_com, tmpP_STOCK_POS, tmpP_STOCK_REC, Len(tmpP_STOCK_REC), K2_tmpP_STOCK, Len(K2_tmpP_STOCK), 2)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "tmpήI΅Γή°ΐ")
            Exit Sub
    End Select
    
    
    
If Trim(StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode)) = "K203" Then
    Debug.Print
End If
    
    
    'iΤ
    Me.Fields("DET_HIN_GAI") = StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode)
    'iΌ
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(tmpP_STOCK_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(tmpP_STOCK_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(tmpP_STOCK_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Me.Fields("DET_HIN_NAME") = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Me.Fields("DET_HIN_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "iΪΟ½ΐ")
            Exit Sub
    End Select
    'έΙ³
    Me.Fields("DET_G_SYUSHI") = StrConv(tmpP_STOCK_REC.G_SYUSHI, vbUnicode)
    'OέΙ
    If IsNumeric(StrConv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)) Then
        Me.Fields("DET_ZEN_ZAIKO_QTY") = Format(CLng(StrConv(tmpP_STOCK_REC.ZEN_ZAIKO_QTY, vbUnicode)), "#,###")
    Else
        Me.Fields("DET_ZEN_ZAIKO_QTY") = ""
    End If
    'όΙ
        'Clng --> Val 2016.01.08
    Me.Fields("DET_NYUKO_QTY") = Format(Val(StrConv(tmpP_STOCK_REC.NYUKO_QTY, vbUnicode)), "#,###")
    'oΙ
        'Clng --> Val 2016.01.08
    Me.Fields("DET_SYUKO_QTY") = Format(Val(StrConv(tmpP_STOCK_REC.SYUKO_QTY, vbUnicode)), "#,###")
    'έΙ
        'Clng --> Val 2016.01.08
    Me.Fields("DET_ZAIKO_QTY") = Format(Val(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,###")
    'dόPΏ
        'CDbl --> Val 2016.01.08
    Me.Fields("DET_SHI_TANKA") = Format(Val(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)), "#,##0.00")
    'έΙΰz
'    Me.Fields("DET_ZAIKO_KIN") = Format(CDbl(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) * CLng(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)), "#,###")
'
    '>>>>>>>>>> 2016.01.08
    If Not IsNumeric(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) Then
        Call UniCode_Conv(tmpP_STOCK_REC.TANKA, "00000000")
    End If
    If Not IsNumeric(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)) Then
        Call UniCode_Conv(tmpP_STOCK_REC.ZAIKO_QTY, "00000000")
    End If
    '>>>>>>>>>> 2016.01.08
    
    Me.Fields("DET_ZAIKO_KIN") = Format(ToRoundUp(CCur(StrConv(tmpP_STOCK_REC.TANKA, vbUnicode)) * _
                                    CCur(StrConv(tmpP_STOCK_REC.ZAIKO_QTY, vbUnicode)), 0), "#,###")
    
    'έΙ³
    Me.Fields("DET_SHI_CODE") = StrConv(tmpP_STOCK_REC.CODE, vbUnicode)
    'ΕIoΧϊ
    If Trim(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode)) = "" Then
        Me.Fields("DET_LAST_SYUKA_DT") = ""
    Else
        Me.Fields("DET_LAST_SYUKA_DT") = Left(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 4) & "/" & _
                                            Mid(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                           Right(StrConv(tmpP_STOCK_REC.LAST_SYUKA_DT, vbUnicode), 2)
    End If
    
    Me.Fields("DET_LAST_SYUKA_QTY") = ""
    If IsNumeric(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) Then
        If CLng(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)) <> 0 Then
            Me.Fields("DET_LAST_SYUKA_QTY") = Format(CLng(StrConv(tmpP_STOCK_REC.LAST_SYUKA_QTY, vbUnicode)), "#,###")
        End If
    End If
    'OΨ
    If IsNumeric(StrConv(tmpP_STOCK_REC.MAEGARI_QTY, vbUnicode)) Then
        Me.Fields("DET_MAEGARI_QTY") = Format(CDbl(StrConv(tmpP_STOCK_REC.MAEGARI_QTY, vbUnicode)), "#,###")
    Else
        Me.Fields("DET_MAEGARI_QTY") = ""
    End If
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    
    

End Sub

Private Sub ActiveReport_Initialize()

    
    
Dim sts     As Integer
    
    
    
    
    
    
    
    
    
    
    DET_com = BtOpGetFirst
    
    lblYY.Caption = Left(PR000301.Text1(1).Text, 4)
    lblMM.Caption = Right(PR000301.Text1(1).Text, 2)
 
 
    SHIME_YMD.Caption = "χϊF" & PR000301.Text1(0)
 
 
 
 
    Center_Name.Caption = StrConv(P_KANRIREC.Center_Name, vbUnicode)



    DET_com = BtOpGetFirst
 


End Sub


Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    lblPrint_Now = Format(Now, "YYYY/MM/DD HH:MM")
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 25
    Me.PageRightMargin = 25

    Me.documentName = "ήI΅\"

    DoEvents

End Sub

