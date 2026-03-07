Partial Class CP_01_001
    Inherits AuthBasePage

    'CP_01_001_97
    'CP_01_001_98
    'CP_01_001_99
    'CP_01_001_03 (2013)'空白
    '資料
    'CP_01_001_add_99 (暫用)
    'XX CP_01_001_add_03 (2013)(無法使用暫用99年) 'CLASS_VISITOR3

    '(CLASS_VISITOR3) CLASS_VISITOR3DELDATA
    '(CLASS_VISITOR) CLASS_VISITORDELDATA
    'VIEW_VISITOR(CLASS_VISITOR)
    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False
    Dim rqDOCID As String = ""

    'Cells/Column
    Const cst_序號 As Integer = 0
    Const cst_訓練機構 As Integer = 1
    'Const cst_班別名稱 As Integer = 2
    'Const cst_期別 As Integer = 3
    'Const cst_訪視單位 As Integer = 4
    'Const cst_開訓日期 As Integer = 5
    'Const cst_結訓日期 As Integer = 6
    'Const cst_訪查日期 As Integer = 7
    'Const cst_訪查人員 As Integer = 8 'Item32
    'Const cst_訪查結果 As Integer = 9 'Item32
    'Const cst_綜合建議 As Integer = 10 'Item31Note
    Const cst_功能 As Integer = 11

    '空白表
    Const cst_printFN1 As String = "CP_01_001_97" '2008年之前。
    Const cst_printFN2 As String = "CP_01_001_98" '2009年
    Const cst_printFN3 As String = "CP_01_001_99" '2010年啟用。
    'Const cst_printFN4 As String = "CP_01_001_03" '2014年啟用。'07:接受企業委託訓練
    Const cst_printFN5 As String = "CP_01_001_06" '2017年啟用。(在職)'07:接受企業委託訓練
    Const cst_printFN6 As String = "CP_01_001_19" '2019年啟用。70:區域產業據點職業訓練計畫(在職)
    'iReport報表
    Const cst_printFN0add As String = "CP_01_001_Rpt" '2007年之前。
    Const cst_printFN1add As String = "CP_01_001_add_97" '2008年之前。
    Const cst_printFN2add As String = "CP_01_001_add_98" '2009年
    Const cst_printFN3add As String = "CP_01_001_add_99" '2010年啟用。
    'Const cst_printFN4add As String = "CP_01_001_add_03" '2014年啟用。'07:接受企業委託訓練
    Const cst_printFN5add As String = "CP_01_001_add_06" '2017年啟用。(在職)'07:接受企業委託訓練
    Const cst_printFN6add As String = "CP_01_001_add_19" '2019年啟用。70:區域產業據點職業訓練計畫(在職)
    'aspx程式
    Const cst_prg_addaspx1 As String = "CP_01_001_add.aspx"
    'Const cst_prg_addaspx2 As String = "CP_01_001_add_97.aspx"
    Const cst_prg_addaspx3 As String = "CP_01_001_add_98.aspx"
    Const cst_prg_addaspx4 As String = "CP_01_001_add_99.aspx"
    'Const cst_prg_addaspx5 As String = "CP_01_001_add_03.aspx" '07:接受企業委託訓練
    Const cst_prg_addaspx6 As String = "CP_01_001_add_06.aspx" '(在職)'07:接受企業委託訓練
    Const cst_prg_addaspx7 As String = "CP_01_001_add_19.aspx" '2019年啟用。70:區域產業據點職業訓練計畫(在職)

    Const cst_othValue_v1 As String = "&view=1"
    Const cst_othValue_add As String = "add"

    'Public Shared gflagTest1 As Boolean = False 'TIMS.sUtl_ChkTest() '測試環境參數

    Const cst_SearchStr As String = "SearchStr"
    'Dim FunDr As DataRow
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

#Region "CONST 1-IMP"

    Const cst_aOCID As Integer = 0          '班別(OCID)
    'Const cst_aSEQNO As Integer = 2
    Const cst_aAPPLYDATE As Integer = 1     '訪查日期	日期格式
    Const cst_aAPPLYDATEHH1 As Integer = 2  '訪查日期開始時	數字
    Const cst_aAPPLYDATEMI1 As Integer = 3  '訪查日期開始分	數字
    Const cst_aAPPLYDATEHH2 As Integer = 4  '訪查日期結束時	數字
    Const cst_aAPPLYDATEMI2 As Integer = 5  '訪查日期結束分	數字
    Const cst_aVISTIMES As Integer = 6      '第n次訪問	數字	1~n　

    Const cst_aAPPROVEDCOUNT As Integer = 7 '核定人數	數字	1~n
    Const cst_aAUTHCOUNT As Integer = 8     '開訓人數	數字	1~n
    Const cst_aTURTHCOUNT As Integer = 9    '實到人數	數字	0~n
    Const cst_aTURNOUTCOUNT As Integer = 10 '請假人數	數字	0~n
    Const cst_aTRUANCYCOUNT As Integer = 11 '缺(曠)課人數	數字	0~n
    Const cst_aLEAVECOUNT As Integer = 12   '離訓人數	數字	0~n
    Const cst_aREJECTCOUNT As Integer = 13  '退訓人數	數字	0~n
    'Const cst_aADVJOBCOUNT As Integer = 13 '提前就業人數	數字	0~n

    Const cst_aDATA1 As Integer = 14 '書面資料1	數字	1.備齊2.未備3.部份備有4.免提供
    Const cst_aDATACOPY1 As Integer = 15 '書面資料1如附件	字串長度50	
    Const cst_aD1CMM As Integer = 16 '書面資料1備齊月	字串2	　
    Const cst_aD1CDD As Integer = 17 '書面資料1備齊日	字串9	

    Const cst_aDATA2 As Integer = 18 '書面資料2	數字	1.備齊2.未備3.部份備有4.免提供
    Const cst_aDATACOPY2 As Integer = 19 '書面資料2如附件	字串長度50	　
    Const cst_aD2CMM As Integer = 20 '書面資料2備齊月	字串2
    Const cst_aD2CDD As Integer = 21    '書面資料2備齊日	字串9	

    Const cst_aDATA3 As Integer = 22    '書面資料3	數字	1.備齊2.未備3.部份備有4.免提供
    Const cst_aDATACOPY3 As Integer = 23 '書面資料3如附件	字串長度50	　
    Const cst_aD3C As Integer = 24      '書面資料3備註選項	數字	1.攜回2.免提供3.其他(請說明)
    Const cst_aD3CMM As Integer = 25    '書面資料3 攜回月	字串2	
    Const cst_aD3CDD As Integer = 26    '書面資料3 攜回日	字串9	　
    Const cst_aD3NOTE As Integer = 27   '書面資料3 其他說明	字串長度100	　

    'Const cst_aDATA4 As Integer = 28    '書面資料4	數字	1.備齊2.未備3.部份備有4.免提供
    'Const cst_aDATACOPY4 As Integer = 29 '書面資料4如附件	字串長度50	　
    'Const cst_aD4C As Integer = 30      '書面資料4選項	數字	1.攜回2.免提供3.免提供4.其他(請說明)　
    'Const cst_aD4NOTE As Integer = 31   '書面資料4其它說明	字串長度100	
    'Const cst_aDATA5 As Integer = 32    '書面資料5	數字	1.備齊2.未備3.部份備有4.免提供
    'Const cst_aDATACOPY5 As Integer = 33 '書面資料5如附件	字串長度50	　
    'Const cst_aD5C As Integer = 34      '書面資料5選項	數字	1.攜回2.免提供3.免提供4.免提供
    'Const cst_aD5NOTE As Integer = 35   '書面資料5說明	字串長度100	　
    'Const cst_aDATA6 As Integer = 36    '書面資料6	數字	1.備齊2.未備3.部份備有4.免提供
    'Const cst_aDATACOPY6 As Integer = 37 '書面資料6如附件	字串長度50	　
    'Const cst_aD6C As Integer = 38      '書面資料6選項	數字	1攜回影本2.免提供
    'Const cst_aDATA62 As Integer = 39   '書面資料7	數字	1.備齊2.未備3.部份備有4.免提供
    'Const cst_aDATACOPY62 As Integer = 40 '書面資料7如附件	字串長度50	　
    'Const cst_aD62C As Integer = 41     '書面資料7選項	數字	1攜回影本

    Const cst_aITEM1_1 As Integer = 28 '42  '課程(師資)實施狀況 1. 選項	數字	1有2無 3 免填
    Const cst_aITEM1_2 As Integer = 29 '43  '課程(師資)實施狀況 2. 選項	數字	1是2否 3 免填
    Const cst_aITEM1_COUR As Integer = 30 '44 '課程(師資)實施狀況 3. 課目	字串長度100	
    Const cst_aITEM1_3 As Integer = 31 '45  '課程(師資)實施狀況 教師與助教. 選項	數字	1是2否 3 免填　
    Const cst_aITEM1_TEACHER As Integer = 32 ' 46 '課程(師資)實施狀況 教師與助教. 教師：	字串長度100	　
    Const cst_aITEM1_ASSISTANT As Integer = 33 '47 '課程(師資)實施狀況 教師與助教. 助教：	字串長度100	　

    Const cst_aITEM1PROS As Integer = 34 ' 48 '課程(師資)實施狀況 處理情形	字串長度500	　
    Const cst_aITEM1NOTE As Integer = 35 '49 '課程(師資)實施狀況 備註	字串長度500	　
    Const cst_aITEM2_1 As Integer = 36 '50 '1.有無書籍(講義)領用表?	數字	1是2否 3 免填
    Const cst_aITEM2_2 As Integer = 37 '51 '2.有無材料領用表?	數字	1是2否 3 免填
    'Const cst_aITEM2_3 As Integer = 52 '3.訓練設施設備是否依契約提供學員使用?	數字	1是2否 3 免填
    Const cst_aITEM2NOTE As Integer = 38 '52 '53 '教材設施運用狀況 處理情形	字串長度500	　
    Const cst_aITEM2PROS As Integer = 39 '53 '54 '教材設施運用狀況 備註	字串長度500	

    Const cst_aITEM3_1 As Integer = 40 '54 '55 '1.教學(訓練)日誌是否確實填寫?	數字	1是2否 3 免填
    Const cst_aITEM3_2 As Integer = 41 '55 '56 '2.有否按時呈主管核閱?	數字	1是2否 3 免填
    'Const cst_aITEM3_3 As Integer = 57 '3.學員生活、就業輔導與管理機制是否依契約挸範辦理?	數字	1是2否 3 免填
    'Const cst_aITEM3_4 As Integer = 58 '4.是否依契約規範提供學員問題反應申訴管道?	數字	1是2否 3 免填
    'Const cst_aITEM3_5 As Integer = 59 '5.是否依契約規範公告學員權益義務管理狀況義務或編製參訓學員服務手冊?	數字	1是2否 3 免填
    Const cst_aITEM3PROS As Integer = 42 '56 '60 '教務管理狀況 處理情形	字串長度500	　
    Const cst_aITEM3NOTE As Integer = 43 '57 '61 '教務管理狀況 備註	字串長度500	

    'Const cst_aITEM4_1 As Integer = 62 '1.是否依規定於開訓後15日內收齊職業訓練生活津貼申請書及相關證明文件後送委訓單位審查？	數字	1是2否 3 免填
    'Const cst_aITEM4_2 As Integer = 63 '2.培訓單位於收到本署所屬分署核撥之津貼後，是否按月即時（不超過3個工作日）轉發給受訓學員。	數字	1是2否 3 免填　
    'Const cst_aITEM4_3 As Integer = 64 '3.申請人離、退訓時，培訓單位是否按月覈實繳回職業訓練生活津貼。	數字	1是2否 3 免填
    'Const cst_aITEM4NOTE As Integer = 65 '免填原因說明	字串長度100
    'Const cst_aITEM4PROS As Integer = 66 '費用(津貼)收核狀況 處理情形	字串長度500	
    Const cst_aITEM7NOTE As Integer = 44 '58 '67 '訓學員反映意見及問題	字串長度500	
    Const cst_aITEM7NOTE2 As Integer = 45 '59 '68 '學員反映意見之委訓單位反應說明	字串長度500	
    Const cst_aITEM31NOTE As Integer = 46 '60 '69 '綜合建議	字串長度500	　
    Const cst_aITEM32 As Integer = 47 '61 '70 '缺失處理	數字	4無缺失 1限期改善，研提檢討報告 2擇期進行訪查 3其他(請說明)
    Const cst_aITEM32NOTE As Integer = 48 '62 '71 '缺失處理 其他說明內容	字串長度500	　

    'Const cst_aCURSENAME As Integer = 72
    'Const cst_aVISITORNAME As Integer = 73
    Const cst_aVISITORNAME As Integer = 49 ' 63 '72 '訪查人員	字串長度10	　
    Const cst_iMaxLength1 As Integer = 50 '64 '73 '總欄位數

    'Const cst_aD1C As Integer = 21
    'Const cst_aD2C As Integer = 22
    'Const cst_aD1NOTE As Integer = 33
    'Const cst_aD2NOTE As Integer = 34
    'Const cst_aD6NOTE As Integer = 38
    'Const cst_aRID As Integer = 77
    'Const cst_aMODIFYACCT As Integer = 78
    'Const cst_aMODIFYDATE As Integer = 79

#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        'gflagTest1 = TIMS.sUtl_ChkTest() 'TIMS.sUtl_ChkTest() '測試環境參數
        '檢查帳號的功能權限-----------------------------------Start
        'Button2.Enabled = False
        'If blnCanAdds Then Button2.Enabled = True
        'Button1.Enabled = False
        'If blnCanSech Then Button1.Enabled = True
        '檢查帳號的功能權限-----------------------------------End

        rqDOCID = TIMS.ClearSQM(Request("DOCID"))

        If Not IsPostBack Then
            cCreate1()

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button5.Attributes("onclick") = s_javascript_btn2
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub cCreate1()

        Const cst_s_msg2 As String = "* 配合103年度訪查作業規範修訂，訪查次數更新如下：<br />自辦訓練，分署訪查培訓單位：<br />(1)訓練期程在450小時以下之班次，訪查1次。<br />(2)訓練期程在451至900小時之班次，訪查2次。<br />(3)訓練期程在901小時以上之班次，訪查3次。<br />"
        Labmsg2.Text = ""
        Select Case sm.UserInfo.TPlanID
            Case TIMS.Cst_TPlanID70, TIMS.Cst_TPlanID06, TIMS.Cst_TPlanID07
            Case Else
                Labmsg2.Text = cst_s_msg2
        End Select

        msg.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        end_date.Text = TIMS.Cdate3(Now.Date)
        DataGridTable.Visible = False
        Button1.Attributes("onclick") = "javascript:return search()"

        If Session(cst_SearchStr) IsNot Nothing Then
            center.Text = TIMS.GetMyValue(Session(cst_SearchStr), "center")
            RIDValue.Value = TIMS.GetMyValue(Session(cst_SearchStr), "RIDValue")
            TMID1.Text = TIMS.GetMyValue(Session(cst_SearchStr), "TMID1")
            OCID1.Text = TIMS.GetMyValue(Session(cst_SearchStr), "OCID1")
            TMIDValue1.Value = TIMS.GetMyValue(Session(cst_SearchStr), "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(Session(cst_SearchStr), "OCIDValue1")
            start_date.Text = TIMS.GetMyValue(Session(cst_SearchStr), "start_date")
            end_date.Text = TIMS.GetMyValue(Session(cst_SearchStr), "end_date")
            PageControler1.PageIndex = 0
            'PageControler1.PageIndex = TIMS.GetMyValue(Session(cst_SearchStr), "PageIndex")
            Dim MyValue As String = TIMS.GetMyValue(Session(cst_SearchStr), "PageIndex")
            If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                MyValue = CInt(MyValue)
                PageControler1.PageIndex = MyValue
            End If
            If TIMS.GetMyValue(Session(cst_SearchStr), "Button1") = "true" Then
                'Button1_Click(sender, e)
                Call sUtl_Search1(1)
            End If
            Session(cst_SearchStr) = Nothing
        End If

        Button7.Visible = False

        If rqDOCID <> "" Then
            '新興資軟查核
            Dim dr As DataRow = TIMS.GetOCIDDate(rqDOCID)
            If dr IsNot Nothing Then
                TMID1.Text = "[" & dr("TrainID").ToString & "]" & dr("TrainName").ToString
                TMIDValue1.Value = dr("TMID").ToString
                OCID1.Text = dr("CLASSCNAME2").ToString
                OCIDValue1.Value = dr("OCID").ToString
                center.Text = dr("OrgName")
                RIDValue.Value = dr("RID")
                Button5.Disabled = True
                Button6.Disabled = True
                '新興資軟查核
                If Not Session("_SearchStr") Is Nothing Then
                    Me.ViewState("_SearchStr") = Session("_SearchStr")
                    Session("_SearchStr") = Nothing
                End If
                Button7.Visible = True
            End If
        End If
        Call SHOW_HyperlinkA()
    End Sub

    '範例檔案下載(ZIP)
    Sub SHOW_HyperlinkA()
        Const cst_NavigateUrl1 As String = "../../Doc/ClassVisitor.zip"
        Const cst_NavigateUrl2 As String = "../../Doc/ClassVisitor2009.zip"
        Const cst_NavigateUrl3 As String = "../../Doc/ClassVisitor2010.zip"
        Const cst_NavigateUrl4 As String = "../../Doc/ClassVisitor2016.zip"
        'Const cst_NavigateUrl21 As String = "../../Doc/ClassVisitor_v21.zip"
        Const cst_NavigateUrl21b As String = "../../Doc/ClassVisitor_v21b.zip"

        '目前最新
        Hyperlink1.Visible = False
        Hyperlink1.NavigateUrl = ""
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        '2008年之前。
        If sm.UserInfo.Years <= 2008 Then Hyperlink1.NavigateUrl = cst_NavigateUrl1
        '2009年
        If sm.UserInfo.Years = 2009 Then Hyperlink1.NavigateUrl = cst_NavigateUrl2
        '2010年~'2015年
        If sm.UserInfo.Years >= 2010 And sm.UserInfo.Years <= 2015 Then Hyperlink1.NavigateUrl = cst_NavigateUrl3
        '2016年
        If sm.UserInfo.Years >= 2016 Then Hyperlink1.NavigateUrl = cst_NavigateUrl4
        '2021年
        If sm.UserInfo.Years >= 2021 Then Hyperlink1.NavigateUrl = cst_NavigateUrl21b

        If Hyperlink1.NavigateUrl <> "" Then Hyperlink1.Visible = True
    End Sub

    '此班級尚未開訓,無法新增! 'True:已經過了開訓時 False:尚未到達開訓時間。
    Function chkocid() As Boolean
        Dim rst As Boolean = False
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drOC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drOC Is Nothing Then Return rst
        HidThours.Value = Val(drOC("Thours"))
        HIDVISCOUNT.Value = GetVISTIMES(OCIDValue1.Value, objconn)
        If drOC("STDate") <= drOC("Today1") Then rst = True
        Return rst
    End Function

    Public Shared Sub UTL_DEL_CLS_VISITOR(ByRef MyPage As Page, ByRef oConn As SqlConnection, ByRef sCmdArg As String)
        If sCmdArg = "" Then Return
        Dim sVER As String = TIMS.GetMyValue(sCmdArg, "VER")
        Dim sOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim sSEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
        If sVER = "" Then Return
        If sOCID = "" Then Return
        If sSEQNO = "" Then Return

        Dim parms As Hashtable = New Hashtable()
        Select Case sVER
            Case "1"
                Dim sql As String = ""
                sql = ""
                sql &= " SELECT * FROM CLASS_VISITOR WHERE OCID = @OCID AND SEQNO = @SEQNO "
                parms.Clear()
                parms.Add("OCID", sOCID)
                parms.Add("SEQNO", sSEQNO)
                Dim dtVT As DataTable
                dtVT = DbAccess.GetDataTable(sql, oConn, parms)
                '使用欄位 '計算欄位
                Dim s_ColName As String = ""
                For i As Integer = 0 To dtVT.Columns.Count - 1 '所有欄位
                    If s_ColName <> "" Then s_ColName &= ","
                    s_ColName &= dtVT.Columns(i).ColumnName
                Next

                'Dim sql As String = ""
                sql = ""
                sql &= " INSERT INTO CLASS_VISITORDELDATA (" & s_ColName & ")"
                sql &= " SELECT " & s_ColName & " FROM CLASS_VISITOR WHERE OCID = @OCID AND SEQNO = @SEQNO"
                parms.Clear()
                parms.Add("OCID", sOCID)
                parms.Add("SEQNO", sSEQNO)
                DbAccess.ExecuteNonQuery(sql, oConn, parms)

                sql = " DELETE CLASS_VISITOR WHERE OCID=@OCID AND SEQNO=@SEQNO "
                parms.Clear()
                parms.Add("OCID", sOCID)
                parms.Add("SEQNO", sSEQNO)
                DbAccess.ExecuteNonQuery(sql, oConn, parms)
                Common.MessageBox(MyPage, "刪除成功!!")

            Case "3"
                Dim sql As String = ""
                sql = ""
                sql &= " SELECT * FROM CLASS_VISITOR3 WHERE OCID=@OCID AND SEQNO = @SEQNO"
                parms.Clear()
                parms.Add("OCID", sOCID)
                parms.Add("SEQNO", sSEQNO)
                Dim dtVT As DataTable
                dtVT = DbAccess.GetDataTable(sql, oConn, parms)
                '使用欄位 '計算欄位
                Dim s_ColName As String = ""
                For i As Integer = 0 To dtVT.Columns.Count - 1 '所有欄位
                    If s_ColName <> "" Then s_ColName &= ","
                    s_ColName &= dtVT.Columns(i).ColumnName
                Next

                'Dim sql As String = ""
                sql = ""
                sql &= " INSERT INTO CLASS_VISITOR3DELDATA (" & s_ColName & ")"
                sql &= " SELECT " & s_ColName & " FROM CLASS_VISITOR3 WHERE OCID=@OCID AND SEQNO = @SEQNO"
                parms.Clear()
                parms.Add("OCID", sOCID)
                parms.Add("SEQNO", sSEQNO)
                DbAccess.ExecuteNonQuery(sql, oConn, parms)

                sql = " DELETE CLASS_VISITOR3 WHERE OCID = @OCID AND SEQNO = @SEQNO"
                parms.Clear()
                parms.Add("OCID", sOCID)
                parms.Add("SEQNO", sSEQNO)
                DbAccess.ExecuteNonQuery(sql, oConn, parms)
                Common.MessageBox(MyPage, "刪除成功!!")
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                Dim rVER As String = TIMS.GetMyValue(e.CommandArgument, "VER")
                Dim rYears As String = TIMS.GetMyValue(e.CommandArgument, "Years")
                Dim rOCID As String = TIMS.GetMyValue(e.CommandArgument, "OCID")
                Dim rSEQNO As String = TIMS.GetMyValue(e.CommandArgument, "SEQNO")
                Dim rDOCID As String = TIMS.GetMyValue(e.CommandArgument, "DOCID")
                Dim sACT As String = TIMS.GetMyValue(e.CommandArgument, "ACT")
                Dim othValue As String = ""
                Dim sUrl As String = ""
                If sACT = "VIEW" Then '動作改為供查詢
                    othValue = cst_othValue_v1 '"&view=1"
                    sUrl = Get_sUrl_add(rVER, rYears, rOCID, rSEQNO, rDOCID, othValue)
                Else
                    sUrl = Get_sUrl_add(rVER, rYears, rOCID, rSEQNO, rDOCID, othValue)
                End If
                Dim dr As DataRow = TIMS.GetOCIDDate(rOCID, objconn)
                HidThours.Value = Val(dr("Thours"))
                HIDVISCOUNT.Value = GetVISTIMES(rOCID, objconn)
                Call GetSearchStr()
                '新興資軟查核
                Session("_SearchStr") = Me.ViewState("_SearchStr")
                'Response.Redirect(sUrl)
                Call TIMS.Utl_Redirect(Me, objconn, sUrl)

            Case "del"
                Dim sCmdArg As String = e.CommandArgument
                Call UTL_DEL_CLS_VISITOR(Me, objconn, sCmdArg)
                'Button1_Click(Button1, e)
                Call sUtl_Search1(1)

            Case "prt"
                '使用JAVASCRIPT.
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim LabITEM32 As Label = e.Item.FindControl("LabITEM32") '訪查結果
                Dim lItem31Note As Label = e.Item.FindControl("lItem31Note") '綜合建議

                e.Item.Cells(cst_序號).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim s_CBLITEM32 As String = Convert.ToString(drv("ITEM32"))
                If Convert.ToString(drv("CBL_ITEM32")) <> "" Then s_CBLITEM32 = Convert.ToString(drv("CBL_ITEM32"))
                LabITEM32.Text = s_CBLITEM32 '訪查結果
                'lItem31Note.Text = Convert.ToString(drv("ITEM31NOTE")) '綜合建議 ITEM31NOTE_30
                lItem31Note.Text = Convert.ToString(drv("ITEM31NOTE_30")) '綜合建議 ITEM31NOTE_30

                Dim ParentName As String = TIMS.Get_ParentRID(drv("Relship"), objconn)
                If ParentName <> "" Then e.Item.Cells(cst_訓練機構).Text = ParentName & "-" & Convert.ToString(drv("OrgName"))
                Dim mybut1 As Button = e.Item.FindControl("Button3") '修改edit
                Dim mybut2 As Button = e.Item.FindControl("Button4") '刪除del
                Dim mybut3 As Button = e.Item.FindControl("Button8") '列印prt
                'Dim mybut4 As Button = e.Item.Cells(3).FindControl("Button11")
                If Not IsDBNull(drv("RID")) Then
                    If drv("RID") <> sm.UserInfo.RID Then
                        mybut1.Text = "檢視"
                        mybut2.Visible = False
                    End If
                End If

                If IsDBNull(drv("applydate")) Then
                    mybut1.Visible = False
                    mybut2.Visible = False
                    mybut3.Visible = False
                End If

                'Dim rDOCID As String = ""  '有可能 DOCID 為空
                Dim rVER As String = Convert.ToString(drv("VER"))
                Dim rYears As String = Convert.ToString(drv("Years"))
                Dim rOCID As String = Convert.ToString(drv("OCID"))
                Dim rSeqNo As String = Convert.ToString(drv("SeqNo"))
                Dim othValue As String = ""
                'If rqDOCID <> "" Then rDOCID = Convert.ToString(rqDOCID)

                Dim cmdArg As String = ""
                'cmdArg = ""
                Call TIMS.SetMyValue(cmdArg, "VER", drv("VER"))
                Call TIMS.SetMyValue(cmdArg, "Years", drv("Years"))
                Call TIMS.SetMyValue(cmdArg, "OCID", drv("OCID"))
                Call TIMS.SetMyValue(cmdArg, "SEQNO", drv("SEQNO"))
                If rqDOCID <> "" Then Call TIMS.SetMyValue(cmdArg, "DOCID", rqDOCID)
                '查詢/修改。
                If Convert.ToString(drv("RID")) <> sm.UserInfo.RID Then
                    Call TIMS.SetMyValue(cmdArg, "ACT", "VIEW") '動作改為供查詢
                    mybut1.Text = "查詢"
                End If
                mybut1.CommandArgument = cmdArg

                '刪除。
                mybut2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                mybut2.CommandArgument = cmdArg '"OCID='" & drv("OCID") & "' and SeqNo='" & drv("SeqNo") & "'"

                '2008/03/06 Ellen Add '列印
                Dim sUrl As String = Get_sUrl_Prt(rVER, rYears, rOCID, rSeqNo)
                If sUrl <> "" Then mybut3.Attributes("onclick") = sUrl 'Get_sUrl_Prt(rVER, rYears, rOCID, rSeqNo)
                If sUrl = "" Then mybut3.Visible = False '若沒有資訊，則不顯示。
        End Select
    End Sub

    '保留搜尋條件。
    Sub GetSearchStr()
        Dim str_SearchStr_1 As String = ""
        str_SearchStr_1 = "k=1"
        str_SearchStr_1 += "&center=" & center.Text
        str_SearchStr_1 += "&RIDValue=" & RIDValue.Value
        str_SearchStr_1 += "&TMID1=" & TMID1.Text
        str_SearchStr_1 += "&OCID1=" & OCID1.Text
        str_SearchStr_1 += "&TMIDValue1=" & TMIDValue1.Value
        str_SearchStr_1 += "&OCIDValue1=" & OCIDValue1.Value
        str_SearchStr_1 += "&start_date=" & start_date.Text
        str_SearchStr_1 += "&end_date=" & end_date.Text
        str_SearchStr_1 += "&VISCOUNT=" & HIDVISCOUNT.Value '新增或修改都會傳遞此值。
        str_SearchStr_1 += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1

        str_SearchStr_1 += If(DataGrid1.Visible, "&Button1=true", "&Button1=false")
        Session(cst_SearchStr) = str_SearchStr_1
    End Sub

    Function sUtl_Search1_dt(ByRef sType As Integer) As DataTable
        Dim dt As DataTable = Nothing

        'stype:1:一般搜尋 2:匯出:本署訪查資料鈕 3:匯出:縣市政府訪查資料鈕。
        Dim Relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)
        If Relship = "" Then Return dt ' Exit Sub

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT b.OCID " & vbCrLf
        'VIEW_VISITOR(CLASS_VISITOR) 1:舊版 3:目前新版2014
        sql &= " ,ISNULL(a.VER,3) VER ,a.SeqNo ,a.RID ,CONVERT(VARCHAR, a.APPLYDATE, 111) ApplyDate " & vbCrLf
        sql &= " ,CASE a.ITEM32 WHEN 1 THEN '限期改善，研提檢討報告' " & vbCrLf
        sql &= "  WHEN 2 THEN '擇期進行訪查' " & vbCrLf
        sql &= "  WHEN 3 THEN '其他' " & vbCrLf
        sql &= "  WHEN 4 THEN '無缺失' END ITEM32 " & vbCrLf
        sql &= " ,dbo.FN_REPX(1,a.CBL_ITEM32) CBL_ITEM32 " & vbCrLf
        sql &= " ,a.ITEM31NOTE " & vbCrLf
        sql &= " ,case when len(a.ITEM31NOTE)>31 then concat(SUBSTRING(a.ITEM31NOTE,1,30),'...') else a.ITEM31NOTE end ITEM31NOTE_30" & vbCrLf
        sql &= " ,b.RID RIDValue " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,b.CLASSCNAME" & vbCrLf
        sql &= " ,b.CYCLTYPE" & vbCrLf
        sql &= " ,FORMAT(b.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,FORMAT(b.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,b.CyclType ,d.OrgName " & vbCrLf
        'sql &= " ,c.ClassID ClassID " & vbCrLf
        sql &= " ,d.Relship ,ip.Years " & vbCrLf
        sql &= " ,e.OrgName COrgName " & vbCrLf
        '訪查人員 VISITORNAME
        sql &= " ,a.VISITORNAME" & vbCrLf
        sql &= " ,concat(ISNULL(C2.APPLYDATEHH1,'　'),'時',ISNULL(C2.APPLYDATEMI1,'　'),'分')" & vbCrLf
        sql &= " +concat('至',ISNULL(C2.APPLYDATEHH2,'　'),'時',ISNULL(C2.APPLYDATEMI2,'　'),'分') APPLYTIME" & vbCrLf

        sql &= " FROM CLASS_CLASSINFO b " & vbCrLf
        sql &= " JOIN ID_CLASS c ON b.CLSID = c.CLSID " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PlanID = b.PlanID " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID ='" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            '無傳入值限定計畫
            If rqDOCID = "" Then sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        sql &= " JOIN dbo.VIEW_RIDNAME d ON b.RID = d.RID " & vbCrLf
        'sql += "JOIN org_orginfo oo ON oo.comidno = b.comidno " & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_VISITOR a ON a.OCID = b.OCID " & vbCrLf
        sql &= " LEFT JOIN dbo.CLASS_VISITOR3 C2 ON C2.OCID=a.OCID AND C2.SEQNO=a.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_RIDNAME e ON a.RID = e.RID " & vbCrLf
        sql &= " WHERE 1=1 "
        '排除不開班的班級  NotOpen='N'
        sql &= " AND b.NotOpen = 'N' " & vbCrLf
        sql &= " AND b.RID IN (SELECT RID FROM Auth_Relship WHERE Relship LIKE '" & Relship & "%') " & vbCrLf

        '班級
        If OCIDValue1.Value <> "" Then sql &= " AND b.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
        '判斷抽訪狀況
        Dim v_interview As String = TIMS.GetListValue(interview)
        Select Case v_interview 'interview.SelectedValue
            Case "1" '全部
                '抽訪區間
                If start_date.Text <> "" Then sql &= " AND a.ApplyDate >= " & TIMS.To_date(start_date.Text) & vbCrLf
                If end_date.Text <> "" Then sql &= " AND a.ApplyDate <= " & TIMS.To_date(end_date.Text) & vbCrLf '" & end_date.Text & "'" & vbCrLf
            Case "2" '有抽訪
                '抽訪區間
                If start_date.Text <> "" Then sql &= " AND a.ApplyDate >= " & TIMS.To_date(start_date.Text) & vbCrLf
                If end_date.Text <> "" Then sql &= " AND a.ApplyDate <= " & TIMS.To_date(end_date.Text) & vbCrLf '" & end_date.Text & "'" & vbCrLf
                sql &= " AND a.APPLYDATE IS NOT NULL" & vbCrLf
            Case "3" '未抽訪
                If start_date.Text <> "" Then sql &= " AND b.STDATE >= " & TIMS.To_date(start_date.Text) & vbCrLf
                If end_date.Text <> "" Then sql &= " AND b.STDATE <= " & TIMS.To_date(end_date.Text) & vbCrLf '" & end_date.Text & "'" & vbCrLf
                sql &= " AND a.APPLYDATE IS NULL" & vbCrLf
        End Select
        'stype:1:一般搜尋 2:匯出:本署訪查資料鈕 3:匯出:縣市政府訪查資料鈕。
        Select Case sType
            Case 1
                sql &= " ORDER BY a.RID, b.OCID " & vbCrLf
            Case 2
                sql &= " ORDER BY d.OrgName, a.RID, b.OCID " & vbCrLf
                'Case 3 'sql &= " ORDER BY e.OrgName, a.RID, b.OCID " & vbCrLf
        End Select

        dt = DbAccess.GetDataTable(sql, objconn)
        Return dt
    End Function

    '查詢SQL
    Sub sUtl_Search1(ByRef sType As Integer)
        Dim dt As DataTable = Nothing

        dt = sUtl_Search1_dt(sType)

        Select Case sType
            Case 1
                msg.Text = "查無資料!"
                DataGridTable.Visible = False
                If dt.Rows.Count > 0 Then
                    msg.Text = ""
                    DataGridTable.Visible = True
                    PageControler1.PageDataTable = dt
                    PageControler1.Sort = "RIDValue ,CyclType"
                    PageControler1.ControlerLoad()
                End If
            Case Else
                'stype:1:一般搜尋 2:匯出:本署訪查資料鈕 3:匯出:縣市政府訪查資料鈕。
                msg.Text = ""
                DataGridTable.Visible = True
                DataGrid1.DataSource = dt
                DataGrid1.DataBind()
        End Select
    End Sub

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sUtl_Search1(1)
    End Sub

    '新增鈕
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If OCID1.Text = "" Then
            Common.MessageBox(Me, "請選擇班級!!")
            Exit Sub
        End If
        If Not chkocid() Then
            Common.MessageBox(Me, "此班級尚未開訓,無法新增!")
            Exit Sub
        End If

        Call GetSearchStr()
        Dim sUrl As String
        'Dim rDOCID As String = ""  '有可能 DOCID 為空
        Dim rYears As String = Convert.ToString(sm.UserInfo.Years)
        Dim othValue As String = cst_othValue_add '"add" '新增。
        'If rqDOCID <> "" Then rDOCID = Convert.ToString(rqDOCID)

        sUrl = Get_sUrl_add("3", rYears, "", "", rqDOCID, "add")
        'Response.Redirect(sUrl)
        'Dim url1 As String = ""
        Call TIMS.Utl_Redirect(Me, objconn, sUrl)
    End Sub

    '列印空白實地訪查紀錄表鈕
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim TPlanID As String = sm.UserInfo.TPlanID
        Dim OrgID As String = sm.UserInfo.OrgID  '列印空白實地訪查紀錄表
        Dim Years As String = sm.UserInfo.Years
        Dim RID As String = sm.UserInfo.RID
        Dim OCID As String = OCIDValue1.Value
        If RIDValue.Value <> "" Then RID = RIDValue.Value
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇班級!!")
            Exit Sub
        End If
        'Dim sfilename As String = "CP_01_001_03" '最新報表。 '2014年啟用。
        Dim sMyValue As String = ""
        sMyValue = ""
        sMyValue &= "&OrgID=" & OrgID
        sMyValue &= "&Years=" & Years
        sMyValue &= "&RID=" & RID
        sMyValue &= "&OCID=" & OCID
        sMyValue &= "&TPlanID=" & TPlanID
        Call sUtl_prtiReport1(sMyValue) '列印空白表件。
    End Sub

    '回查核頁 (新興資軟查核)鈕
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr") '新興資軟查核
        Dim url1 As String = "../05/CP_05_001.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    ''' <summary>
    ''' 取得報表名稱。
    ''' </summary>
    ''' <param name="iType"></param>
    ''' <param name="ss"></param>
    ''' <returns></returns>
    Public Shared Function Get_sFileName1(ByVal iType As Integer, ByVal ss As String) As String
        Dim rVRE As String = TIMS.GetMyValue(ss, "rVRE")
        Dim rYears As String = TIMS.GetMyValue(ss, "rYears")
        Dim ssYears As String = TIMS.GetMyValue(ss, "ssYears")
        Dim ssTPlanID As String = TIMS.GetMyValue(ss, "ssTPlanID")
        Dim sfilename As String = ""
        Select Case iType 'iType: 1:空白表 /2:資料表
            Case 1
                '只有班級名稱的空白表
                'sfilename = cst_printFN4 '"CP_01_001_03" '最新報表。 '2014年啟用。
                If ssYears <= 2008 Then
                    sfilename = cst_printFN1 '"CP_01_001_97"  '2008年之前。
                ElseIf ssYears = 2009 Then
                    sfilename = cst_printFN2 '"CP_01_001_98"  '2009年
                ElseIf ssYears >= 2010 AndAlso ssYears <= 2013 Then
                    sfilename = cst_printFN3 '"CP_01_001_99"  '2010年啟用。
                End If
                If TIMS.Cst_TPlanID07.IndexOf(ssTPlanID) > -1 Then sfilename = cst_printFN5 '"CP_01_001_06"
                If TIMS.Cst_TPlanID06_CP01001.IndexOf(ssTPlanID) > -1 Then sfilename = cst_printFN5 '"CP_01_001_06"
                If TIMS.Cst_TPlanID70.IndexOf(ssTPlanID) > -1 Then sfilename = cst_printFN6 '"CP_01_001_19"

            Case 2
                '班級的資料表
                'sfilename = cst_printFN4add '"CP_01_001_add_03" '新Prt filename
                Select Case rVRE
                    Case "1" '沿用舊版
                        sfilename = cst_printFN3add '"CP_01_001_add_99" '最新Prt filename
                End Select
                'Dim sfilename As String = "CP_01_001_add_99" '最新Prt filename
                If rYears <= "2007" Then
                    sfilename = cst_printFN0add '"CP_01_001_Rpt"
                ElseIf rYears = "2008" Then
                    sfilename = cst_printFN1add '"CP_01_001_add_97"
                ElseIf rYears = "2009" Then
                    sfilename = cst_printFN2add '"CP_01_001_add_98"
                ElseIf rYears >= "2010" AndAlso rYears <= "2013" Then
                    sfilename = cst_printFN3add '"CP_01_001_add_99"
                End If
                If TIMS.Cst_TPlanID07.IndexOf(ssTPlanID) > -1 Then sfilename = cst_printFN5add '"CP_01_001_06"
                If TIMS.Cst_TPlanID06_CP01001.IndexOf(ssTPlanID) > -1 Then sfilename = cst_printFN5add '"CP_01_001_add_06"
                If TIMS.Cst_TPlanID70.IndexOf(ssTPlanID) > -1 Then sfilename = cst_printFN6add '"CP_01_001_add_19"

        End Select
        Return sfilename
    End Function

    ''' <summary>
    ''' 取得最新導向網頁。ASPX
    ''' </summary>
    ''' <param name="rVER"></param>
    ''' <param name="rYears"></param>
    ''' <param name="rOCID"></param>
    ''' <param name="rSeqNo"></param>
    ''' <param name="rDOCID"></param>
    ''' <param name="othValue"></param>
    ''' <returns></returns>
    Function Get_sUrl_add(ByVal rVER As String, ByVal rYears As String, ByVal rOCID As String, ByVal rSeqNo As String, ByVal rDOCID As String, ByVal othValue As String) As String
        Dim rst As String = ""
        Dim sPage As String = cst_prg_addaspx4 ' "CP_01_001_add_99.aspx" '舊Page 
        'Select Case rVER
        '    Case "3" '使用新版
        '        sPage = cst_prg_addaspx5 '"CP_01_001_add_03.aspx" '新Page  
        'End Select
        If rYears <= "2007" Then
            sPage = cst_prg_addaspx1 '"CP_01_001_add.aspx"
        ElseIf rYears = "2008" Then
            sPage = cst_prg_addaspx1 '"CP_01_001_add_97.aspx"
        ElseIf rYears = "2009" Then
            sPage = cst_prg_addaspx3 '"CP_01_001_add_98.aspx"
        ElseIf rYears >= "2010" AndAlso rYears <= "2013" Then
            sPage = cst_prg_addaspx4 '"CP_01_001_add_99.aspx"
        End If
        If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then sPage = cst_prg_addaspx6
        If TIMS.Cst_TPlanID06_CP01001.IndexOf(sm.UserInfo.TPlanID) > -1 Then sPage = cst_prg_addaspx6 '"CP_01_001_add_06.aspx"
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then sPage = cst_prg_addaspx7 '"CP_01_001_add_06.aspx"

        Select Case othValue '"add
            Case cst_othValue_add '"add" '新增
                '有可能 DOCID 為空
                rst = sPage & "?ID=" & Request("ID")
                If rDOCID <> "" Then rst &= "&DOCID=" & rDOCID
            Case Else '修改、其他。
                '有可能 DOCID 為空
                rst = sPage & "?ID=" & Request("ID") & "&OCID=" & rOCID & "&SeqNo=" & rSeqNo & othValue
                If rDOCID <> "" Then rst &= "&DOCID=" & rDOCID
        End Select
        Return rst
    End Function

    ''' <summary>
    ''' 取得最新導向Prt。JRXML
    ''' </summary>
    ''' <param name="rVRE"></param>
    ''' <param name="rYears"></param>
    ''' <param name="rOCID"></param>
    ''' <param name="rSeqNo"></param>
    ''' <returns></returns>
    Function Get_sUrl_Prt(ByVal rVRE As String, ByVal rYears As String, ByVal rOCID As String, ByVal rSeqNo As String) As String
        Dim rst As String = ""
        Dim ss As String = ""
        TIMS.SetMyValue(ss, "rVRE", rVRE)
        TIMS.SetMyValue(ss, "rYears", rYears)
        TIMS.SetMyValue(ss, "ssYears", Convert.ToString(sm.UserInfo.Years))
        TIMS.SetMyValue(ss, "ssTPlanID", Convert.ToString(sm.UserInfo.TPlanID))
        Dim sfilename As String = Get_sFileName1(2, ss)
        rst = ReportQuery.ReportScript(Me, sfilename, "OCID=" & rOCID & "&SeqNo=" & rSeqNo)
        Return rst
    End Function

    ''' <summary>
    ''' 列印空白表件。SmartQuery.PrintReport
    ''' </summary>
    ''' <param name="sMyValue"></param>
    Sub sUtl_prtiReport1(ByVal sMyValue As String)
        '報表名稱。
        Dim ss As String = ""
        TIMS.SetMyValue(ss, "rVRE", "") '無值
        TIMS.SetMyValue(ss, "rYears", "") '無值
        TIMS.SetMyValue(ss, "ssYears", Convert.ToString(sm.UserInfo.Years))
        TIMS.SetMyValue(ss, "ssTPlanID", Convert.ToString(sm.UserInfo.TPlanID))
        Dim sfilename As String = Get_sFileName1(1, ss)
        If sfilename = "" Then
            Common.MessageBox(Me, "該計畫年度，未設定報表!!")
            Exit Sub
        End If
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sfilename, sMyValue)
    End Sub

    '訪視次數至少N次，本次為第N次訪視
    Public Shared Function GetVISTIMES(ByVal OCID As String, ByRef objconn As SqlConnection) As Integer
        Dim rst As Integer = 0
        '* 配合103年度訪查作業規範修訂，訪查次數更新如下：
        '1:.自辦訓練, 分署訪查培訓單位
        '(1)訓練期程在450小時以下之班次，訪查1次。
        '(2)訓練期程在451至900小時之班次，訪查2次。
        '(3)訓練期程在901小時以上之班次，訪查3次。
        '2.委外(補助)訓練，分署訪查培訓單位：(包括委外職前訓練、偏遠地區原住民、中長期及新移民等特定對象失業者職前訓練、辦理照顧服務職類職業訓練、推動事業單位辦理職前培訓計畫)
        '(1)訓練期程在180小時以下之班次，訪查1次。
        '(2)訓練期程在181至540小時之班次，訪查2次。
        '(3)訓練期程在541小時以上之班次，訪查3次。
        '3.提升勞工基礎數位能力研習-實體班，分署訪查培訓單位：訓練期程為36小時之實體課程，每一訓練地點應至少訪查1次，且每年訪查訓練班次比率至少應達實際總開班課程班次之60%以上。
        '4:.辦理失業者職業訓練, 彙管單位訪查培訓單位
        '(1)訓練期程在180小時以下之班次，訪查1次。
        '(2)訓練期程在181至540小時之班次，訪查2次。
        '(3)訓練期程在541小時以上之班次，訪查3次。
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sql As String = "SELECT dbo.fn_GETVISTIMES(@OCID) TIMES "
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCID)
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = Val(dt.Rows(0)("TIMES"))
        Return rst
    End Function

#Region "EXP"

    '2:匯出:本署訪查資料鈕 
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        DataGrid1.Columns(cst_功能).Visible = False
        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim dtX1 As DataTable = sUtl_Search1_dt(2)

        Dim sPattern As String = "" '序號,
        Dim sColumn As String = ""
        sPattern = "訓練機構,班別名稱,期別,訪視單位,開訓日期,結訓日期,訪查日期,訪查時間,訪查人員,訪查結果,綜合建議"
        sColumn = "ORGNAME,CLASSCNAME,CYCLTYPE,CORGNAME,STDATE,FTDATE,APPLYDATE,APPLYTIME,VISITORNAME,ITEM32,ITEM31NOTE"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        'DataGrid1.AllowPaging = False '關閉分頁功能
        'DataGrid1.EnableViewState = False  '把ViewState給關了
        'Dim Swr As New System.IO.StringWriter
        'Dim Htw As New System.Web.UI.HtmlTextWriter(Swr)
        'Div1.RenderControl(Htw)
        'DataGrid1.Visible = False
        ''Common.RespWrite(Me, Convert.ToString(Swr))
        'Dim sFileName1 As String = "本署訪查資料"
        'Dim strHTML As String = ""
        'strHTML &= Convert.ToString(Swr)

        Dim sFileName1 As String = "本署訪查資料"
        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim strHTML As String = ""
        strHTML &= "<div>"
        strHTML &= "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
        'Common.RespWrite(Me, "<tr>")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr &= "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "序號") '& vbTab
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        '建立資料面
        Dim iNum As Integer = 0
        For Each dr As DataRow In dtX1.Rows
            iNum += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", iNum) '& vbTab
            For i As Integer = 0 To sColumnA.Length - 1
                ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.CloseDbConn(objconn)
        Response.End()
    End Sub

    ''3:匯出:縣市政府訪查資料鈕
    'Protected Sub btnExport2_Click(sender As Object, e As EventArgs) Handles btnExport2.Click
    '    DataGrid1.Columns(cst_功能).Visible = False
    '    DataGrid1.AllowPaging = False '關閉分頁功能
    '    DataGrid1.EnableViewState = False  '把ViewState給關了

    '    Call sUtl_Search1(3)

    '    'DataGrid1.AllowPaging = False '關閉分頁功能
    '    'DataGrid1.EnableViewState = False  '把ViewState給關了

    '    Dim objStringWriter As New System.IO.StringWriter
    '    Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
    '    Div1.RenderControl(objHtmlTextWriter)
    '    'Common.RespWrite(Me, Convert.ToString(objStringWriter))
    '    DataGrid1.Visible = False

    '    Dim sFileName1 As String = "縣市政府訪查資料"
    '    Dim strHTML As String = ""
    '    strHTML &= Convert.ToString(objStringWriter)

    '    Dim v_ExpType As String = TIMS.GetListValue(RBListExpType)
    '    Dim parmsExp As New Hashtable
    '    parmsExp.Add("ExpType", v_ExpType) 'EXCEL/PDF/ODS
    '    parmsExp.Add("FileName", sFileName1)
    '    'parmsExp.Add("strSTYLE", strSTYLE)
    '    parmsExp.Add("strHTML", strHTML)
    '    parmsExp.Add("ResponseNoEnd", "Y")
    '    TIMS.Utl_ExportRp1(Me, parmsExp)

    '    TIMS.CloseDbConn(objconn)
    '    Response.End()
    'End Sub

#End Region

#Region "sub_XLSImp3"

    ''' <summary>
    ''' 匯入名冊
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Btn_XlsImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_XlsImport.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim impType As String = ""
        'Const cst_imp1 As String = "imp1"
        'Const cst_imp3 As String = "imp3"
        'impType = cst_imp1
        'If sm.UserInfo.Years >= "2016" Then impType = cst_imp3
        'Select Case impType
        '    Case cst_imp1
        '        Call sub_XLSImp1()
        '    Case cst_imp3
        '        Call sub_XLSImp3() '2016匯入
        'End Select


        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "ods") Then Return
        End If

        Const Cst_FileSavePath As String = "~/CP/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        'Call sImport2(FullFileName1)
        Call Sub_XLSImp3(FullFileName1)

    End Sub

    ''' <summary>
    ''' 匯入名冊
    ''' </summary>
    ''' <param name="FullFileName1"></param>
    Sub Sub_XLSImp3(ByRef FullFileName1 As String)
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "FullFileName1", FullFileName1)
        TIMS.SetMyValue2(htSS, "FirstCol", "班別(OCID)") '"身分證字號" '任1欄位名稱(必填)
        Dim Reason As String = ""
        '上傳檔案/取得內容
        Dim dt_xls As DataTable = TIMS.Get_File1data(File1, Reason, htSS, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        '儲存錯誤資料的DataTable
        Dim dtWrong As New DataTable
        Dim drWrong As DataRow = Nothing
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("OCID"))
        dtWrong.Columns.Add(New DataColumn("VisitorName"))
        dtWrong.Columns.Add(New DataColumn("ApplyDate"))
        dtWrong.Columns.Add(New DataColumn("Reason"))

        Dim iRowIndex As Integer = 1
        Dim sReason As String = "" '儲存錯誤的原因
        Dim tConn As SqlConnection = DbAccess.GetConnection
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Dim iCmd As SqlCommand = GetxICmd(oTrans, tConn)
        Try
            '有資料
            sReason = ""
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                If iRowIndex <> 0 Then
                    Dim colArray As Array = dt_xls.Rows(i).ItemArray
                    sReason = CheckImportData3(colArray) '依據匯入檔判斷錯誤
                    If sReason = "" Then
                        Dim dr1ARY As DataRow = ChgImpData(colArray) 'out@dr1
                        If dr1ARY IsNot Nothing Then Call Savedata3(iCmd, dr1ARY) '無錯誤存檔 '匯入資料
                    End If

                    If sReason <> "" Then
                        '錯誤資料，填入錯誤資料表
                        drWrong = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)
                        drWrong("Index") = iRowIndex 'Index 第幾筆錯誤
                        If colArray.Length > cst_aOCID Then drWrong("OCID") = Convert.ToString(colArray(cst_aOCID)) '班級代碼
                        If colArray.Length > cst_aAPPLYDATE Then drWrong("ApplyDate") = TIMS.Cdate3(colArray(cst_aAPPLYDATE)) '抽訪日期
                        Dim s_VisitorName As String = "查無此欄位"
                        '訪查人員 VISITORNAME
                        If colArray.Length > cst_aVISITORNAME Then
                            s_VisitorName = If(Convert.ToString(colArray(cst_aVISITORNAME)) <> "", Convert.ToString(colArray(cst_aVISITORNAME)), "未填寫")
                        End If
                        drWrong("VisitorName") = s_VisitorName
                        drWrong("Reason") = sReason '原因
                    End If
                End If
                iRowIndex += 1
            Next
            'DbAccess.UpdateDataTable(dt, da, oTrans)
            DbAccess.CommitTrans(oTrans)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)

            DbAccess.RollbackTrans(oTrans)
            Call TIMS.CloseDbConn(tConn)
            Common.MessageBox(Me, "儲存失敗!!")
            'Common.MessageBox(Me, ex.ToString)
            Exit Sub
        End Try
        'tConn.BeginTransaction() 'DbAccess.GetConnection
        'Call TIMS.OpenDbConn(tConn)
        Call TIMS.CloseDbConn(tConn)
        '判斷匯出資料是否有誤
        'Dim explain, explain2 As String
        Dim explain As String = ""
        Dim explain2 As String = ""
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        If dtWrong.Rows.Count > 0 Then
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('CP_01_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            Exit Sub
        End If
        If sReason <> "" Then
            Common.MessageBox(Me, explain & sReason)
            Exit Sub
        End If
        Common.MessageBox(Me, explain)
        Exit Sub
    End Sub

    ''' <summary>
    ''' 取得一個iCmd
    ''' </summary>
    ''' <param name="oTrans"></param>
    ''' <param name="tConn"></param>
    ''' <returns></returns>
    Public Shared Function GetxICmd(ByRef oTrans As SqlTransaction, ByRef tConn As SqlConnection) As SqlCommand
        Call TIMS.OpenDbConn(tConn)
        Dim sql As String = ""
        sql = ""
        sql &= " INSERT INTO CLASS_VISITOR3 ( " & vbCrLf
        sql &= "   OCID " & vbCrLf
        sql &= "   ,SEQNO " & vbCrLf
        sql &= "   ,APPLYDATE " & vbCrLf
        sql &= "   ,APPLYDATEHH1 " & vbCrLf
        sql &= "   ,APPLYDATEMI1 " & vbCrLf
        sql &= "   ,APPLYDATEHH2 " & vbCrLf
        sql &= "   ,APPLYDATEMI2 " & vbCrLf
        sql &= "   ,VISTIMES " & vbCrLf
        sql &= "   ,DATA1 " & vbCrLf
        sql &= "   ,DATA2 " & vbCrLf
        sql &= "   ,DATA3 " & vbCrLf
        'sql &= "   ,DATA4 " & vbCrLf
        'sql &= "   ,DATA5 " & vbCrLf
        'sql &= "   ,DATA6 " & vbCrLf
        sql &= "   ,DATACOPY1 " & vbCrLf
        sql &= "   ,DATACOPY2 " & vbCrLf
        sql &= "   ,DATACOPY3 " & vbCrLf
        'sql &= "   ,DATACOPY4 " & vbCrLf
        'sql &= "   ,DATACOPY5 " & vbCrLf
        'sql &= "   ,DATACOPY6 " & vbCrLf
        'sql &= "  ,D1C " & vbCrLf
        'sql &= "  ,D2C " & vbCrLf
        sql &= "   ,D3C " & vbCrLf
        'sql &= "   ,D4C " & vbCrLf
        'sql &= "   ,D5C " & vbCrLf
        'sql &= "   ,D6C " & vbCrLf
        'sql &= "   ,DATA62 " & vbCrLf
        'sql &= "   ,DATACOPY62 " & vbCrLf
        'sql &= "   ,D62C " & vbCrLf
        sql &= "   ,ITEM7NOTE2 " & vbCrLf
        sql &= "   ,D1CMM " & vbCrLf
        sql &= "   ,D1CDD " & vbCrLf
        sql &= "   ,D2CMM " & vbCrLf
        sql &= "   ,D2CDD " & vbCrLf
        sql &= "   ,D3CMM " & vbCrLf
        sql &= "   ,D3CDD " & vbCrLf
        'sql &= "  ,D1NOTE " & vbCrLf
        'sql &= "  ,D2NOTE " & vbCrLf
        sql &= "   ,D3NOTE " & vbCrLf
        'sql &= "   ,D4NOTE " & vbCrLf
        'sql &= "  ,D5NOTE " & vbCrLf
        'sql &= "  ,D6NOTE " & vbCrLf
        sql &= "   ,APPROVEDCOUNT " & vbCrLf
        sql &= "   ,AUTHCOUNT " & vbCrLf
        sql &= "   ,TURTHCOUNT " & vbCrLf
        sql &= "   ,TURNOUTCOUNT " & vbCrLf
        sql &= "   ,TRUANCYCOUNT " & vbCrLf
        sql &= "   ,LEAVECOUNT " & vbCrLf
        sql &= "   ,REJECTCOUNT " & vbCrLf
        'sql &= "   ,ADVJOBCOUNT " & vbCrLf
        sql &= "   ,ITEM1_1 " & vbCrLf
        sql &= "   ,ITEM1_2 " & vbCrLf
        sql &= "   ,ITEM1_COUR " & vbCrLf
        sql &= "   ,ITEM1_3 " & vbCrLf
        sql &= "   ,ITEM1_TEACHER " & vbCrLf
        sql &= "   ,ITEM1_ASSISTANT " & vbCrLf
        sql &= "   ,ITEM2_1 " & vbCrLf
        sql &= "   ,ITEM2_2 " & vbCrLf
        'sql &= "   ,ITEM2_3 " & vbCrLf
        sql &= "   ,ITEM3_1 " & vbCrLf
        sql &= "   ,ITEM3_2 " & vbCrLf
        'sql &= "   ,ITEM3_3 " & vbCrLf
        'sql &= "   ,ITEM3_4 " & vbCrLf
        'sql &= "   ,ITEM3_5 " & vbCrLf
        'sql &= "   ,ITEM4_1 " & vbCrLf
        'sql &= "   ,ITEM4_2 " & vbCrLf
        'sql &= "   ,ITEM4_3 " & vbCrLf
        'sql &= "   ,ITEM4NOTE " & vbCrLf
        sql &= "   ,ITEM7NOTE " & vbCrLf
        sql &= "   ,ITEM31NOTE " & vbCrLf
        sql &= "   ,ITEM32 " & vbCrLf
        sql &= "   ,ITEM32NOTE " & vbCrLf
        sql &= "   ,ITEM1PROS " & vbCrLf
        sql &= "   ,ITEM2PROS " & vbCrLf
        sql &= "   ,ITEM3PROS " & vbCrLf
        'sql &= "   ,ITEM4PROS " & vbCrLf
        sql &= "   ,ITEM1NOTE " & vbCrLf
        sql &= "   ,ITEM2NOTE " & vbCrLf
        sql &= "   ,ITEM3NOTE " & vbCrLf
        'sql &= "   ,CURSENAME " & vbCrLf
        sql &= "   ,VISITORNAME" & vbCrLf
        sql &= "   ,RID" & vbCrLf
        sql &= "   ,MODIFYACCT" & vbCrLf
        sql &= "   ,MODIFYDATE" & vbCrLf
        sql &= " ) VALUES ( " & vbCrLf
        sql &= "   @OCID " & vbCrLf
        sql &= "   ,@SEQNO " & vbCrLf
        sql &= "   ,@APPLYDATE " & vbCrLf
        sql &= "   ,@APPLYDATEHH1 " & vbCrLf
        sql &= "   ,@APPLYDATEMI1 " & vbCrLf
        sql &= "   ,@APPLYDATEHH2 " & vbCrLf
        sql &= "   ,@APPLYDATEMI2 " & vbCrLf
        sql &= "   ,@VISTIMES " & vbCrLf
        sql &= "   ,@DATA1 " & vbCrLf
        sql &= "   ,@DATA2 " & vbCrLf
        sql &= "   ,@DATA3 " & vbCrLf
        'sql &= "   ,@DATA4 " & vbCrLf
        'sql &= "   ,@DATA5 " & vbCrLf
        'sql &= "   ,@DATA6 " & vbCrLf
        sql &= "   ,@DATACOPY1 " & vbCrLf
        sql &= "   ,@DATACOPY2 " & vbCrLf
        sql &= "   ,@DATACOPY3 " & vbCrLf
        'sql &= "   ,@DATACOPY4 " & vbCrLf
        'sql &= "   ,@DATACOPY5 " & vbCrLf
        'sql &= "   ,@DATACOPY6 " & vbCrLf
        'sql &= "  ,@D1C " & vbCrLf
        'sql &= "  ,@D2C " & vbCrLf
        sql &= "   ,@D3C " & vbCrLf
        'sql &= "   ,@D4C " & vbCrLf
        'sql &= "   ,@D5C " & vbCrLf
        'sql &= "   ,@D6C " & vbCrLf
        'sql &= "   ,@DATA62 " & vbCrLf
        'sql &= "   ,@DATACOPY62 " & vbCrLf
        'sql &= "   ,@D62C " & vbCrLf
        sql &= "   ,@ITEM7NOTE2 " & vbCrLf
        sql &= "   ,@D1CMM " & vbCrLf
        sql &= "   ,@D1CDD " & vbCrLf
        sql &= "   ,@D2CMM " & vbCrLf
        sql &= "   ,@D2CDD " & vbCrLf
        sql &= "   ,@D3CMM " & vbCrLf
        sql &= "   ,@D3CDD " & vbCrLf
        'sql &= "  ,@D1NOTE " & vbCrLf
        'sql &= "  ,@D2NOTE " & vbCrLf
        sql &= "   ,@D3NOTE " & vbCrLf
        'sql &= "   ,@D4NOTE " & vbCrLf
        'sql &= "  ,@D5NOTE " & vbCrLf
        'sql &= "  ,@D6NOTE " & vbCrLf
        sql &= "   ,@APPROVEDCOUNT " & vbCrLf
        sql &= "   ,@AUTHCOUNT " & vbCrLf
        sql &= "   ,@TURTHCOUNT " & vbCrLf
        sql &= "   ,@TURNOUTCOUNT " & vbCrLf
        sql &= "   ,@TRUANCYCOUNT " & vbCrLf
        sql &= "   ,@LEAVECOUNT" & vbCrLf
        sql &= "   ,@REJECTCOUNT " & vbCrLf
        'sql &= "   ,@ADVJOBCOUNT " & vbCrLf
        sql &= "   ,@ITEM1_1 " & vbCrLf
        sql &= "   ,@ITEM1_2 " & vbCrLf
        sql &= "   ,@ITEM1_COUR " & vbCrLf
        sql &= "   ,@ITEM1_3 " & vbCrLf
        sql &= "   ,@ITEM1_TEACHER " & vbCrLf
        sql &= "   ,@ITEM1_ASSISTANT " & vbCrLf
        sql &= "   ,@ITEM2_1 " & vbCrLf
        sql &= "   ,@ITEM2_2 " & vbCrLf
        'sql &= "   ,@ITEM2_3 " & vbCrLf
        sql &= "   ,@ITEM3_1 " & vbCrLf
        sql &= "   ,@ITEM3_2 " & vbCrLf
        'sql &= "   ,@ITEM3_3 " & vbCrLf
        'sql &= "   ,@ITEM3_4 " & vbCrLf
        'sql &= "   ,@ITEM3_5 " & vbCrLf
        'sql &= "   ,@ITEM4_1 " & vbCrLf
        'sql &= "   ,@ITEM4_2 " & vbCrLf
        'sql &= "   ,@ITEM4_3 " & vbCrLf
        'sql &= "   ,@ITEM4NOTE " & vbCrLf
        sql &= "   ,@ITEM7NOTE " & vbCrLf
        sql &= "   ,@ITEM31NOTE " & vbCrLf
        sql &= "   ,@ITEM32 " & vbCrLf
        sql &= "   ,@ITEM32NOTE " & vbCrLf
        sql &= "   ,@ITEM1PROS " & vbCrLf
        sql &= "   ,@ITEM2PROS " & vbCrLf
        sql &= "   ,@ITEM3PROS " & vbCrLf
        'sql &= "   ,@ITEM4PROS " & vbCrLf
        sql &= "   ,@ITEM1NOTE " & vbCrLf
        sql &= "   ,@ITEM2NOTE " & vbCrLf
        sql &= "   ,@ITEM3NOTE " & vbCrLf
        'sql &= "   ,@CURSENAME " & vbCrLf
        sql &= "   ,@VISITORNAME " & vbCrLf
        sql &= "   ,@RID " & vbCrLf
        sql &= "   ,@MODIFYACCT " & vbCrLf
        sql &= "   ,GETDATE() " & vbCrLf
        sql &= " ) " & vbCrLf
        Dim iCMD As New SqlCommand(sql, tConn, oTrans)
        Return iCMD
    End Function

    '新增iCmd
    Sub Savedata3(ByRef iCmd As SqlCommand, ByVal dr1ARY As DataRow)
        'Dim iSEQNO As Integer = 0
        'iSEQNO = sUtl_GetSEQNO()
        'Dim iD3C As Integer = 0
        'If D3C1.Checked Then iD3C = 1
        'If D3C2.Checked Then iD3C = 2
        'If D3C3.Checked Then iD3C = 3
        'Dim dt As New DataTable
        'Dim oCmd As New SqlCommand(sql, objconn)

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim oTrans As SqlTransaction = iCmd.Transaction
        Dim sql As String = ""
        sql = "SELECT dbo.NVL(MAX(SEQNO),0) SEQNO FROM CLASS_VISITOR3 WHERE OCID='" & dr1ARY("OCID") & "'"
        Dim drS As DataRow = DbAccess.GetOneRow(sql, oTrans)
        Dim iSeqNo As Integer = 0 '依班別得到最大序號(無資料為0)
        If Not drS Is Nothing Then iSeqNo = Val(drS("SeqNO"))
        iSeqNo += 1 '(本次新增加1)
        'dr1("SEQNO") = iSeqNo

        With iCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = dr1ARY("OCID")
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = iSeqNo 'dr1ARY("SEQNO")
            .Parameters.Add("APPLYDATE", SqlDbType.DateTime).Value = dr1ARY("APPLYDATE")
            .Parameters.Add("APPLYDATEHH1", SqlDbType.VarChar).Value = dr1ARY("APPLYDATEHH1")
            .Parameters.Add("APPLYDATEMI1", SqlDbType.VarChar).Value = dr1ARY("APPLYDATEMI1")
            .Parameters.Add("APPLYDATEHH2", SqlDbType.VarChar).Value = dr1ARY("APPLYDATEHH2")
            .Parameters.Add("APPLYDATEMI2", SqlDbType.VarChar).Value = dr1ARY("APPLYDATEMI2")
            .Parameters.Add("VISTIMES", SqlDbType.Int).Value = dr1ARY("VISTIMES")
            .Parameters.Add("DATA1", SqlDbType.VarChar).Value = dr1ARY("DATA1") 'IIf(DATA1.SelectedValue = "", Convert.DBNull, DATA1.SelectedValue)
            .Parameters.Add("DATA2", SqlDbType.VarChar).Value = dr1ARY("DATA2") 'IIf(DATA2.SelectedValue = "", Convert.DBNull, DATA2.SelectedValue)
            .Parameters.Add("DATA3", SqlDbType.VarChar).Value = dr1ARY("DATA3") 'IIf(DATA3.SelectedValue = "", Convert.DBNull, DATA3.SelectedValue)
            '.Parameters.Add("DATA4", SqlDbType.VarChar).Value = dr1ARY("DATA4") 'IIf(DATA4.SelectedValue = "", Convert.DBNull, DATA4.SelectedValue)
            '.Parameters.Add("DATA5", SqlDbType.VarChar).Value = dr1ARY("DATA5") 'IIf(DATA5.SelectedValue = "", Convert.DBNull, DATA5.SelectedValue)
            '.Parameters.Add("DATA6", SqlDbType.VarChar).Value = dr1ARY("DATA6") 'IIf(DATA6.SelectedValue = "", Convert.DBNull, DATA6.SelectedValue)
            .Parameters.Add("DATACOPY1", SqlDbType.VarChar).Value = dr1ARY("DATACOPY1") 'DATACOPY1.Text
            .Parameters.Add("DATACOPY2", SqlDbType.VarChar).Value = dr1ARY("DATACOPY2") 'DATACOPY2.Text
            .Parameters.Add("DATACOPY3", SqlDbType.VarChar).Value = dr1ARY("DATACOPY3") 'DATACOPY3.Text
            '.Parameters.Add("DATACOPY4", SqlDbType.VarChar).Value = dr1ARY("DATACOPY4") 'DATACOPY4.Text
            '.Parameters.Add("DATACOPY5", SqlDbType.VarChar).Value = dr1ARY("DATACOPY5") 'DATACOPY5.Text
            '.Parameters.Add("DATACOPY6", SqlDbType.VarChar).Value = dr1ARY("DATACOPY6") 'DATACOPY6.Text
            '.Parameters.Add("D1C", SqlDbType.VarChar).Value = D1C
            '.Parameters.Add("D2C", SqlDbType.VarChar).Value = D2C
            .Parameters.Add("D3C", SqlDbType.Int).Value = dr1ARY("D3C") 'IIf(iD3C = 0, Convert.DBNull, iD3C)
            '.Parameters.Add("D4C", SqlDbType.Int).Value = dr1ARY("D4C") 'Val(D4C.SelectedValue)
            '.Parameters.Add("D5C", SqlDbType.Int).Value = dr1ARY("D5C") 'Val(D5C.SelectedValue)
            '.Parameters.Add("D6C", SqlDbType.Int).Value = dr1ARY("D6C") 'Val(D6C.SelectedValue)
            '.Parameters.Add("DATA62", SqlDbType.VarChar).Value = dr1ARY("DATA62") 'IIf(DATA62.SelectedValue = "", Convert.DBNull, DATA62.SelectedValue)
            '.Parameters.Add("DATACOPY62", SqlDbType.VarChar).Value = dr1ARY("DATACOPY62") 'IIf(DATACOPY62.Text = "", Convert.DBNull, DATACOPY62.Text)
            '.Parameters.Add("D62C", SqlDbType.Int).Value = dr1ARY("D62C") 'Val(D62C.SelectedValue)
            .Parameters.Add("ITEM7NOTE2", SqlDbType.VarChar).Value = dr1ARY("ITEM7NOTE2") 'IIf(ITEM7NOTE2.Text = "", Convert.DBNull, ITEM7NOTE2.Text)
            .Parameters.Add("D1CMM", SqlDbType.VarChar).Value = dr1ARY("D1CMM") 'IIf(D1CMM.Text = "", Convert.DBNull, D1CMM.Text)
            .Parameters.Add("D1CDD", SqlDbType.VarChar).Value = dr1ARY("D1CDD") 'IIf(D1CDD.Text = "", Convert.DBNull, D1CDD.Text)
            .Parameters.Add("D2CMM", SqlDbType.VarChar).Value = dr1ARY("D2CMM") 'IIf(D2CMM.Text = "", Convert.DBNull, D2CMM.Text)
            .Parameters.Add("D2CDD", SqlDbType.VarChar).Value = dr1ARY("D2CDD") 'IIf(D2CDD.Text = "", Convert.DBNull, D2CDD.Text)
            .Parameters.Add("D3CMM", SqlDbType.VarChar).Value = dr1ARY("D3CMM") 'IIf(D3CMM.Text = "", Convert.DBNull, D3CMM.Text)
            .Parameters.Add("D3CDD", SqlDbType.VarChar).Value = dr1ARY("D3CDD") 'IIf(D3CDD.Text = "", Convert.DBNull, D3CDD.Text)
            '.Parameters.Add("D1NOTE", SqlDbType.VarChar).Value = D1NOTE.text
            '.Parameters.Add("D2NOTE", SqlDbType.VarChar).Value = D2NOTE.text
            .Parameters.Add("D3NOTE", SqlDbType.VarChar).Value = dr1ARY("D3NOTE") 'D3NOTE.Text
            '.Parameters.Add("D4NOTE", SqlDbType.VarChar).Value = dr1ARY("D4NOTE") 'D4NOTE.Text
            '.Parameters.Add("D5NOTE", SqlDbType.VarChar).Value = D5NOTE.text
            '.Parameters.Add("D6NOTE", SqlDbType.VarChar).Value = D6NOTE.text
            .Parameters.Add("APPROVEDCOUNT", SqlDbType.Int).Value = dr1ARY("APPROVEDCOUNT") 'Val(APPROVEDCOUNT.Text)
            .Parameters.Add("AUTHCOUNT", SqlDbType.Int).Value = dr1ARY("AUTHCOUNT") 'Val(AUTHCOUNT.Text)
            .Parameters.Add("TURTHCOUNT", SqlDbType.Int).Value = dr1ARY("TURTHCOUNT") 'Val(TURTHCOUNT.Text)
            .Parameters.Add("TURNOUTCOUNT", SqlDbType.Int).Value = dr1ARY("TURNOUTCOUNT") 'Val(TURNOUTCOUNT.Text)
            .Parameters.Add("TRUANCYCOUNT", SqlDbType.Int).Value = dr1ARY("TRUANCYCOUNT") 'Val(TRUANCYCOUNT.Text)
            .Parameters.Add("LEAVECOUNT", SqlDbType.Int).Value = dr1ARY("LEAVECOUNT") 'Val(REJECTCOUNT.Text)
            .Parameters.Add("REJECTCOUNT", SqlDbType.Int).Value = dr1ARY("REJECTCOUNT") 'Val(REJECTCOUNT.Text)
            '.Parameters.Add("ADVJOBCOUNT", SqlDbType.Int).Value = dr1ARY("ADVJOBCOUNT") 'Val(ADVJOBCOUNT.Text)

            .Parameters.Add("ITEM1_1", SqlDbType.VarChar).Value = dr1ARY("ITEM1_1") ' IIf(ITEM1_1.SelectedValue = "", Convert.DBNull, ITEM1_1.SelectedValue)
            .Parameters.Add("ITEM1_2", SqlDbType.VarChar).Value = dr1ARY("ITEM1_2") 'IIf(ITEM1_2.SelectedValue = "", Convert.DBNull, ITEM1_2.SelectedValue)
            .Parameters.Add("ITEM1_COUR", SqlDbType.VarChar).Value = dr1ARY("ITEM1_COUR") 'ITEM1_COUR.Text
            .Parameters.Add("ITEM1_3", SqlDbType.VarChar).Value = dr1ARY("ITEM1_3") 'IIf(ITEM1_3.SelectedValue = "", Convert.DBNull, ITEM1_3.SelectedValue)
            .Parameters.Add("ITEM1_TEACHER", SqlDbType.VarChar).Value = dr1ARY("ITEM1_TEACHER") 'ITEM1_TEACHER.Text
            .Parameters.Add("ITEM1_ASSISTANT", SqlDbType.VarChar).Value = dr1ARY("ITEM1_ASSISTANT") 'ITEM1_ASSISTANT.Text
            .Parameters.Add("ITEM2_1", SqlDbType.VarChar).Value = dr1ARY("ITEM2_1") 'IIf(ITEM2_1.SelectedValue = "", Convert.DBNull, ITEM2_1.SelectedValue)
            .Parameters.Add("ITEM2_2", SqlDbType.VarChar).Value = dr1ARY("ITEM2_2") 'IIf(ITEM2_2.SelectedValue = "", Convert.DBNull, ITEM2_2.SelectedValue)
            '.Parameters.Add("ITEM2_3", SqlDbType.VarChar).Value = dr1ARY("ITEM2_3") 'IIf(ITEM2_3.SelectedValue = "", Convert.DBNull, ITEM2_3.SelectedValue)
            .Parameters.Add("ITEM3_1", SqlDbType.VarChar).Value = dr1ARY("ITEM3_1") 'IIf(ITEM3_1.SelectedValue = "", Convert.DBNull, ITEM3_1.SelectedValue)
            .Parameters.Add("ITEM3_2", SqlDbType.VarChar).Value = dr1ARY("ITEM3_2") 'IIf(ITEM3_2.SelectedValue = "", Convert.DBNull, ITEM3_2.SelectedValue)
            '.Parameters.Add("ITEM3_3", SqlDbType.VarChar).Value = dr1ARY("ITEM3_3") 'IIf(ITEM3_3.SelectedValue = "", Convert.DBNull, ITEM3_3.SelectedValue)
            '.Parameters.Add("ITEM3_4", SqlDbType.VarChar).Value = dr1ARY("ITEM3_4") 'IIf(ITEM3_4.SelectedValue = "", Convert.DBNull, ITEM3_4.SelectedValue)
            '.Parameters.Add("ITEM3_5", SqlDbType.VarChar).Value = dr1ARY("ITEM3_5") 'IIf(ITEM3_5.SelectedValue = "", Convert.DBNull, ITEM3_5.SelectedValue)
            '.Parameters.Add("ITEM4_1", SqlDbType.VarChar).Value = dr1ARY("ITEM4_1") 'IIf(ITEM4_1.SelectedValue = "", Convert.DBNull, ITEM4_1.SelectedValue)
            '.Parameters.Add("ITEM4_2", SqlDbType.VarChar).Value = dr1ARY("ITEM4_2") 'IIf(ITEM4_2.SelectedValue = "", Convert.DBNull, ITEM4_2.SelectedValue)
            '.Parameters.Add("ITEM4_3", SqlDbType.VarChar).Value = dr1ARY("ITEM4_3") 'IIf(ITEM4_3.SelectedValue = "", Convert.DBNull, ITEM4_3.SelectedValue)
            '.Parameters.Add("ITEM4NOTE", SqlDbType.VarChar).Value = dr1ARY("ITEM4NOTE") 'ITEM4NOTE.Text
            .Parameters.Add("ITEM7NOTE", SqlDbType.VarChar).Value = dr1ARY("ITEM7NOTE") 'ITEM7NOTE.Text
            .Parameters.Add("ITEM31NOTE", SqlDbType.VarChar).Value = dr1ARY("ITEM31NOTE") 'TEM31NOTE.Text
            .Parameters.Add("ITEM32", SqlDbType.VarChar).Value = dr1ARY("ITEM32") 'IIf(ITEM32.SelectedValue = "", Convert.DBNull, ITEM32.SelectedValue)
            .Parameters.Add("ITEM32NOTE", SqlDbType.VarChar).Value = dr1ARY("ITEM32NOTE") 'ITEM32NOTE.Text
            .Parameters.Add("ITEM1PROS", SqlDbType.VarChar).Value = dr1ARY("ITEM1PROS") 'ITEM1PROS.Text
            .Parameters.Add("ITEM2PROS", SqlDbType.VarChar).Value = dr1ARY("ITEM2PROS") 'ITEM2PROS.Text
            .Parameters.Add("ITEM3PROS", SqlDbType.VarChar).Value = dr1ARY("ITEM3PROS") 'ITEM3PROS.Text
            '.Parameters.Add("ITEM4PROS", SqlDbType.VarChar).Value = dr1ARY("ITEM4PROS") 'ITEM4PROS.Text
            .Parameters.Add("ITEM1NOTE", SqlDbType.VarChar).Value = dr1ARY("ITEM1NOTE") 'ITEM1NOTE.Text
            .Parameters.Add("ITEM2NOTE", SqlDbType.VarChar).Value = dr1ARY("ITEM2NOTE") 'ITEM2NOTE.Text
            .Parameters.Add("ITEM3NOTE", SqlDbType.VarChar).Value = dr1ARY("ITEM3NOTE") 'ITEM3NOTE.Text
            '.Parameters.Add("CURSENAME", SqlDbType.VarChar).Value = dr1ARY("CURSENAME") 'CURSENAME.Text '培訓單位人員姓名
            .Parameters.Add("VISITORNAME", SqlDbType.VarChar).Value = dr1ARY("VISITORNAME") 'VISITORNAME.Text '訪視人員姓名
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
            'dt.Load(.ExecuteReader())
            .ExecuteNonQuery()
            'rst = .ExecuteScalar()
        End With
    End Sub

    Const cst_st數字必填 As String = "數字必填"
    Const cst_st日期必填 As String = "日期必填"
    Const cst_st字串必填 As String = "字串必填"
    'Const cst_st文字必填 As String = "文字必填"
    Const cst_st字串 As String = "字串"
    Const cst_st整數 As String = "整數"

    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String) As String
        Return ChkValue1(fN1, vN1, sType, 0)
    End Function

    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String, ByVal iSize As Integer) As String
        Dim rst As String = ""
        Select Case sType
            Case cst_st字串
                If Convert.ToString(vN1) <> "" Then
                    If Convert.ToString(vN1).Length > iSize Then rst &= fN1 & "字串長度 必須小於等於" & iSize & "字數<br>"
                End If
            Case cst_st字串必填
                If Convert.ToString(vN1) <> "" Then
                    If Convert.ToString(vN1).Length > iSize Then rst &= fN1 & "字串長度 必須小於等於" & iSize & "字數<br>"
                Else
                    rst &= fN1 & " 為必填資料<br>"
                End If
            Case cst_st數字必填
                If Convert.ToString(vN1) <> "" Then
                    If Not IsNumeric(vN1) Then rst &= fN1 & "必需為數字<br>"
                Else
                    rst &= fN1 & "必須填寫<br>"
                End If
            Case cst_st日期必填
                If Convert.ToString(vN1) <> "" Then
                    If Not IsDate(vN1) Then
                        rst &= fN1 & "必須是西元年格式(yyyy/MM/dd)<br>"
                    Else
                        If CDate(vN1) < "1900/1/1" Or CDate(vN1) > "2100/1/1" Then rst &= fN1 & "範圍有誤<br>"
                    End If
                Else
                    rst &= fN1 & "必須填寫<br>"
                End If
        End Select
        Return rst
    End Function

    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String, ByVal iMin As Integer, ByVal iMax As Integer) As String
        Dim rst As String = ""
        Select Case sType
            Case cst_st整數
                If Not TIMS.IsInt(vN1) Then
                    rst &= fN1 & " 必須為整數數字<br>"
                    Return rst
                End If
                If Val(vN1) < Val(iMin) Then
                    rst &= fN1 & " 數字超過範圍<br>"
                    Return rst
                End If
                If Val(vN1) > Val(iMax) Then
                    rst &= fN1 & " 數字超過範圍<br>"
                    Return rst
                End If
        End Select
        Return rst
    End Function

    ''' <summary>
    ''' 轉換匯入資料變成 DataRow
    ''' </summary>
    ''' <param name="colArray"></param>
    ''' <returns></returns>
    Function ChgImpData(ByVal colArray As Array) As DataRow
        Dim dr1 As DataRow = Nothing
        Dim sql As String = ""
        sql = " SELECT * FROM CLASS_VISITOR3 WHERE 1<>1 "
        Dim dtV As DataTable = DbAccess.GetDataTable(sql, objconn) 'SELECT * FROM CLASS_VISITOR3 WHERE 1<>1
        dr1 = dtV.NewRow
        dr1("OCID") = TIMS.ClearSQM(colArray(cst_aOCID))
        'dr1("SEQNO") = TIMS.ClearSQM(colArray(cst_aSEQNO))
        dr1("APPLYDATE") = TIMS.Cdate2(colArray(cst_aAPPLYDATE))
        dr1("APPLYDATEHH1") = TIMS.ClearSQM(colArray(cst_aAPPLYDATEHH1))
        dr1("APPLYDATEMI1") = TIMS.ClearSQM(colArray(cst_aAPPLYDATEMI1))
        dr1("APPLYDATEHH2") = TIMS.ClearSQM(colArray(cst_aAPPLYDATEHH2))
        dr1("APPLYDATEMI2") = TIMS.ClearSQM(colArray(cst_aAPPLYDATEMI2))
        dr1("VISTIMES") = TIMS.ClearSQM(colArray(cst_aVISTIMES))

        dr1("APPROVEDCOUNT") = TIMS.ClearSQM(colArray(cst_aAPPROVEDCOUNT))
        dr1("AUTHCOUNT") = TIMS.ClearSQM(colArray(cst_aAUTHCOUNT))
        dr1("TURTHCOUNT") = TIMS.ClearSQM(colArray(cst_aTURTHCOUNT))
        dr1("TURNOUTCOUNT") = TIMS.ClearSQM(colArray(cst_aTURNOUTCOUNT))
        dr1("TRUANCYCOUNT") = TIMS.ClearSQM(colArray(cst_aTRUANCYCOUNT))
        dr1("LEAVECOUNT") = TIMS.ClearSQM(colArray(cst_aLEAVECOUNT)) '離訓人數
        dr1("REJECTCOUNT") = TIMS.ClearSQM(colArray(cst_aREJECTCOUNT))
        'dr1("ADVJOBCOUNT") = TIMS.ClearSQM(colArray(cst_aADVJOBCOUNT))

        dr1("DATA1") = TIMS.ClearSQM(colArray(cst_aDATA1))
        dr1("DATACOPY1") = TIMS.ClearSQM(colArray(cst_aDATACOPY1))
        dr1("D1CMM") = TIMS.ClearSQM(colArray(cst_aD1CMM))
        dr1("D1CDD") = TIMS.ClearSQM(colArray(cst_aD1CDD))

        dr1("DATA2") = TIMS.ClearSQM(colArray(cst_aDATA2))
        dr1("DATACOPY2") = TIMS.ClearSQM(colArray(cst_aDATACOPY2))
        dr1("D2CMM") = TIMS.ClearSQM(colArray(cst_aD2CMM))
        dr1("D2CDD") = TIMS.ClearSQM(colArray(cst_aD2CDD))

        dr1("DATA3") = TIMS.ClearSQM(colArray(cst_aDATA3))
        dr1("DATACOPY3") = TIMS.ClearSQM(colArray(cst_aDATACOPY3))
        dr1("D3C") = TIMS.ClearSQM(colArray(cst_aD3C))
        dr1("D3CMM") = TIMS.ClearSQM(colArray(cst_aD3CMM))
        dr1("D3CDD") = TIMS.ClearSQM(colArray(cst_aD3CDD))
        dr1("D3NOTE") = TIMS.ClearSQM(colArray(cst_aD3NOTE))

        'dr1("DATA4") = TIMS.ClearSQM(colArray(cst_aDATA4))
        'dr1("DATACOPY4") = TIMS.ClearSQM(colArray(cst_aDATACOPY4))
        'dr1("D4C") = TIMS.ClearSQM(colArray(cst_aD4C))
        'dr1("D4NOTE") = TIMS.ClearSQM(colArray(cst_aD4NOTE))

        'dr1("DATA5") = TIMS.ClearSQM(colArray(cst_aDATA5))
        'dr1("DATACOPY5") = TIMS.ClearSQM(colArray(cst_aDATACOPY5))
        'dr1("D5C") = TIMS.ClearSQM(colArray(cst_aD5C))
        'dr1("D5NOTE") = TIMS.ClearSQM(colArray(cst_aD5NOTE))

        'dr1("DATA6") = TIMS.ClearSQM(colArray(cst_aDATA6))
        'dr1("DATACOPY6") = TIMS.ClearSQM(colArray(cst_aDATACOPY6))
        'dr1("D6C") = TIMS.ClearSQM(colArray(cst_aD6C))
        'dr1("DATA62") = TIMS.ClearSQM(colArray(cst_aDATA62))
        'dr1("DATACOPY62") = TIMS.ClearSQM(colArray(cst_aDATACOPY62))
        'dr1("D62C") = TIMS.ClearSQM(colArray(cst_aD62C))

        'dr1("D1C") = TIMS.ClearSQM(colArray(cst_aD1C))
        'dr1("D2C") = TIMS.ClearSQM(colArray(cst_aD2C))
        'dr1("D1NOTE") = TIMS.ClearSQM(colArray(cst_aD1NOTE))
        'dr1("D2NOTE") = TIMS.ClearSQM(colArray(cst_aD2NOTE))
        'dr1("D6NOTE") = TIMS.ClearSQM(colArray(cst_aD6NOTE))

        dr1("ITEM1_1") = TIMS.ClearSQM(colArray(cst_aITEM1_1))
        dr1("ITEM1_2") = TIMS.ClearSQM(colArray(cst_aITEM1_2))
        dr1("ITEM1_COUR") = TIMS.ClearSQM(colArray(cst_aITEM1_COUR))
        dr1("ITEM1_3") = TIMS.ClearSQM(colArray(cst_aITEM1_3))
        dr1("ITEM1_TEACHER") = TIMS.ClearSQM(colArray(cst_aITEM1_TEACHER))
        dr1("ITEM1_ASSISTANT") = TIMS.ClearSQM(colArray(cst_aITEM1_ASSISTANT))

        dr1("ITEM1PROS") = TIMS.ClearSQM(colArray(cst_aITEM1PROS))
        dr1("ITEM1NOTE") = TIMS.ClearSQM(colArray(cst_aITEM1NOTE))
        dr1("ITEM2_1") = TIMS.ClearSQM(colArray(cst_aITEM2_1))
        dr1("ITEM2_2") = TIMS.ClearSQM(colArray(cst_aITEM2_2))
        'dr1("ITEM2_3") = TIMS.ClearSQM(colArray(cst_aITEM2_3))
        dr1("ITEM2NOTE") = TIMS.ClearSQM(colArray(cst_aITEM2NOTE))
        dr1("ITEM2PROS") = TIMS.ClearSQM(colArray(cst_aITEM2PROS))

        dr1("ITEM3_1") = TIMS.ClearSQM(colArray(cst_aITEM3_1))
        dr1("ITEM3_2") = TIMS.ClearSQM(colArray(cst_aITEM3_2))
        'dr1("ITEM3_3") = TIMS.ClearSQM(colArray(cst_aITEM3_3))
        'dr1("ITEM3_4") = TIMS.ClearSQM(colArray(cst_aITEM3_4))
        'dr1("ITEM3_5") = TIMS.ClearSQM(colArray(cst_aITEM3_5))
        dr1("ITEM3PROS") = TIMS.ClearSQM(colArray(cst_aITEM3PROS))
        dr1("ITEM3NOTE") = TIMS.ClearSQM(colArray(cst_aITEM3NOTE))

        'dr1("ITEM4_1") = TIMS.ClearSQM(colArray(cst_aITEM4_1))
        'dr1("ITEM4_2") = TIMS.ClearSQM(colArray(cst_aITEM4_2))
        'dr1("ITEM4_3") = TIMS.ClearSQM(colArray(cst_aITEM4_3))
        'dr1("ITEM4NOTE") = TIMS.ClearSQM(colArray(cst_aITEM4NOTE))
        dr1("ITEM7NOTE") = TIMS.ClearSQM(colArray(cst_aITEM7NOTE))
        dr1("ITEM7NOTE2") = TIMS.ClearSQM(colArray(cst_aITEM7NOTE2))
        dr1("ITEM31NOTE") = TIMS.ClearSQM(colArray(cst_aITEM31NOTE))
        dr1("ITEM32") = TIMS.ClearSQM(colArray(cst_aITEM32))
        dr1("ITEM32NOTE") = TIMS.ClearSQM(colArray(cst_aITEM32NOTE))

        'dr1("ITEM4PROS") = TIMS.ClearSQM(colArray(cst_aITEM4PROS))
        'dr1("CURSENAME") = TIMS.ClearSQM(colArray(cst_aCURSENAME))
        dr1("VISITORNAME") = TIMS.ClearSQM(colArray(cst_aVISITORNAME))

        dr1("RID") = sm.UserInfo.RID 'TIMS.ClearSQM(colArray(cst_aRID))
        dr1("MODIFYACCT") = sm.UserInfo.UserID 'TIMS.ClearSQM(colArray(cst_aMODIFYACCT))
        'dr1("MODIFYDATE") = TIMS.ClearSQM(colArray(cst_aMODIFYDATE))
        Return dr1
    End Function

    ''' <summary>
    ''' 檢查輸入資料
    ''' </summary>
    ''' <param name="colArray"></param>
    ''' <returns></returns>
    Function CheckImportData3(ByVal colArray As Array) As String
        Dim Reason As String = ""
        'Const cst_iMaxLength1 As Integer = 73
        If colArray.Length < cst_iMaxLength1 Then
            'Reason += "欄位數量不正確(應該為58個欄位)<BR>"
            Reason &= String.Format("欄位對應有誤 ,{0}/{1}<BR>", colArray.Length, cst_iMaxLength1)
            Return Reason
        End If

        Reason &= ChkValue1("班別(OCID)", colArray(cst_aOCID), cst_st數字必填)
        Reason &= ChkValue1("訪查日期", colArray(cst_aAPPLYDATE), cst_st日期必填)
        Reason &= ChkValue1("訪查日期開始時", colArray(cst_aAPPLYDATEHH1), cst_st數字必填)
        Reason &= ChkValue1("訪查日期開始分", colArray(cst_aAPPLYDATEMI1), cst_st數字必填)
        Reason &= ChkValue1("訪查日期結束時", colArray(cst_aAPPLYDATEHH2), cst_st數字必填)
        Reason &= ChkValue1("訪查日期結束分", colArray(cst_aAPPLYDATEMI2), cst_st數字必填)
        Reason &= ChkValue1("第n次訪問", colArray(cst_aVISTIMES), cst_st數字必填)

        Reason &= ChkValue1("核定人數", colArray(cst_aAPPROVEDCOUNT), cst_st數字必填)
        Reason &= ChkValue1("開訓人數", colArray(cst_aAUTHCOUNT), cst_st數字必填)
        Reason &= ChkValue1("實到人數", colArray(cst_aTURTHCOUNT), cst_st數字必填)
        Reason &= ChkValue1("請假人數", colArray(cst_aTURNOUTCOUNT), cst_st數字必填)
        Reason &= ChkValue1("缺(曠)課人數", colArray(cst_aTRUANCYCOUNT), cst_st數字必填)
        Reason &= ChkValue1("離訓人數", colArray(cst_aLEAVECOUNT), cst_st數字必填)
        Reason &= ChkValue1("退訓人數", colArray(cst_aREJECTCOUNT), cst_st數字必填)
        'Reason &= ChkValue1("提前就業人數", colArray(13), cst_st數字必填)

        Reason &= ChkValue1("書面資料1", colArray(cst_aDATA1), cst_st數字必填)
        Reason &= ChkValue1("書面資料1如附件", colArray(cst_aDATACOPY1), cst_st字串, 50)
        Reason &= ChkValue1("書面資料1備齊月", colArray(cst_aD1CMM), cst_st字串, 2)
        Reason &= ChkValue1("書面資料1備齊日", colArray(cst_aD1CDD), cst_st字串, 9)

        Reason &= ChkValue1("書面資料2", colArray(cst_aDATA2), cst_st數字必填)
        Reason &= ChkValue1("書面資料2如附件", colArray(cst_aDATACOPY2), cst_st字串, 50)
        Reason &= ChkValue1("書面資料2備齊月", colArray(cst_aD2CMM), cst_st字串, 2)
        Reason &= ChkValue1("書面資料2備齊日", colArray(cst_aD2CDD), cst_st字串, 9)

        Reason &= ChkValue1("書面資料3", colArray(cst_aDATA3), cst_st數字必填)
        Reason &= ChkValue1("書面資料3如附件", colArray(cst_aDATACOPY3), cst_st字串, 50)
        Reason &= ChkValue1("書面資料3說明選項", colArray(cst_aD3C), cst_st數字必填)
        Reason &= ChkValue1("書面資料3 攜回月", colArray(cst_aD3CMM), cst_st字串, 2)
        Reason &= ChkValue1("書面資料3 攜回日", colArray(cst_aD3CDD), cst_st字串, 9)
        Reason &= ChkValue1("書面資料3 其他說明", colArray(cst_aD3NOTE), cst_st字串, 100)

        'Reason &= ChkValue1("書面資料4", colArray(cst_aDATA4), cst_st數字必填)
        'Reason &= ChkValue1("書面資料4如附件", colArray(cst_aDATACOPY4), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料4選項", colArray(cst_aD4C), cst_st數字必填)
        'Reason &= ChkValue1("書面資料4其它說明", colArray(cst_aD4NOTE), cst_st字串, 100)

        'Reason &= ChkValue1("書面資料5", colArray(cst_aDATA5), cst_st數字必填)
        'Reason &= ChkValue1("書面資料5如附件", colArray(cst_aDATACOPY5), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料5選項", colArray(cst_aD5C), cst_st數字必填)
        'Reason &= ChkValue1("書面資料5說明", colArray(cst_aD5NOTE), cst_st字串, 100)

        'Reason &= ChkValue1("書面資料6", colArray(cst_aDATA6), cst_st數字必填)
        'Reason &= ChkValue1("書面資料6如附件", colArray(cst_aDATACOPY6), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料6選項", colArray(cst_aD6C), cst_st數字必填)
        'Reason &= ChkValue1("書面資料7", colArray(cst_aDATA62), cst_st數字必填)
        'Reason &= ChkValue1("書面資料7如附件", colArray(cst_aDATACOPY62), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料7選項", colArray(cst_aD62C), cst_st數字必填)

        Reason &= ChkValue1("課程(師資)實施狀況 1.選項", colArray(cst_aITEM1_1), cst_st數字必填)
        Reason &= ChkValue1("課程(師資)實施狀況 2.選項", colArray(cst_aITEM1_2), cst_st數字必填)
        Reason &= ChkValue1("課程(師資)實施狀況 3.課目", colArray(cst_aITEM1_COUR), cst_st字串, 100)
        Reason &= ChkValue1("課程(師資)實施狀況 教師與助教. 選項", colArray(cst_aITEM1_3), cst_st數字必填)
        Reason &= ChkValue1("課程(師資)實施狀況 教師與助教. 教師：", colArray(cst_aITEM1_TEACHER), cst_st字串必填, 100)
        Reason &= ChkValue1("課程(師資)實施狀況 教師與助教. 助教：", colArray(cst_aITEM1_ASSISTANT), cst_st字串, 100)

        Reason &= ChkValue1("課程(師資)實施狀況 處理情形", colArray(cst_aITEM1PROS), cst_st字串, 500)
        Reason &= ChkValue1("課程(師資)實施狀況 備註", colArray(cst_aITEM1NOTE), cst_st字串, 500)
        Reason &= ChkValue1("1.有無書籍(講義)領用表?", colArray(cst_aITEM2_1), cst_st數字必填)
        Reason &= ChkValue1("2.有無材料領用表?", colArray(cst_aITEM2_2), cst_st數字必填)
        'Reason &= ChkValue1("3.訓練設施設備是否依契約提供學員使用?", colArray(cst_aITEM2_3), cst_st數字必填)
        Reason &= ChkValue1("教材設施運用狀況 處理情形", colArray(cst_aITEM2NOTE), cst_st字串, 500)
        Reason &= ChkValue1("教材設施運用狀況 備註", colArray(cst_aITEM2PROS), cst_st字串, 500)

        Reason &= ChkValue1("1.教學(訓練)日誌是否確實填寫?", colArray(cst_aITEM3_1), cst_st數字必填)
        Reason &= ChkValue1("2.有否按時呈主管核閱?", colArray(cst_aITEM3_2), cst_st數字必填)
        'Reason &= ChkValue1("3.學員生活、就業輔導與管理機制是否依契約挸範辦理?", colArray(cst_aITEM3_3), cst_st數字必填)
        'Reason &= ChkValue1("4.是否依契約規範提供學員問題反應申訴管道?", colArray(cst_aITEM3_4), cst_st數字必填)
        'Reason &= ChkValue1("5.是否依契約規範公告學員權益義務管理狀況義務或編製參訓學員服務手冊?", colArray(cst_aITEM3_5), cst_st數字必填)
        Reason &= ChkValue1("教務管理狀況 處理情形", colArray(cst_aITEM3PROS), cst_st字串, 500)
        Reason &= ChkValue1("教務管理狀況 備註", colArray(cst_aITEM3NOTE), cst_st字串, 500)

        'Reason &= ChkValue1("1.是否依規定於開訓後15日內收齊職業訓練生活津貼申請書及相關證明文件後送委訓單位審查？", colArray(cst_aITEM4_1), cst_st數字必填)
        'Reason &= ChkValue1("2.培訓單位於收到本署所屬分署核撥之津貼後，是否按月即時（不超過3個工作日）轉發給受訓學員。", colArray(cst_aITEM4_2), cst_st數字必填)
        'Reason &= ChkValue1("3.申請人離、退訓時，培訓單位是否按月覈實繳回職業訓練生活津貼。", colArray(cst_aITEM4_3), cst_st數字必填)
        'Reason &= ChkValue1("免填原因說明", colArray(cst_aITEM4NOTE), cst_st字串, 100)
        'Reason &= ChkValue1("費用(津貼)收核狀況 處理情形", colArray(cst_aITEM4PROS), cst_st字串, 500)
        Reason &= ChkValue1("訓學員反映意見及問題", colArray(cst_aITEM7NOTE), cst_st字串, 500)
        Reason &= ChkValue1("學員反映意見之委訓單位反應說明", colArray(cst_aITEM7NOTE2), cst_st字串, 500)
        Reason &= ChkValue1("綜合建議", colArray(cst_aITEM31NOTE), cst_st字串, 500)
        Reason &= ChkValue1("缺失處理", colArray(cst_aITEM32), cst_st數字必填)
        Reason &= ChkValue1("缺失處理其他說明內容", colArray(cst_aITEM32NOTE), cst_st字串, 500)

        'Reason &= ChkValue1("培訓姓名", colArray(cst_aCURSENAME), cst_st字串必填, 10)
        'Reason &= ChkValue1("訪視姓名", colArray(cst_aVISITORNAME), cst_st字串必填, 10)
        Reason &= ChkValue1("訪查人員", colArray(cst_aVISITORNAME), cst_st字串必填, 10) '訪查人員

        If Reason <> "" Then Return Reason

        '23578423847239847235897
        Dim drC As DataRow = TIMS.GetOCIDDate(colArray(0), objconn)
        If drC Is Nothing Then Reason &= "班別(OCID) 查無資料<br>"
        If sm.UserInfo.LID <> 0 AndAlso drC IsNot Nothing Then
            If drC("PlanID") <> sm.UserInfo.PlanID Then Reason &= "班別計畫有誤<br>"
        End If
        'Reason &= ChkValue1("訪查日期", colArray(1), cst_st日期必填)
        If Val(colArray(cst_aAPPLYDATEHH1)) > 23 Then Reason &= "訪查日期開始時 有誤<br>"
        If Val(colArray(cst_aAPPLYDATEMI1)) > 59 Then Reason &= "訪查日期開始分 有誤<br>"

        If Val(colArray(cst_aAPPLYDATEHH2)) > 23 Then Reason &= "訪查日期結束時 有誤<br>"
        If Val(colArray(cst_aAPPLYDATEMI2)) > 59 Then Reason &= "訪查日期結束分 有誤<br>"

        Reason &= ChkValue1("訪查日期開始時", colArray(cst_aAPPLYDATEHH1), cst_st整數, 0, 23)
        Reason &= ChkValue1("訪查日期開始分", colArray(cst_aAPPLYDATEMI1), cst_st整數, 0, 59)
        Reason &= ChkValue1("訪查日期結束時", colArray(cst_aAPPLYDATEHH2), cst_st整數, 0, 23)
        Reason &= ChkValue1("訪查日期結束分", colArray(cst_aAPPLYDATEMI2), cst_st整數, 0, 59)
        Reason &= ChkValue1("第n次訪問", colArray(cst_aVISTIMES), cst_st整數, 1, 99)

        Reason &= ChkValue1("核定人數", colArray(cst_aAPPROVEDCOUNT), cst_st整數, 1, 999)
        Reason &= ChkValue1("開訓人數", colArray(cst_aAUTHCOUNT), cst_st整數, 1, 999)
        Reason &= ChkValue1("實到人數", colArray(cst_aTURTHCOUNT), cst_st整數, 1, 999)
        Reason &= ChkValue1("請假人數", colArray(cst_aTURNOUTCOUNT), cst_st整數, 0, 999)
        Reason &= ChkValue1("缺(曠)課人數", colArray(cst_aTRUANCYCOUNT), cst_st整數, 0, 999)
        Reason &= ChkValue1("離訓人數", colArray(cst_aLEAVECOUNT), cst_st整數, 0, 999)
        Reason &= ChkValue1("退訓人數", colArray(cst_aREJECTCOUNT), cst_st整數, 0, 999)
        'Reason &= ChkValue1("提前就業人數", colArray(cst_aADVJOBCOUNT), cst_st整數, 0, 999)

        Reason &= ChkValue1("書面資料1", colArray(cst_aDATA1), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料1如附件", colArray(15), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料1備齊月", colArray(16), cst_st字串, 2)
        'Reason &= ChkValue1("書面資料1備齊日", colArray(17), cst_st字串, 9)
        Reason &= ChkValue1("書面資料2", colArray(cst_aDATA2), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料2如附件", colArray(19), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料2備齊月", colArray(20), cst_st字串, 2)
        'Reason &= ChkValue1("書面資料2備齊日", colArray(21), cst_st字串, 9)
        Reason &= ChkValue1("書面資料3", colArray(cst_aDATA3), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料3如附件", colArray(23), cst_st字串, 50)
        Reason &= ChkValue1("書面資料3說明選項", colArray(cst_aD3C), cst_st整數, 1, 3)
        'Reason &= ChkValue1("書面資料3 攜回月", colArray(25), cst_st字串, 2)
        'Reason &= ChkValue1("書面資料3 攜回日", colArray(26), cst_st字串, 9)
        'Reason &= ChkValue1("書面資料3 其他說明", colArray(27), cst_st字串, 100)

        'Reason &= ChkValue1("書面資料4", colArray(cst_aDATA4), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料4如附件", colArray(29), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料4選項", colArray(cst_aD4C), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料4其它說明", colArray(31), cst_st字串, 100)

        'Reason &= ChkValue1("書面資料5", colArray(cst_aDATA5), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料5如附件", colArray(33), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料5選項", colArray(cst_aD5C), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料5說明", colArray(35), cst_st字串, 100)

        'Reason &= ChkValue1("書面資料6", colArray(cst_aDATA6), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料6如附件", colArray(37), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料6選項", colArray(cst_aD6C), cst_st整數, 1, 2)

        'Reason &= ChkValue1("書面資料7", colArray(cst_aDATA62), cst_st整數, 1, 4)
        'Reason &= ChkValue1("書面資料7如附件", colArray(40), cst_st字串, 50)
        'Reason &= ChkValue1("書面資料7選項", colArray(cst_aD62C), cst_st整數, 1, 1)

        Reason &= ChkValue1("課程(師資)實施狀況 1. 選項", colArray(cst_aITEM1_1), cst_st整數, 1, 3)
        Reason &= ChkValue1("課程(師資)實施狀況 2. 選項", colArray(cst_aITEM1_2), cst_st整數, 1, 3)
        'Reason &= ChkValue1("課程(師資)實施狀況 3. 課目", colArray(44), cst_st字串, 100)
        Reason &= ChkValue1("課程(師資)實施狀況 教師與助教. 選項", colArray(cst_aITEM1_3), cst_st整數, 1, 3)
        'Reason &= ChkValue1("課程(師資)實施狀況 教師與助教. 教師：", colArray(46), cst_st字串, 100)
        'Reason &= ChkValue1("課程(師資)實施狀況 教師與助教. 助教：", colArray(47), cst_st字串, 100)
        'Reason &= ChkValue1("課程(師資)實施狀況 處理情形", colArray(48), cst_st字串, 500)
        'Reason &= ChkValue1("課程(師資)實施狀況 備註", colArray(49), cst_st字串, 500)
        Reason &= ChkValue1("1.有無書籍(講義)領用表?", colArray(cst_aITEM2_1), cst_st整數, 1, 3)
        Reason &= ChkValue1("2.有無材料領用表?", colArray(cst_aITEM2_2), cst_st整數, 1, 3)
        'Reason &= ChkValue1("3.訓練設施設備是否依契約提供學員使用?", colArray(cst_aITEM2_3), cst_st整數, 1, 3)
        'Reason &= ChkValue1("教材設施運用狀況 處理情形", colArray(53), cst_st字串, 500)
        'Reason &= ChkValue1("教材設施運用狀況 備註", colArray(54), cst_st字串, 500)

        Reason &= ChkValue1("1.教學(訓練)日誌是否確實填寫?", colArray(cst_aITEM3_1), cst_st整數, 1, 3)
        Reason &= ChkValue1("2.有否按時呈主管核閱?", colArray(cst_aITEM3_2), cst_st整數, 1, 3)
        'Reason &= ChkValue1("3.學員生活、就業輔導與管理機制是否依契約挸範辦理?", colArray(cst_aITEM3_3), cst_st整數, 1, 3)
        'Reason &= ChkValue1("4.是否依契約規範提供學員問題反應申訴管道?", colArray(cst_aITEM3_4), cst_st整數, 1, 3)
        'Reason &= ChkValue1("5.是否依契約規範公告學員權益義務管理狀況義務或編製參訓學員服務手冊?", colArray(cst_aITEM3_5), cst_st整數, 1, 3)
        'Reason &= ChkValue1("教務管理狀況 處理情形", colArray(60), cst_st字串, 500)
        'Reason &= ChkValue1("教務管理狀況 備註", colArray(61), cst_st字串, 500)
        'Reason &= ChkValue1("1.是否依規定於開訓後15日內收齊職業訓練生活津貼申請書及相關證明文件後送委訓單位審查？", colArray(cst_aITEM4_1), cst_st整數, 1, 3)
        'Reason &= ChkValue1("2.培訓單位於收到本署所屬分署核撥之津貼後，是否按月即時（不超過3個工作日）轉發給受訓學員。", colArray(cst_aITEM4_2), cst_st整數, 1, 3)
        'Reason &= ChkValue1("3.申請人離、退訓時，培訓單位是否按月覈實繳回職業訓練生活津貼。", colArray(cst_aITEM4_3), cst_st整數, 1, 3)
        'Reason &= ChkValue1("免填原因說明", colArray(cst_aITEM4NOTE), cst_st字串, 100)
        'Reason &= ChkValue1("費用(津貼)收核狀況 處理情形", colArray(cst_aITEM4PROS), cst_st字串, 500)
        'Reason &= ChkValue1("訓學員反映意見及問題", colArray(cst_aITEM7NOTE), cst_st字串, 500)
        'Reason &= ChkValue1("學員反映意見之委訓單位反應說明", colArray(cst_aITEM7NOTE2), cst_st字串, 500)
        'Reason &= ChkValue1("綜合建議", colArray(cst_aITEM31NOTE), cst_st字串, 500)
        Reason &= ChkValue1("缺失處理", colArray(cst_aITEM32), cst_st整數, 1, 4)
        'Reason &= ChkValue1("缺失處理其他說明內容", colArray(cst_aITEM32NOTE), cst_st字串, 500)
        'Reason &= ChkValue1("培訓姓名", colArray(cst_aCURSENAME), cst_st字串, 10)
        'Reason &= ChkValue1("訪視姓名", colArray(cst_aVISITORNAME), cst_st字串, 10)
        Return Reason
    End Function

#End Region

#Region "NO USE - sub_XLSImp1"
    '匯入名冊
    'Sub sub_XLSImp1()
    '    Dim RowIndex As Integer = 1
    '    Dim Reason As String = "" '儲存錯誤的原因
    '    Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
    '    Dim drWrong As DataRow = Nothing

    '    '建立錯誤資料格式Table
    '    dtWrong.Columns.Add(New DataColumn("Index"))
    '    dtWrong.Columns.Add(New DataColumn("OCID"))
    '    dtWrong.Columns.Add(New DataColumn("VisitorName"))
    '    dtWrong.Columns.Add(New DataColumn("ApplyDate"))
    '    dtWrong.Columns.Add(New DataColumn("Reason"))

    '    Dim MyFileName As String = ""
    '    Dim MyFileType As String = ""
    '    Dim dt_xls As DataTable = Nothing
    '    Const Cst_FileSavePath As String = "~/CP/01/Temp/"

    '    If Me.File1.Value = "" Then
    '        Common.MessageBox(Me, "未輸入匯入檔案位置")
    '        Exit Sub
    '    End If
    '    If File1.PostedFile.ContentLength = 0 Then
    '        Common.MessageBox(Me, "檔案位置錯誤!")
    '        Exit Sub
    '    End If
    '    '取出檔案名稱
    '    MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
    '    '取出檔案類型
    '    If MyFileName.IndexOf(".") = -1 Then
    '        Common.MessageBox(Me, "檔案類型錯誤!")
    '        Exit Sub
    '    End If
    '    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
    '    If MyFileType <> "xls" Then
    '        Common.MessageBox(Me, "檔案類型錯誤，必須為XLS檔!")
    '        Exit Sub
    '    End If

    '    '上傳檔案
    '    File1.PostedFile.SaveAs(Server.MapPath(Cst_FileSavePath & MyFileName))

    '    '取得內容
    '    Dim fullFilNm1 As String = Server.MapPath(Cst_FileSavePath & MyFileName).ToString
    '    dt_xls = TIMS.GetDataTable_XlsFile(fullFilNm1, "", Reason, "班別(OCID)")

    '    IO.File.Delete(Server.MapPath(Cst_FileSavePath & MyFileName)) '刪除檔案

    '    If Reason <> "" Then
    '        Common.MessageBox(Me, Reason)
    '        Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
    '        Exit Sub
    '    End If

    '    'xls 方式 讀取寫入資料庫
    '    If dt_xls.Rows.Count = 0 Then '有資料
    '        Common.MessageBox(Me, "查無匯入資料!!")
    '        Exit Sub
    '    End If

    '    Dim tConn As SqlConnection = DbAccess.GetConnection
    '    Call TIMS.OpenDbConn(tConn)

    '    '有資料
    '    Reason = ""
    '    For i As Integer = 0 To dt_xls.Rows.Count - 1
    '        If RowIndex <> 0 Then
    '            Dim colArray As Array = dt_xls.Rows(i).ItemArray
    '            Reason = CheckImportData(colArray)
    '            If Reason = "" Then Call Savedata2(tConn, colArray) '無錯誤存檔 '匯入資料
    '            If Reason <> "" Then
    '                '錯誤資料，填入錯誤資料表
    '                drWrong = dtWrong.NewRow
    '                dtWrong.Rows.Add(drWrong)
    '                drWrong("Index") = RowIndex
    '                If colArray.Length > 5 Then
    '                    drWrong("OCID") = colArray(0)
    '                    If colArray.Length > 82 Then
    '                        If Convert.ToString(colArray(82)) <> "" Then
    '                            drWrong("VisitorName") = Convert.ToString(colArray(82))
    '                        Else
    '                            drWrong("VisitorName") = "未填寫"
    '                        End If
    '                    Else
    '                        drWrong("VisitorName") = "查無此欄位"
    '                    End If
    '                    drWrong("ApplyDate") = colArray(1)
    '                    drWrong("Reason") = Reason
    '                End If
    '            End If
    '        End If
    '        RowIndex += 1
    '    Next

    '    Call TIMS.CloseDbConn(tConn)

    '    '判斷匯出資料是否有誤
    '    Dim explain, explain2 As String
    '    explain = ""
    '    explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
    '    explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
    '    explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
    '    explain2 = ""
    '    explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
    '    explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
    '    explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"
    '    If dtWrong.Rows.Count > 0 Then
    '        Session("MyWrongTable") = dtWrong
    '        Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('CP_01_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
    '        Exit Sub
    '    End If
    '    If Reason <> "" Then
    '        Common.MessageBox(Me, explain & Reason)
    '        Exit Sub
    '    End If
    '    Common.MessageBox(Me, explain)
    '    Exit Sub
    'End Sub

    'Sub Savedata2(ByRef tConn As SqlConnection, ByRef colArray As Array)
    '    Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
    '    Dim sql As String = ""
    '    Try
    '        Dim iSeqNo As Integer = 0
    '        sql = " SELECT MAX(SeqNO) SeqNO FROM Class_Visitor WHERE OCID = '" & colArray(0) & "' "
    '        Dim dr As DataRow = DbAccess.GetOneRow(sql, trans)
    '        If dr IsNot Nothing Then
    '            iSeqNo = If(IsDBNull(dr("SeqNO")), 0, CInt(dr("SeqNO")))
    '        End If

    '        Dim da As SqlDataAdapter = Nothing
    '        Dim dt As DataTable = Nothing
    '        sql = " SELECT * FROM Class_Visitor WHERE OCID = '" & colArray(0) & "' AND ApplyDate = CONVERT(DATETIME, '" & colArray(1) & "', 111) ORDER BY SeqNo DESC "
    '        dt = DbAccess.GetDataTable(sql, da, trans)
    '        If dt.Rows.Count = 0 Then
    '            dr = dt.NewRow()
    '            dt.Rows.Add(dr)
    '            iSeqNo += 1
    '        Else
    '            dr = dt.Rows(0)
    '        End If

    '        If sm.UserInfo.Years <= 2008 Then
    '            dr("OCID") = colArray(0).ToString '班別(OCID)
    '            dr("ApplyDate") = colArray(1).ToString '訪查日期 
    '            dr("AuthCount") = colArray(2).ToString '核定人數
    '            dr("TurthCount") = colArray(3).ToString '實到人數
    '            dr("TurnoutCount") = colArray(4).ToString '請假人數
    '            dr("TruancyCount") = colArray(5).ToString '缺(曠)課人數
    '            dr("RejectCount") = colArray(6).ToString '退訓人數
    '            dr("Data1") = colArray(7).ToString '書面資料1
    '            If colArray(8).ToString <> "" Then dr("DataCopy1") = colArray(8).ToString '書面資料1攜回影本
    '            If colArray(9).ToString <> "" Then dr("Data1Note") = colArray(9).ToString '書面資料1說明
    '            dr("Data2") = colArray(10).ToString '書面資料2
    '            If colArray(11).ToString <> "" Then dr("DataCopy2") = colArray(11).ToString '書面資料2攜回影本
    '            If colArray(12).ToString <> "" Then dr("Data2Note") = colArray(12).ToString '書面資料2說明
    '            dr("Data3") = colArray(13).ToString '書面資料3
    '            If colArray(14).ToString <> "" Then dr("DataCopy3") = colArray(14).ToString '書面資料3攜回影本
    '            If colArray(15).ToString <> "" Then dr("Data3Note") = colArray(15).ToString '書面資料3說明
    '            dr("Data4") = colArray(16).ToString '書面資料4
    '            If colArray(17).ToString <> "" Then dr("DataCopy4") = colArray(17).ToString '書面資料4攜回影本
    '            If colArray(18).ToString <> "" Then dr("Data4Note") = colArray(18).ToString '書面資料4說明
    '            dr("Data5") = colArray(19).ToString '書面資料5
    '            If colArray(20).ToString <> "" Then dr("DataCopy5") = colArray(20).ToString '書面資料5攜回影本
    '            If colArray(21).ToString <> "" Then dr("Data5Note") = colArray(21).ToString '書面資料5說明
    '            dr("Data6") = colArray(22).ToString '書面資料6
    '            If colArray(23).ToString <> "" Then dr("DataCopy6") = colArray(23).ToString '書面資料6攜回影本
    '            If colArray(24).ToString <> "" Then dr("Data6Note") = colArray(24).ToString '書面資料6說明
    '            dr("Data7") = colArray(25).ToString '書面資料7
    '            If colArray(26).ToString <> "" Then dr("DataCopy7") = colArray(26).ToString '書面資料7攜回影本
    '            If colArray(27).ToString <> "" Then dr("Data7Note") = colArray(27).ToString '書面資料7說明
    '            dr("Item1_1") = colArray(28).ToString '項次1_1有無週(月)課程表?
    '            dr("Item1_2") = colArray(29).ToString '項次1_2是否依課程表授課?
    '            dr("Item1_3") = colArray(30).ToString '項次1_3課目或課題為何?
    '            If colArray(31).ToString <> "" Then dr("Item1Pros") = colArray(31).ToString '項次1處理情形
    '            If colArray(32).ToString <> "" Then dr("Item1Note") = colArray(32).ToString '項次1備註
    '            dr("Item2_1") = colArray(33).ToString '項次2_1教學(訓練)日誌是否確實填寫?
    '            dr("Item2_2") = colArray(34).ToString '項次2_2有否按時呈主管核閱?
    '            If colArray(35).ToString <> "" Then dr("Item2Pros") = colArray(35).ToString '項次2處理情形
    '            If colArray(36).ToString <> "" Then dr("Item2Note") = colArray(36).ToString '項次2備註
    '            dr("Item3_1") = colArray(37).ToString '項次3_1教師(職業訓練師)與助教姓名_是否與計畫相符?
    '            dr("Item3_1Tech") = colArray(38).ToString '項次3_1講師
    '            dr("Item3_1Tutor") = colArray(39).ToString '項次3_1助教
    '            dr("Item3_2") = colArray(40).ToString '項次3_2學員學習情況是否良好?
    '            If colArray(41).ToString <> "" Then dr("Item3Pros") = colArray(41).ToString '項次3處理情形
    '            If colArray(42).ToString <> "" Then dr("Item3Note") = colArray(42).ToString '項次3備註
    '            dr("Item4_1") = colArray(43).ToString '項次4教學環境(教室或訓練工場)是否整齊
    '            If colArray(44).ToString <> "" Then dr("Item4Pros") = colArray(44).ToString '項次4處理情形
    '            If colArray(45).ToString <> "" Then dr("Item4Note") = colArray(45).ToString '項次4備註
    '            dr("Item5_1") = colArray(46).ToString '項次5有無具體輔導活動或其他事項?
    '            If colArray(47).ToString <> "" Then dr("Item5Pros") = colArray(47).ToString '項次5處理情形
    '            If colArray(48).ToString <> "" Then dr("Item5Note") = colArray(48).ToString '項次5備註
    '            dr("Item6_1") = colArray(49).ToString '項次6職業訓練生活津貼是否依規定申請並發放?
    '            dr("Item6Count1") = colArray(50).ToString '項次6_請領人數
    '            dr("Item6Count2") = colArray(51).ToString '項次6_無異狀人數
    '            dr("Item6Count3") = colArray(52).ToString '項次6_需追蹤人數
    '            If colArray(53).ToString <> "" Then dr("Item6Names") = colArray(53).ToString '項次6_需追蹤姓名
    '            If colArray(54).ToString <> "" Then dr("Item6Note") = colArray(54).ToString '項次6備註
    '            If colArray(55).ToString <> "" Then dr("Item7Note") = colArray(55).ToString '項次7訓學員反映意見及問題:
    '            dr("CurseName") = colArray(56).ToString '培訓單位人員姓名
    '            dr("VisitorName") = colArray(57).ToString '訪視人員姓名
    '        ElseIf sm.UserInfo.Years = 2009 Then
    '            dr("OCID") = colArray(0).ToString '班別(OCID)
    '            dr("ApplyDate") = colArray(1).ToString '訪查日期 
    '            dr("AuthCount") = colArray(2).ToString '核定人數
    '            dr("TurthCount") = colArray(3).ToString '實到人數
    '            dr("TurnoutCount") = colArray(4).ToString '請假人數
    '            dr("TruancyCount") = colArray(5).ToString '缺(曠)課人數
    '            dr("RejectCount") = colArray(6).ToString '退訓人數
    '            dr("AheadjobCount") = colArray(7).ToString '提前就業人數
    '            dr("Data1") = colArray(8).ToString '書面資料1
    '            If colArray(9).ToString <> "" Then dr("D1c") = colArray(9).ToString '佐證資料1check
    '            If colArray(10).ToString <> "" Then dr("DataCopy1") = colArray(10).ToString '書面資料1如附件
    '            If colArray(11).ToString <> "" Then dr("Data1Note") = colArray(11).ToString '書面資料1說明
    '            dr("Data3") = colArray(12).ToString '書面資料2
    '            If colArray(13).ToString <> "" Then dr("D2c") = colArray(13).ToString '佐證資料2check
    '            If colArray(14).ToString <> "" Then dr("DataCopy3") = colArray(14).ToString '書面資料2如附件
    '            If colArray(15).ToString <> "" Then dr("Data3Note") = colArray(15).ToString '書面資料2說明
    '            dr("Data5") = colArray(16).ToString '書面資料3
    '            If colArray(17).ToString <> "" Then dr("D3c") = colArray(17).ToString '佐證資料3check
    '            If colArray(18).ToString <> "" Then dr("DataCopy5") = colArray(18).ToString '書面資料3如附件
    '            If colArray(19).ToString <> "" Then dr("D3c3") = colArray(19).ToString '書面資料說明選項
    '            If colArray(20).ToString <> "" Then dr("Data5Note") = colArray(20).ToString '書面資料3說明
    '            dr("Data6") = colArray(21).ToString '書面資料4
    '            If colArray(22).ToString <> "" Then dr("D4c") = colArray(22).ToString '佐證資料4check
    '            If colArray(23).ToString <> "" Then dr("DataCopy6") = colArray(23).ToString '書面資料4如附件
    '            If colArray(24).ToString <> "" Then dr("D4c3") = colArray(24).ToString '說明選項4
    '            If colArray(25).ToString <> "" Then dr("Data6Note") = colArray(25).ToString '書面資料4說明
    '            dr("Data7") = colArray(26).ToString '書面資料5
    '            If colArray(27).ToString <> "" Then dr("D5c") = colArray(27).ToString '佐證資料5check
    '            If colArray(28).ToString <> "" Then dr("DataCopy7") = colArray(28).ToString '書面資料5如附件
    '            If colArray(29).ToString <> "" Then dr("D5c3") = colArray(29).ToString '說明選項5
    '            If colArray(30).ToString <> "" Then dr("Data7Note") = colArray(30).ToString '書面資料5說明
    '            dr("Data9") = colArray(31).ToString '書面資料6
    '            If colArray(32).ToString <> "" Then dr("D6c") = colArray(32).ToString '佐證資料6check
    '            If colArray(33).ToString <> "" Then dr("DataCopy9") = colArray(33).ToString '書面資料6如附件
    '            If colArray(34).ToString <> "" Then dr("D6c3") = colArray(34).ToString '說明選項6
    '            If colArray(35).ToString <> "" Then dr("Data9Note") = colArray(35).ToString '書面資料6說明
    '            dr("Data10") = colArray(36).ToString '書面資料7
    '            If colArray(37).ToString <> "" Then dr("D7c") = colArray(37).ToString '佐證資料7check
    '            If colArray(38).ToString <> "" Then dr("DataCopy10") = colArray(38).ToString '書面資料7如附件
    '            If colArray(39).ToString <> "" Then dr("D7c3") = colArray(39).ToString '說明選項7
    '            If colArray(40).ToString <> "" Then dr("Data10Note") = colArray(40).ToString '書面資料7說明
    '            dr("Data11") = colArray(41).ToString '書面資料8
    '            If colArray(42).ToString <> "" Then dr("D8c") = colArray(42).ToString '佐證資料8check
    '            If colArray(43).ToString <> "" Then dr("DataCopy11") = colArray(43).ToString '書面資料8如附件
    '            If colArray(44).ToString <> "" Then dr("D8c3") = colArray(44).ToString '說明選項8
    '            If colArray(45).ToString <> "" Then dr("Data11Note") = colArray(45).ToString '書面資料8說明
    '            dr("Item1_1") = colArray(46).ToString '項次1_1有無週(月)課程表?
    '            dr("Item1_2") = colArray(47).ToString '項次1_2是否依課程表授課?
    '            dr("Item1_3") = colArray(48).ToString '項次1_3課目或課題為何?
    '            dr("Item3_1") = colArray(49).ToString '項次3_1教師(職業訓練師)與助教姓名_是否與計畫相符?
    '            dr("Item3_1Tech") = colArray(50).ToString '項次3_1講師
    '            dr("Item3_1Tutor") = colArray(51).ToString '項次3_1助教
    '            If colArray(52).ToString <> "" Then dr("Item1Pros") = colArray(52).ToString '項次1處理情形
    '            If colArray(53).ToString <> "" Then dr("Item1Note") = colArray(53).ToString '項次1備註
    '            dr("Item19") = colArray(54).ToString '項次有無書籍(講義)領用表?
    '            dr("Item20") = colArray(55).ToString '項次有無材料領用表?
    '            dr("Item21") = colArray(56).ToString '項次訓練設施設備是否依契約提供學員使用?
    '            If colArray(57).ToString <> "" Then dr("Item2Pros") = colArray(57).ToString '項次2處理情形
    '            If colArray(58).ToString <> "" Then dr("Item2Note") = colArray(58).ToString '項次2備註
    '            dr("Item2_1") = colArray(59).ToString '項次2_1教學(訓練)日誌是否確實填寫?
    '            dr("Item2_2") = colArray(60).ToString '項次2_2有否按時呈主管核閱?
    '            dr("Item23") = colArray(61).ToString '項次學員生活就業輔導與管理機制是否依契約規範辦理?
    '            dr("Item24") = colArray(62).ToString '項次是否依契約規範提供學員問題反應申訴管道?
    '            dr("Item25") = colArray(63).ToString '項次是否為參訓學員辦理勞工保險加退保?
    '            dr("Item26") = colArray(64).ToString '項次是否依契約規範公告學員權益教務管理狀況義務或編製參訓學員服務手冊?
    '            If colArray(65).ToString <> "" Then dr("Item3Pros") = colArray(65).ToString '項次3處理情形
    '            If colArray(66).ToString <> "" Then dr("Item3Note") = colArray(66).ToString '項次3備註
    '            dr("Item28") = colArray(67).ToString '項次有無自費參訓學員?
    '            If colArray(68).ToString <> "" Then dr("Item28count") = colArray(68).ToString '項次幾人?
    '            dr("Item28_2") = colArray(69).ToString '項次訓練單位是否繳交主辦單位?
    '            dr("Item6_1") = colArray(70).ToString '項次職業訓練生活津貼是否依規定申請並核發?
    '            dr("Item29") = colArray(71).ToString '項次培訓單位是否巧立名目強制收取費用?
    '            If colArray(72).ToString <> "" Then dr("Item4Pros") = colArray(72).ToString '項次4處理情形
    '            If colArray(73).ToString <> "" Then dr("Item4Note") = colArray(73).ToString '項次4備註
    '            dr("Item30") = colArray(74).ToString '項次職業訓練機構是否依規定懸掛設立許可證書?
    '            If colArray(75).ToString <> "" Then dr("Item5Pros") = colArray(75).ToString '項次5處理情形
    '            If colArray(76).ToString <> "" Then dr("Item5Note") = colArray(76).ToString '項次5備註
    '            If colArray(77).ToString <> "" Then dr("Item7Note") = colArray(77).ToString '項次7訓學員反映意見及問題: 
    '            If colArray(78).ToString <> "" Then dr("Item31Note") = colArray(78).ToString '項次綜合建議 
    '            If colArray(79).ToString <> "" Then dr("Item32") = colArray(79).ToString '缺失處理
    '            If colArray(80).ToString <> "" Then dr("Item32Note") = colArray(80).ToString '說明內容
    '            dr("CurseName") = colArray(81).ToString '培訓單位人員姓名
    '            dr("VisitorName") = colArray(82).ToString '訪視人員姓名
    '            dr("RID") = sm.UserInfo.RID    '取得RID
    '        ElseIf sm.UserInfo.Years >= 2010 Then
    '            dr("OCID") = colArray(0).ToString '班別(OCID)
    '            dr("ApplyDate") = colArray(1).ToString '訪查日期 
    '            dr("AuthCount") = colArray(2).ToString '核定人數
    '            dr("TurthCount") = colArray(3).ToString '實到人數
    '            dr("TurnoutCount") = colArray(4).ToString '請假人數
    '            dr("TruancyCount") = colArray(5).ToString '缺(曠)課人數
    '            dr("RejectCount") = colArray(6).ToString '退訓人數
    '            dr("AheadjobCount") = colArray(7).ToString '提前就業人數
    '            dr("Data1") = colArray(8).ToString '書面資料1
    '            If colArray(9).ToString <> "" Then dr("D1c") = colArray(9).ToString '佐證資料1check
    '            If colArray(10).ToString <> "" Then dr("DataCopy1") = colArray(10).ToString '書面資料1如附件
    '            If colArray(11).ToString <> "" Then dr("Data1Note") = colArray(11).ToString '書面資料1說明
    '            dr("Data3") = colArray(12).ToString '書面資料2
    '            If colArray(13).ToString <> "" Then dr("D2c") = colArray(13).ToString '佐證資料2check
    '            If colArray(14).ToString <> "" Then dr("DataCopy3") = colArray(14).ToString '書面資料2如附件
    '            If colArray(15).ToString <> "" Then dr("Data3Note") = colArray(15).ToString '書面資料2說明
    '            dr("Data5") = colArray(16).ToString '書面資料3
    '            If colArray(17).ToString <> "" Then dr("D3c") = colArray(17).ToString '佐證資料3check
    '            If colArray(18).ToString <> "" Then dr("DataCopy5") = colArray(18).ToString '書面資料3如附件
    '            If colArray(19).ToString <> "" Then dr("D3c3") = colArray(19).ToString '書面資料說明選項
    '            If colArray(20).ToString <> "" Then dr("Data5Note") = colArray(20).ToString '書面資料3說明
    '            dr("Data6") = colArray(21).ToString '書面資料4
    '            If colArray(22).ToString <> "" Then dr("D4c") = colArray(22).ToString '佐證資料4check
    '            If colArray(23).ToString <> "" Then dr("DataCopy6") = colArray(23).ToString '書面資料4如附件
    '            If colArray(24).ToString <> "" Then dr("D4c3") = colArray(24).ToString '說明選項4
    '            If colArray(25).ToString <> "" Then dr("Data6Note") = colArray(25).ToString '書面資料4說明
    '            dr("Data10") = colArray(26).ToString '書面資料5
    '            If colArray(27).ToString <> "" Then dr("D7c") = colArray(27).ToString '佐證資料5check
    '            If colArray(28).ToString <> "" Then dr("DataCopy10") = colArray(28).ToString '書面資料5如附件
    '            If colArray(29).ToString <> "" Then dr("D7c3") = colArray(29).ToString '說明選項5
    '            If colArray(30).ToString <> "" Then dr("Data10Note") = colArray(30).ToString '書面資料5說明
    '            dr("Item1_1") = colArray(31).ToString '項次1_1有無週(月)課程表?
    '            dr("Item1_2") = colArray(32).ToString '項次1_2是否依課程表授課?
    '            dr("Item1_3") = colArray(33).ToString '項次1_3課目或課題為何?
    '            dr("Item3_1") = colArray(34).ToString '項次3_1教師(職業訓練師)與助教姓名_是否與計畫相符?
    '            dr("Item3_1Tech") = colArray(35).ToString '項次3_1講師
    '            dr("Item3_1Tutor") = colArray(36).ToString '項次3_1助教
    '            If colArray(37).ToString <> "" Then dr("Item1Pros") = colArray(37).ToString '項次1處理情形
    '            If colArray(38).ToString <> "" Then dr("Item1Note") = colArray(38).ToString '項次1備註
    '            dr("Item19") = colArray(39).ToString '項次有無書籍(講義)領用表?
    '            dr("Item20") = colArray(40).ToString '項次有無材料領用表?
    '            dr("Item21") = colArray(41).ToString '項次訓練設施設備是否依契約提供學員使用?
    '            If colArray(42).ToString <> "" Then dr("Item2Pros") = colArray(42).ToString '項次2處理情形
    '            If colArray(43).ToString <> "" Then dr("Item2Note") = colArray(43).ToString '項次2備註
    '            dr("Item2_1") = colArray(44).ToString '項次2_1教學(訓練)日誌是否確實填寫?
    '            dr("Item2_2") = colArray(45).ToString '項次2_2有否按時呈主管核閱?
    '            dr("Item23") = colArray(46).ToString '項次學員生活就業輔導與管理機制是否依契約規範辦理?
    '            dr("Item24") = colArray(47).ToString '項次是否依契約規範提供學員問題反應申訴管道?
    '            dr("Item25") = colArray(48).ToString '項次是否為參訓學員辦理勞工保險加退保?
    '            dr("Item26") = colArray(49).ToString '項次是否依契約規範公告學員權益教務管理狀況義務或編製參訓學員服務手冊?
    '            If colArray(50).ToString <> "" Then dr("Item3Pros") = colArray(50).ToString '項次3處理情形
    '            If colArray(51).ToString <> "" Then dr("Item3Note") = colArray(51).ToString '項次3備註
    '            dr("Item6_1") = colArray(52).ToString '項次職業訓練生活津貼是否依規定申請並核發?
    '            If colArray(53).ToString <> "" Then dr("Item4Pros") = colArray(53).ToString '項次4處理情形
    '            If colArray(54).ToString <> "" Then dr("Item4Note") = colArray(54).ToString '項次4備註
    '            If colArray(55).ToString <> "" Then dr("Item7Note") = colArray(55).ToString '項次7訓學員反映意見及問題:
    '            If colArray(56).ToString <> "" Then dr("Item31Note") = colArray(56).ToString '項次綜合建議 
    '            If colArray(57).ToString <> "" Then dr("Item32") = colArray(57).ToString '缺失處理
    '            If colArray(58).ToString <> "" Then dr("Item32Note") = colArray(58).ToString '說明內容
    '            dr("CurseName") = colArray(59).ToString '培訓單位人員姓名
    '            dr("VisitorName") = colArray(60).ToString '訪視人員姓名
    '            dr("RID") = sm.UserInfo.RID    '取得RID
    '        End If
    '        dr("ModifyAcct") = sm.UserInfo.UserID '異動者
    '        dr("ModifyDate") = Now() '異動時間
    '        If dr("SeqNo").ToString = "" Then dr("SeqNo") = iSeqNo
    '        DbAccess.UpdateDataTable(dt, da, trans)
    '        DbAccess.CommitTrans(trans)
    '    Catch ex As Exception
    '        DbAccess.RollbackTrans(trans)
    '        Call TIMS.CloseDbConn(tConn)
    '        Throw ex
    '    End Try
    'End Sub

    'Function CheckImportData(ByVal colArray As Array) As String
    '    Dim Reason As String = ""
    '    'Dim i, j, subCount As Integer
    '    'Dim sql As String

    '    If sm.UserInfo.Years <= 2008 Then
    '        If colArray.Length <> 58 Then
    '            'Reason += "欄位數量不正確(應該為58個欄位)<BR>"
    '            Reason += "欄位對應有誤<BR>"
    '        Else
    '            If colArray(0).ToString = "" Then
    '                Reason += "班別(OCID)必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(0)) = False Then Reason += "班別(OCID)必需為數字<BR>"
    '            End If
    '            If colArray(1).ToString = "" Then
    '                Reason += "訪查日期必須填寫<Br>"
    '            Else
    '                If IsDate(colArray(1)) = False Then
    '                    Reason += "訪查日期必須是西元年格式(yyyy/MM/dd)<BR>"
    '                Else
    '                    If CDate(colArray(1)) < "1900/1/1" Or CDate(colArray(1)) > "2100/1/1" Then Reason += "訪查日期範圍有誤<BR>"
    '                End If
    '            End If
    '            If colArray(2).ToString = "" Then
    '                Reason += "核定人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(2)) = False Then Reason += "核定人數必需為數字<BR>"
    '            End If
    '            If colArray(3).ToString = "" Then
    '                Reason += "實到人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(3)) = False Then Reason += "實到人數必需為數字<BR>"
    '            End If
    '            If colArray(4).ToString = "" Then
    '                Reason += "請假人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(4)) = False Then Reason += "請假人數必需為數字<BR>"
    '            End If
    '            If colArray(5).ToString = "" Then
    '                Reason += "缺(曠)課人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(5)) = False Then Reason += "缺(曠)課人數必需為數字<BR>"
    '            End If
    '            If colArray(6).ToString = "" Then
    '                Reason += "退訓人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(6)) = False Then Reason += "退訓人數必需為數字<BR>"
    '            End If
    '            If colArray(7).ToString = "" Then
    '                Reason += "書面資料1必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(7)) = False Then Reason += "書面資料1必需為數字<BR>"
    '            End If
    '            If colArray(8).ToString <> "" Then
    '                If (colArray(8).ToString.Length > 50) Then Reason += "書面資料1攜回影本必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(9).ToString <> "" Then
    '                If (colArray(9).ToString.Length > 100) Then Reason += "書面資料1說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(10).ToString = "" Then
    '                Reason += "書面資料2必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(10)) = False Then
    '                    Reason += "書面資料2必需為數字<BR>"
    '                End If
    '            End If
    '            If colArray(11).ToString <> "" Then
    '                If (colArray(11).ToString.Length > 50) Then Reason += "書面資料2攜回影本必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(12).ToString <> "" Then
    '                If (colArray(12).ToString.Length > 100) Then Reason += "書面資料2說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(13).ToString = "" Then
    '                Reason += "書面資料3必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(13)) = False Then Reason += "書面資料3必需為數字<BR>"
    '            End If
    '            If colArray(14).ToString <> "" Then
    '                If (colArray(14).ToString.Length > 50) Then Reason += "書面資料3攜回影本必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(15).ToString <> "" Then
    '                If (colArray(15).ToString.Length > 100) Then Reason += "書面資料3說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(16).ToString = "" Then
    '                Reason += "書面資料4必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(16)) = False Then Reason += "書面資料4必需為數字<BR>"
    '            End If
    '            If colArray(17).ToString <> "" Then
    '                If (colArray(17).ToString.Length > 50) Then Reason += "書面資料4攜回影本必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(18).ToString <> "" Then
    '                If (colArray(18).ToString.Length > 100) Then Reason += "書面資料4說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(19).ToString = "" Then
    '                Reason += "書面資料5必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(19)) = False Then Reason += "書面資料5必需為數字<BR>"
    '            End If
    '            If colArray(20).ToString <> "" Then
    '                If (colArray(20).ToString.Length > 50) Then Reason += "書面資料5攜回影本必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(21).ToString <> "" Then
    '                If (colArray(21).ToString.Length > 100) Then Reason += "書面資料5說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(22).ToString = "" Then
    '                Reason += "書面資料6必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(22)) = False Then Reason += "書面資料6必需為數字<BR>"
    '            End If
    '            If colArray(23).ToString <> "" Then
    '                If (colArray(23).ToString.Length > 50) Then Reason += "書面資料6攜回影本必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(24).ToString <> "" Then
    '                If (colArray(24).ToString.Length > 100) Then Reason += "書面資料6說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(25).ToString = "" Then
    '                Reason += "書面資料7必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(25)) = False Then Reason += "書面資料7必需為數字<BR>"
    '            End If
    '            If colArray(26).ToString <> "" Then
    '                If (colArray(26).ToString.Length > 50) Then Reason += "書面資料7攜回影本必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(27).ToString <> "" Then
    '                If (colArray(27).ToString.Length > 100) Then Reason += "書面資料7說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(28).ToString = "" Then
    '                Reason += "項次1_1有無週(月)課程表?必須填寫<Br>"
    '            Else
    '                Select Case colArray(28).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次1_1有無週(月)課程表?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(29).ToString = "" Then
    '                Reason += "項次1_2是否依課程表授課?必須填寫<Br>"
    '            Else
    '                Select Case colArray(29).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次1_2是否依課程表授課?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(30).ToString = "" Then
    '                Reason += "項次1_3課目或課題為何?必須填寫<Br>"
    '            Else
    '                If (colArray(30).ToString.Length > 50) Then Reason += "項次1_3課目或課題為何?必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(31).ToString <> "" Then
    '                If (colArray(31).ToString.Length > 500) Then Reason += "項次1處理情形必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(32).ToString <> "" Then
    '                If (colArray(32).ToString.Length > 500) Then Reason += "項次1備註說明必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(33).ToString = "" Then
    '                Reason += "項次2_1教學(訓練)日誌是否確實填寫?必須填寫<Br>"
    '            Else
    '                Select Case colArray(33).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次2_1教學(訓練)日誌是否確實填寫?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(34).ToString = "" Then
    '                Reason += "項次2_2有否按時呈主管核閱?必須填寫<Br>"
    '            Else
    '                Select Case colArray(34).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次2_2有否按時呈主管核閱?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(35).ToString <> "" Then
    '                If (colArray(35).ToString.Length > 500) Then Reason += "項次2處理情形必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(36).ToString <> "" Then
    '                If (colArray(36).ToString.Length > 500) Then Reason += "項次2備註必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(37).ToString = "" Then
    '                Reason += "項次3_1教師(職業訓練師)與助教姓名_是否與計畫相符?必須填寫<Br>"
    '            Else
    '                Select Case colArray(37).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次3_1教師(職業訓練師)與助教姓名_是否與計畫相符?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(38).ToString = "" Then
    '                Reason += "項次3_1講師必須填寫<Br>"
    '            Else
    '                If (colArray(38).ToString.Length > 50) Then Reason += "項次3_1講師必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(39).ToString = "" Then
    '                Reason += "項次3_1助教必須填寫<Br>"
    '            Else
    '                If (colArray(39).ToString.Length > 50) Then Reason += "項次3_1助教必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(40).ToString = "" Then
    '                Reason += "項次3_2學員學習情況是否良好?必須填寫<Br>"
    '            Else
    '                Select Case colArray(40).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次3_2學員學習情況是否良好?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(41).ToString <> "" Then
    '                If (colArray(41).ToString.Length > 500) Then Reason += "項次3處理情形必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(42).ToString <> "" Then
    '                If (colArray(42).ToString.Length > 500) Then Reason += "項次3備註必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(43).ToString = "" Then
    '                Reason += "項次4教學環境(教室或訓練工場)是否整齊?必須填寫<Br>"
    '            Else
    '                Select Case colArray(43).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次4教學環境(教室或訓練工場)是否整齊?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(44).ToString <> "" Then
    '                If (colArray(44).ToString.Length > 500) Then Reason += "項次4處理情形必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(45).ToString <> "" Then
    '                If (colArray(45).ToString.Length > 500) Then Reason += "項次4備註必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(46).ToString = "" Then
    '                Reason += "項次5有無具體輔導活動或其他事項?必須填寫<Br>"
    '            Else
    '                Select Case colArray(46).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次5有無具體輔導活動或其他事項?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(47).ToString <> "" Then
    '                If (colArray(47).ToString.Length > 500) Then Reason += "項次5處理情形必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(48).ToString <> "" Then
    '                If (colArray(48).ToString.Length > 500) Then Reason += "項次5備註必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(49).ToString = "" Then
    '                Reason += "項次6職業訓練生活津貼是否依規定申請並發放?必須填寫<Br>"
    '            Else
    '                Select Case colArray(49).ToString
    '                    Case "Y", "N"
    '                    Case Else
    '                        Reason += "項次6職業訓練生活津貼是否依規定申請並發放?只能是Y或者是N<BR>"
    '                End Select
    '            End If
    '            If colArray(50).ToString = "" Then
    '                Reason += "項次6_請領人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(50)) = False Then Reason += "項次6_請領人數必需為數字<BR>"
    '            End If
    '            If colArray(51).ToString = "" Then
    '                Reason += "項次6_無異狀人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(51)) = False Then Reason += "項次6_無異狀人數必需為數字<BR>"
    '            End If
    '            If colArray(52).ToString = "" Then
    '                Reason += "項次6_需追蹤人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(52)) = False Then
    '                    Reason += "項次6_需追蹤人數必需為數字<BR>"
    '                End If
    '            End If
    '            If colArray(53).ToString <> "" Then
    '                If (colArray(53).ToString.Length > 50) Then Reason += "項次6_需追蹤姓名必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(54).ToString <> "" Then
    '                If (colArray(54).ToString.Length > 500) Then Reason += "項次6備註必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(55).ToString <> "" Then
    '                If (colArray(55).ToString.Length > 500) Then Reason += "項次7訓學員反映意見及問題: 必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(56).ToString = "" Then
    '                Reason += "培訓單位人員姓名必須填寫<Br>"
    '            Else
    '                If (colArray(56).ToString.Length > 10) Then Reason += "培訓單位人員姓名必須小於等於10字字數<BR>"
    '            End If
    '            If colArray(57).ToString = "" Then
    '                Reason += "訪視人員姓名必須填寫<Br>"
    '            Else
    '                If (colArray(57).ToString.Length > 10) Then Reason += "訪視人員姓名必須小於等於10字字數<BR>"
    '            End If
    '        End If
    '        Return Reason
    '    ElseIf sm.UserInfo.Years = 2009 Then
    '        If colArray.Length <> 83 Then
    '            'Reason += "欄位數量不正確(應該為58個欄位)<BR>"
    '            Reason += "欄位對應有誤<BR>"
    '        Else
    '            If colArray(0).ToString = "" Then
    '                Reason += "班別(OCID)必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(0)) = False Then Reason += "班別(OCID)必需為數字<BR>"
    '            End If
    '            If colArray(1).ToString = "" Then
    '                Reason += "訪查日期必須填寫<Br>"
    '            Else
    '                If IsDate(colArray(1)) = False Then
    '                    Reason += "訪查日期必須是西元年格式(yyyy/MM/dd)<BR>"
    '                Else
    '                    If CDate(colArray(1)) < "1900/1/1" Or CDate(colArray(1)) > "2100/1/1" Then Reason += "訪查日期範圍有誤<BR>"
    '                End If
    '            End If
    '            If colArray(2).ToString = "" Then
    '                Reason += "核定人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(2)) = False Then Reason += "核定人數必需為數字<BR>"
    '            End If
    '            If colArray(3).ToString = "" Then
    '                Reason += "實到人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(3)) = False Then Reason += "實到人數必需為數字<BR>"
    '            End If
    '            If colArray(4).ToString = "" Then
    '                Reason += "請假人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(4)) = False Then Reason += "請假人數必需為數字<BR>"
    '            End If
    '            If colArray(5).ToString = "" Then
    '                Reason += "缺(曠)課人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(5)) = False Then Reason += "缺(曠)課人數必需為數字<BR>"
    '            End If
    '            If colArray(6).ToString = "" Then
    '                Reason += "退訓人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(6)) = False Then Reason += "退訓人數必需為數字<BR>"
    '            End If
    '            If colArray(7).ToString = "" Then
    '                Reason += "退訓人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(7)) = False Then Reason += "提前就業人數必需為數字<BR>"
    '            End If
    '            If colArray(8).ToString = "" Then
    '                Reason += "書面資料1必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(8)) = False Then Reason += "書面資料1必需為數字<BR>"
    '            End If
    '            If colArray(9).ToString <> "" Then
    '                If IsNumeric(colArray(9)) = False Then Reason += "佐證資料1選項必需為數字<BR>"
    '            End If
    '            If colArray(10).ToString <> "" Then
    '                If (colArray(10).ToString.Length > 50) Then Reason += "書面資料1如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(11).ToString <> "" Then
    '                If (colArray(11).ToString.Length > 100) Then Reason += "書面資料1說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(12).ToString = "" Then
    '                Reason += "書面資料2必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(12)) = False Then Reason += "書面資料2必需為數字<BR>"
    '            End If
    '            If colArray(13).ToString <> "" Then
    '                If IsNumeric(colArray(13)) = False Then Reason += "佐證資料2選項必需為數字<BR>"
    '            End If
    '            If colArray(14).ToString <> "" Then
    '                If (colArray(14).ToString.Length > 50) Then Reason += "書面資料2如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(15).ToString <> "" Then
    '                If (colArray(15).ToString.Length > 100) Then Reason += "書面資料2說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(16).ToString = "" Then
    '                Reason += "書面資料3必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(16)) = False Then Reason += "書面資料3必需為數字<BR>"
    '            End If
    '            If colArray(17).ToString <> "" Then
    '                If IsNumeric(colArray(17)) = False Then Reason += "佐證資料3選項必需為數字<BR>"
    '            End If
    '            If colArray(18).ToString <> "" Then
    '                If (colArray(18).ToString.Length > 50) Then Reason += "書面資料3如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(19).ToString <> "" Then
    '                If IsNumeric(colArray(19)) = False Then Reason += "書面資料3說明選項必需為數字<BR>"
    '            End If
    '            If colArray(20).ToString <> "" Then
    '                If (colArray(20).ToString.Length > 100) Then Reason += "書面資料3說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(21).ToString = "" Then
    '                Reason += "書面資料4必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(21)) = False Then Reason += "書面資料4必需為數字<BR>"
    '            End If
    '            If colArray(22).ToString <> "" Then
    '                If IsNumeric(colArray(22)) = False Then Reason += "佐證資料4選項必需為數字<BR>"
    '            End If
    '            If colArray(23).ToString <> "" Then
    '                If (colArray(23).ToString.Length > 50) Then Reason += "書面資料4如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(24).ToString <> "" Then
    '                If IsNumeric(colArray(24)) = False Then Reason += "書面資料4說明選項必需為數字<BR>"
    '            End If
    '            If colArray(25).ToString <> "" Then
    '                If (colArray(25).ToString.Length > 100) Then Reason += "書面資料4說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(26).ToString = "" Then
    '                Reason += "書面資料5必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(26)) = False Then Reason += "書面資料5必需為數字<BR>"
    '            End If
    '            If colArray(27).ToString <> "" Then
    '                If IsNumeric(colArray(27)) = False Then Reason += "佐證資料5選項必需為數字<BR>"
    '            End If
    '            If colArray(28).ToString <> "" Then
    '                If (colArray(28).ToString.Length > 50) Then Reason += "書面資料5如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(29).ToString <> "" Then
    '                If IsNumeric(colArray(29)) = False Then Reason += "書面資料5說明選項必需為數字<BR>"
    '            End If
    '            If colArray(30).ToString <> "" Then
    '                If (colArray(30).ToString.Length > 100) Then Reason += "書面資料5說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(31).ToString = "" Then
    '                Reason += "書面資料6必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(31)) = False Then Reason += "書面資料6必需為數字<BR>"
    '            End If
    '            If colArray(32).ToString <> "" Then
    '                If IsNumeric(colArray(32)) = False Then Reason += "佐證資料6選項必需為數字<BR>"
    '            End If
    '            If colArray(33).ToString <> "" Then
    '                If (colArray(33).ToString.Length > 50) Then Reason += "書面資料6如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(34).ToString <> "" Then
    '                If IsNumeric(colArray(34)) = False Then Reason += "書面資料6說明選項必需為數字<BR>"
    '            End If
    '            If colArray(35).ToString <> "" Then
    '                If (colArray(35).ToString.Length > 100) Then Reason += "書面資料6說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(36).ToString = "" Then
    '                Reason += "書面資料7必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(36)) = False Then Reason += "書面資料7必需為數字<BR>"
    '            End If
    '            If colArray(37).ToString <> "" Then
    '                If IsNumeric(colArray(37)) = False Then Reason += "佐證資料7選項必需為數字<BR>"
    '            End If
    '            If colArray(38).ToString <> "" Then
    '                If (colArray(38).ToString.Length > 50) Then Reason += "書面資料7如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(39).ToString <> "" Then
    '                If IsNumeric(colArray(39)) = False Then Reason += "書面資料7說明選項必需為數字<BR>"
    '            End If
    '            If colArray(40).ToString <> "" Then
    '                If (colArray(40).ToString.Length > 100) Then Reason += "書面資料7說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(41).ToString = "" Then
    '                Reason += "書面資料8必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(41)) = False Then Reason += "書面資料8必需為數字<BR>"
    '            End If
    '            If colArray(42).ToString <> "" Then
    '                If IsNumeric(colArray(42)) = False Then Reason += "佐證資料8選項必需為數字<BR>"
    '            End If
    '            If colArray(43).ToString <> "" Then
    '                If (colArray(43).ToString.Length > 50) Then Reason += "書面資料8如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(44).ToString <> "" Then
    '                If IsNumeric(colArray(44)) = False Then Reason += "書面資料8說明選項必需為數字<BR>"
    '            End If
    '            If colArray(45).ToString <> "" Then
    '                If (colArray(45).ToString.Length > 100) Then Reason += "書面資料8說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(46).ToString = "" Then
    '                Reason += "有無週(月)課程表?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(46)) = False Then Reason += "有無週(月)課程表?必需為數字<BR>"
    '            End If
    '            If colArray(47).ToString = "" Then
    '                Reason += "是否依課程表授課?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(47)) = False Then Reason += "是否依課程表授課?必需為數字<BR>"
    '            End If
    '            If colArray(48).ToString = "" Then
    '                Reason += "課目或課題為何?必須填寫<Br>"
    '            Else
    '                If (colArray(48).ToString.Length > 50) Then Reason += "課目或課題為何?必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(49).ToString = "" Then
    '                Reason += "教師(職業訓練師)與助教姓名_是否與計畫相符?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(49)) = False Then Reason += "教師(職業訓練師)與助教姓名_是否與計畫相符?必需為數字<BR>"
    '            End If
    '            If colArray(50).ToString = "" Then
    '                Reason += "教師必須填寫<Br>"
    '            Else
    '                If (colArray(50).ToString.Length > 50) Then Reason += "教師必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(51).ToString <> "" Then
    '                If (colArray(51).ToString.Length > 50) Then Reason += "助教必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(52).ToString <> "" Then
    '                If (colArray(52).ToString.Length > 500) Then Reason += "處理情形1必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(53).ToString <> "" Then
    '                If (colArray(53).ToString.Length > 500) Then Reason += "備註1必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(54).ToString = "" Then
    '                Reason += "有無書籍(講義)領用表?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(54)) = False Then Reason += "有無書籍(講義)領用表?必需為數字<BR>"
    '            End If
    '            If colArray(55).ToString = "" Then
    '                Reason += "有無材料領用表?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(55)) = False Then Reason += "有無材料領用表?必需為數字<BR>"
    '            End If
    '            If colArray(56).ToString = "" Then
    '                Reason += "訓練設施設備是否依契約提供學員使用?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(56)) = False Then Reason += "訓練設施設備是否依契約提供學員使用?必需為數字<BR>"
    '            End If
    '            If colArray(57).ToString <> "" Then
    '                If (colArray(57).ToString.Length > 500) Then Reason += "處理情形2必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(58).ToString <> "" Then
    '                If (colArray(58).ToString.Length > 500) Then Reason += "備註2必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(59).ToString = "" Then
    '                Reason += "教學(訓練)日誌是否確實填寫?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(59)) = False Then Reason += "教學(訓練)日誌是否確實填寫?必需為數字<BR>"
    '            End If
    '            If colArray(60).ToString = "" Then
    '                Reason += "有否按時呈主管核閱?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(60)) = False Then Reason += "有否按時呈主管核閱?必需為數字<BR>"
    '            End If
    '            If colArray(61).ToString = "" Then
    '                Reason += "學員生活就業輔導與管理機制是否依契約規範辦理?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(61)) = False Then Reason += "學員生活就業輔導與管理機制是否依契約規範辦理?必需為數字<BR>"
    '            End If
    '            If colArray(62).ToString = "" Then
    '                Reason += "是否依契約規範提供學員問題反應申訴管道?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(62)) = False Then Reason += "是否依契約規範提供學員問題反應申訴管道?必需為數字<BR>"
    '            End If
    '            If colArray(63).ToString = "" Then
    '                Reason += "是否為參訓學員辦理勞工保險加退保?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(63)) = False Then Reason += "是否為參訓學員辦理勞工保險加退保?必需為數字<BR>"
    '            End If
    '            If colArray(64).ToString = "" Then
    '                Reason += "是否依契約規範公告學員權益教務管理狀況義務或編製參訓學員服務手冊?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(64)) = False Then Reason += "是否依契約規範公告學員權益教務管理狀況義務或編製參訓學員服務手冊?必需為數字<BR>"
    '            End If
    '            If colArray(65).ToString <> "" Then
    '                If (colArray(65).ToString.Length > 500) Then Reason += "處理情形3必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(66).ToString <> "" Then
    '                If (colArray(66).ToString.Length > 500) Then Reason += "備註3必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(67).ToString = "" Then
    '                Reason += "有無自費參訓學員?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(67)) = False Then Reason += "有無自費參訓學員?必需為數字<BR>"
    '            End If
    '            If colArray(67).ToString = 1 And colArray(68).ToString = "" Then Reason += "幾人?必須填寫<Br>"
    '            If colArray(68).ToString <> "" Then
    '                If IsNumeric(colArray(68)) = False Then Reason += "幾人?必需為數字<BR>"
    '            End If
    '            If colArray(69).ToString = "" Then
    '                Reason += "訓練單位是否繳交主辦單位?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(69)) = False Then Reason += "訓練單位是否繳交主辦單位?必需為數字<BR>"
    '            End If
    '            If colArray(70).ToString = "" Then
    '                Reason += "職業訓練生活津貼是否依規定申請並核發?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(70)) = False Then Reason += "職業訓練生活津貼是否依規定申請並核發?必需為數字<BR>"
    '            End If
    '            If colArray(71).ToString = "" Then
    '                Reason += "培訓單位是否巧立名目強制收取費用?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(71)) = False Then Reason += "培訓單位是否巧立名目強制收取費用?必需為數字<BR>"
    '            End If
    '            If colArray(72).ToString <> "" Then
    '                If (colArray(72).ToString.Length > 500) Then Reason += "處理情形4必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(73).ToString <> "" Then
    '                If (colArray(73).ToString.Length > 500) Then Reason += "備註4必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(74).ToString = "" Then
    '                Reason += "職業訓練機構是否依規定懸掛設立許可證書?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(74)) = False Then Reason += "職業訓練機構是否依規定懸掛設立許可證書?必需為數字<BR>"
    '            End If
    '            If colArray(75).ToString <> "" Then
    '                If (colArray(75).ToString.Length > 500) Then Reason += "處理情形5必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(76).ToString <> "" Then
    '                If (colArray(76).ToString.Length > 500) Then Reason += "備註5必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(77).ToString <> "" Then
    '                If (colArray(77).ToString.Length > 500) Then Reason += "訓學員反映意見及問題:必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(78).ToString <> "" Then
    '                If (colArray(78).ToString.Length > 500) Then Reason += "綜合建議:必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(79).ToString <> "" Then
    '                If IsNumeric(colArray(79)) = False Then Reason += "缺失處理?必需為數字<BR>"
    '            End If
    '            If colArray(80).ToString <> "" Then
    '                If (colArray(80).ToString.Length > 500) Then Reason += "其他說明內容:必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(81).ToString = "" Then
    '                Reason += "培訓單位人員姓名必須填寫<Br>"
    '            Else
    '                If (colArray(81).ToString.Length > 10) Then Reason += "培訓單位人員姓名必須小於等於10字字數<BR>"
    '            End If
    '            If colArray(82).ToString = "" Then
    '                Reason += "訪視人員姓名必須填寫<Br>"
    '            Else
    '                If (colArray(82).ToString.Length > 10) Then Reason += "訪視人員姓名必須小於等於10字字數<BR>"
    '            End If
    '        End If
    '        Return Reason
    '    ElseIf sm.UserInfo.Years >= 2010 Then
    '        If colArray.Length <> 61 Then
    '            'Reason += "欄位數量不正確(應該為58個欄位)<BR>"
    '            Reason += "欄位對應有誤<BR>"
    '        Else
    '            If colArray(0).ToString = "" Then
    '                Reason += "班別(OCID)必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(0)) = False Then Reason += "班別(OCID)必需為數字<BR>"
    '            End If
    '            If colArray(1).ToString = "" Then
    '                Reason += "訪查日期必須填寫<Br>"
    '            Else
    '                If IsDate(colArray(1)) = False Then
    '                    Reason += "訪查日期必須是西元年格式(yyyy/MM/dd)<BR>"
    '                Else
    '                    If CDate(colArray(1)) < "1900/1/1" Or CDate(colArray(1)) > "2100/1/1" Then Reason += "訪查日期範圍有誤<BR>"
    '                End If
    '            End If
    '            If colArray(2).ToString = "" Then
    '                Reason += "核定人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(2)) = False Then Reason += "核定人數必需為數字<BR>"
    '            End If
    '            If colArray(3).ToString = "" Then
    '                Reason += "實到人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(3)) = False Then Reason += "實到人數必需為數字<BR>"
    '            End If
    '            If colArray(4).ToString = "" Then
    '                Reason += "請假人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(4)) = False Then Reason += "請假人數必需為數字<BR>"
    '            End If
    '            If colArray(5).ToString = "" Then
    '                Reason += "缺(曠)課人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(5)) = False Then Reason += "缺(曠)課人數必需為數字<BR>"
    '            End If
    '            If colArray(6).ToString = "" Then
    '                Reason += "退訓人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(6)) = False Then Reason += "退訓人數必需為數字<BR>"
    '            End If
    '            If colArray(7).ToString = "" Then
    '                Reason += "退訓人數必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(7)) = False Then Reason += "提前就業人數必需為數字<BR>"
    '            End If
    '            If colArray(8).ToString = "" Then
    '                Reason += "書面資料1必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(8)) = False Then Reason += "書面資料1必需為數字<BR>"
    '            End If
    '            If colArray(9).ToString <> "" Then
    '                If IsNumeric(colArray(9)) = False Then Reason += "佐證資料1選項必需為數字<BR>"
    '            End If
    '            If colArray(10).ToString <> "" Then
    '                If (colArray(10).ToString.Length > 50) Then Reason += "書面資料1如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(11).ToString <> "" Then
    '                If (colArray(11).ToString.Length > 100) Then Reason += "書面資料1說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(12).ToString = "" Then
    '                Reason += "書面資料2必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(12)) = False Then Reason += "書面資料2必需為數字<BR>"
    '            End If
    '            If colArray(13).ToString <> "" Then
    '                If IsNumeric(colArray(13)) = False Then Reason += "佐證資料2選項必需為數字<BR>"
    '            End If
    '            If colArray(14).ToString <> "" Then
    '                If (colArray(14).ToString.Length > 50) Then Reason += "書面資料2如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(15).ToString <> "" Then
    '                If (colArray(15).ToString.Length > 100) Then Reason += "書面資料2說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(16).ToString = "" Then
    '                Reason += "書面資料3必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(16)) = False Then Reason += "書面資料3必需為數字<BR>"
    '            End If
    '            If colArray(17).ToString <> "" Then
    '                If IsNumeric(colArray(17)) = False Then Reason += "佐證資料3選項必需為數字<BR>"
    '            End If
    '            If colArray(18).ToString <> "" Then
    '                If (colArray(18).ToString.Length > 50) Then Reason += "書面資料3如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(19).ToString <> "" Then
    '                If IsNumeric(colArray(19)) = False Then Reason += "書面資料3說明選項必需為數字<BR>"
    '            End If
    '            If colArray(20).ToString <> "" Then
    '                If (colArray(20).ToString.Length > 100) Then Reason += "書面資料3說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(21).ToString = "" Then
    '                Reason += "書面資料4必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(21)) = False Then Reason += "書面資料4必需為數字<BR>"
    '            End If
    '            If colArray(22).ToString <> "" Then
    '                If IsNumeric(colArray(22)) = False Then Reason += "佐證資料4選項必需為數字<BR>"
    '            End If
    '            If colArray(23).ToString <> "" Then
    '                If (colArray(23).ToString.Length > 50) Then Reason += "書面資料4如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(24).ToString <> "" Then
    '                If IsNumeric(colArray(24)) = False Then Reason += "書面資料4說明選項必需為數字<BR>"
    '            End If
    '            If colArray(25).ToString <> "" Then
    '                If (colArray(25).ToString.Length > 100) Then Reason += "書面資料4說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(26).ToString = "" Then
    '                Reason += "書面資料5必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(26)) = False Then Reason += "書面資料5必需為數字<BR>"
    '            End If
    '            If colArray(27).ToString <> "" Then
    '                If IsNumeric(colArray(27)) = False Then Reason += "佐證資料5選項必需為數字<BR>"
    '            End If
    '            If colArray(28).ToString <> "" Then
    '                If (colArray(28).ToString.Length > 50) Then Reason += "書面資料5如附件必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(29).ToString <> "" Then
    '                If IsNumeric(colArray(29)) = False Then Reason += "書面資料5說明選項必需為數字<BR>"
    '            End If
    '            If colArray(30).ToString <> "" Then
    '                If (colArray(30).ToString.Length > 100) Then Reason += "書面資料5說明必須小於等於100字字數<BR>"
    '            End If
    '            If colArray(31).ToString = "" Then
    '                Reason += "有無週(月)課程表?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(31)) = False Then Reason += "有無週(月)課程表?必需為數字<BR>"
    '            End If
    '            If colArray(32).ToString = "" Then
    '                Reason += "是否依課程表授課?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(32)) = False Then Reason += "是否依課程表授課?必需為數字<BR>"
    '            End If
    '            If colArray(33).ToString = "" Then
    '                Reason += "課目或課題為何?必須填寫<Br>"
    '            Else
    '                If (colArray(33).ToString.Length > 50) Then Reason += "課目或課題為何?必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(34).ToString = "" Then
    '                Reason += "教師(職業訓練師)與助教姓名_是否與計畫相符?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(34)) = False Then Reason += "教師(職業訓練師)與助教姓名_是否與計畫相符?必需為數字<BR>"
    '            End If
    '            If colArray(35).ToString = "" Then
    '                Reason += "教師必須填寫<Br>"
    '            Else
    '                If (colArray(35).ToString.Length > 50) Then Reason += "教師必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(36).ToString <> "" Then
    '                If (colArray(36).ToString.Length > 50) Then Reason += "助教必須小於等於50字字數<BR>"
    '            End If
    '            If colArray(37).ToString <> "" Then
    '                If (colArray(37).ToString.Length > 500) Then Reason += "處理情形1必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(38).ToString <> "" Then
    '                If (colArray(38).ToString.Length > 500) Then Reason += "備註1必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(39).ToString = "" Then
    '                Reason += "有無書籍(講義)領用表?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(39)) = False Then Reason += "有無書籍(講義)領用表?必需為數字<BR>"
    '            End If
    '            If colArray(40).ToString = "" Then
    '                Reason += "有無材料領用表?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(40)) = False Then Reason += "有無材料領用表?必需為數字<BR>"
    '            End If
    '            If colArray(41).ToString = "" Then
    '                Reason += "訓練設施設備是否依契約提供學員使用?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(41)) = False Then Reason += "訓練設施設備是否依契約提供學員使用?必需為數字<BR>"
    '            End If
    '            If colArray(42).ToString <> "" Then
    '                If (colArray(42).ToString.Length > 500) Then Reason += "處理情形2必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(43).ToString <> "" Then
    '                If (colArray(43).ToString.Length > 500) Then Reason += "備註2必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(44).ToString = "" Then
    '                Reason += "教學(訓練)日誌是否確實填寫?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(44)) = False Then Reason += "教學(訓練)日誌是否確實填寫?必需為數字<BR>"
    '            End If
    '            If colArray(45).ToString = "" Then
    '                Reason += "有否按時呈主管核閱?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(45)) = False Then Reason += "有否按時呈主管核閱?必需為數字<BR>"
    '            End If
    '            If colArray(46).ToString = "" Then
    '                Reason += "學員生活就業輔導與管理機制是否依契約規範辦理?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(46)) = False Then Reason += "學員生活就業輔導與管理機制是否依契約規範辦理?必需為數字<BR>"
    '            End If
    '            If colArray(47).ToString = "" Then
    '                Reason += "是否依契約規範提供學員問題反應申訴管道?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(47)) = False Then Reason += "是否依契約規範提供學員問題反應申訴管道?必需為數字<BR>"
    '            End If
    '            If colArray(48).ToString = "" Then
    '                Reason += "是否為參訓學員辦理勞工保險加退保?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(48)) = False Then Reason += "是否為參訓學員辦理勞工保險加退保?必需為數字<BR>"
    '            End If
    '            If colArray(49).ToString = "" Then
    '                Reason += "是否依契約規範公告學員權益教務管理狀況義務或編製參訓學員服務手冊?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(49)) = False Then Reason += "是否依契約規範公告學員權益教務管理狀況義務或編製參訓學員服務手冊?必需為數字<BR>"
    '            End If
    '            If colArray(50).ToString <> "" Then
    '                If (colArray(50).ToString.Length > 500) Then Reason += "處理情形3必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(51).ToString <> "" Then
    '                If (colArray(51).ToString.Length > 500) Then Reason += "備註3必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(52).ToString = "" Then
    '                Reason += "職業訓練生活津貼是否依規定申請並核發?必須填寫<Br>"
    '            Else
    '                If IsNumeric(colArray(52)) = False Then Reason += "職業訓練生活津貼是否依規定申請並核發?必需為數字<BR>"
    '            End If
    '            If colArray(53).ToString <> "" Then
    '                If (colArray(53).ToString.Length > 500) Then Reason += "處理情形4必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(54).ToString <> "" Then
    '                If (colArray(54).ToString.Length > 500) Then Reason += "備註4必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(55).ToString <> "" Then
    '                If (colArray(55).ToString.Length > 500) Then Reason += "訓學員反映意見及問題:必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(56).ToString <> "" Then
    '                If (colArray(56).ToString.Length > 500) Then Reason += "綜合建議:必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(57).ToString <> "" Then
    '                If IsNumeric(colArray(57)) = False Then Reason += "缺失處理?必需為數字<BR>"
    '            End If
    '            If colArray(58).ToString <> "" Then
    '                If (colArray(58).ToString.Length > 500) Then Reason += "其他說明內容:必須小於等於500字字數<BR>"
    '            End If
    '            If colArray(59).ToString = "" Then
    '                Reason += "培訓單位人員姓名必須填寫<Br>"
    '            Else
    '                If (colArray(59).ToString.Length > 10) Then Reason += "培訓單位人員姓名必須小於等於10字字數<BR>"
    '            End If
    '            If colArray(60).ToString = "" Then
    '                Reason += "訪視人員姓名必須填寫<Br>"
    '            Else
    '                If (colArray(60).ToString.Length > 10) Then Reason += "訪視人員姓名必須小於等於10字字數<BR>"
    '            End If
    '        End If
    '    End If
    '    Return Reason
    'End Function

#End Region

End Class