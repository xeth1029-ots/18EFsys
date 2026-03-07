Partial Class CP_01_008
    Inherits AuthBasePage

    'Const cst_printFN1 As String = "CP_01_006_1" '列印空白抽訪紀錄表
    Const cst_printFN4 As String = "CP_01_006_2" '2019'列印空白抽訪紀錄表

    Const cst_printFN2 As String = "CP_01_007_1" '電話抽訪
    Const cst_printFN3 As String = "CP_01_008" '實地抽訪

    Dim sURL_CP01007ADD As String = ""
    Dim sURL_CP01006ADD As String = ""
    'Const cst_CP_01_006_add_aspx_new As String = "CP_01_006_add9.aspx?ID="

    'Dim FunDr As DataRow
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        sURL_CP01007ADD = String.Concat(TIMS.URL_CP01007ADD, rqMID)
        sURL_CP01006ADD = String.Concat("CP_01_006_add9.aspx?ID=", rqMID)

        iPYNum = TIMS.sUtl_GetPYNum(Me)

        If Not IsPostBack Then
            Query.Attributes("onclick") = "return CheckSearch();"
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DataGridTable.Visible = False
            Save.Visible = False
            Print.Visible = False
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

#Region "(No Use)"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)
        '            If FunDr("Sech") = 1 Then
        '                Query.Enabled = True
        '            Else
        '                Query.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

#End Region

        If Not IsPostBack Then
            Call Utl_UseSessionStr1(sender, e)

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button13_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Utl_UseSessionStr1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Session("SearchStr") Is Nothing Then Return
        Dim s_SearchStr1 As String = Convert.ToString(Session("SearchStr"))
        Session("SearchStr") = Nothing

        Dim MyValue As String = ""

        center.Text = TIMS.GetMyValue(s_SearchStr1, "center")
        RIDValue.Value = TIMS.GetMyValue(s_SearchStr1, "RIDValue")
        TMID1.Text = TIMS.GetMyValue(s_SearchStr1, "TMID1")
        OCID1.Text = TIMS.GetMyValue(s_SearchStr1, "OCID1")
        TMIDValue1.Value = TIMS.GetMyValue(s_SearchStr1, "TMIDValue1")
        OCIDValue1.Value = TIMS.GetMyValue(s_SearchStr1, "OCIDValue1")
        STDate1.Text = TIMS.GetMyValue(s_SearchStr1, "STDate1")
        STDate2.Text = TIMS.GetMyValue(s_SearchStr1, "STDate2")
        FTDate1.Text = TIMS.GetMyValue(s_SearchStr1, "FTDate1")
        FTDate2.Text = TIMS.GetMyValue(s_SearchStr1, "FTDate2")

        MyValue = TIMS.GetMyValue(s_SearchStr1, "PageIndex")
        If MyValue <> "" AndAlso IsNumeric(MyValue) Then
            MyValue = CInt(MyValue)
            PageControler1.PageIndex = MyValue
        End If
        MyValue = TIMS.GetMyValue(s_SearchStr1, "Query")
        If MyValue = "true" Then Call Query_Click(sender, e)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim labSEQNO As Label = e.Item.FindControl("labSEQNO")
                Dim labSTDate As Label = e.Item.FindControl("labSTDate")
                Dim labFTDate As Label = e.Item.FindControl("labFTDate")
                Dim ViewRecord As Label = e.Item.FindControl("ViewRecord")
                Dim VisitorName As TextBox = e.Item.FindControl("VisitorName")
                Dim ExpectDate As TextBox = e.Item.FindControl("ExpectDate")
                Dim Img1 As HtmlImage = e.Item.FindControl("Img1")
                Dim AddUnexpectVisitor As Button = e.Item.FindControl("AddUnexpectVisitor")
                Dim AddUnExpectTel As Button = e.Item.FindControl("AddUnExpectTel")
                Dim PrintUnexpectVisitor As Button = e.Item.FindControl("PrintUnexpectVisitor")
                Dim PrintUnExpectTel As Button = e.Item.FindControl("PrintUnExpectTel")
                ''Dim SeqNo As Label = e.Item.FindControl("SeqNo")
                Dim hid_OCID As HtmlInputHidden = e.Item.FindControl("hid_OCID")
                Dim hid_ceSEQNO As HtmlInputHidden = e.Item.FindControl("hid_ceSEQNO")

                hid_OCID.Value = Convert.ToString(drv("OCID"))
                hid_ceSEQNO.Value = If(Convert.ToString(drv("ceSEQNO")) <> "", Convert.ToString(drv("ceSEQNO")), "1")

                labSEQNO.Text = TIMS.Get_DGSeqNo(sender, e)
                VisitorName.Text = drv("VisitorName").ToString
                '開訓日/ '結訓日
                labSTDate.Text = If(flag_ROC, TIMS.Cdate17(drv("STDate")), TIMS.Cdate3(drv("STDate")))
                labFTDate.Text = If(flag_ROC, TIMS.Cdate17(drv("FTDate")), TIMS.Cdate3(drv("FTDate")))
                ExpectDate.Text = If(flag_ROC, TIMS.Cdate17(drv("ExpectDate")), TIMS.Cdate3(drv("ExpectDate")))
                Img1.Attributes("onclick") = "javascript:show_calendar('" & ExpectDate.ClientID & "','','','CY/MM/DD');"

                ViewRecord.Text = Get_Record_Url(drv("OCID").ToString)

                '新增抽訪紀錄
                'Dim rqMID As String = TIMS.Get_MRqID(Me)
                'Dim sUrl1 As String = "CP_01_006_add.aspx?ID=" & rqMID
                'If iPYNum >= 3 Then sUrl1 = "CP_01_006_add8.aspx?ID=" & rqMID
                Dim rqMID As String = TIMS.Get_MRqID(Me)
                Dim sUrl1 As String = sURL_CP01006ADD 'cst_CP_01_006_add_aspx_new & rqMID
                Dim sUrl2 As String = sURL_CP01007ADD '"CP_01_007_add.aspx?ID=" & rqMID
                AddUnexpectVisitor.CommandArgument = String.Concat(sUrl1, "&OCID=", drv("OCID"), "&State=Add&Type=CV")
                AddUnExpectTel.CommandArgument = String.Concat(sUrl2, "&OCID=", drv("OCID"), "&State=Add&Type=CT")

                '列印空白抽訪紀錄表
                'Dim prtFilename As String = cst_printFN1
                'If TIMS.GetReportQueryPath() = TIMS.cst_Report_TEST Then prtFilename = cst_printFN4 '"CP_01_006_C2"
                'Dim prtFilename As String = cst_printFN4 '"CP_01_006_C2"
                PrintUnexpectVisitor.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN4, "OCID=" & drv("OCID") & "")
                Dim vMyValue2 As String = String.Concat("OCID=", drv("OCID"), "&RID=", drv("RID"), "&PLANID=", drv("PLANID"))
                PrintUnExpectTel.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN2, vMyValue2)

        End Select

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        GetSearchStr()
        'Session("_SearchStr") = Me.ViewState("_SearchStr")
        Select Case e.CommandName
            Case "AddUV"
                TIMS.Utl_Redirect1(Me, e.CommandArgument)
            Case "AddUT"
                TIMS.Utl_Redirect1(Me, e.CommandArgument)
        End Select

    End Sub

    Sub GetSearchStr()
        Dim s_SearchStr1 As String = ""
        TIMS.SetMyValue(s_SearchStr1, "center", center.Text)
        TIMS.SetMyValue(s_SearchStr1, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(s_SearchStr1, "TMID1", TMID1.Text)
        TIMS.SetMyValue(s_SearchStr1, "OCID1", OCID1.Text)
        TIMS.SetMyValue(s_SearchStr1, "TMIDValue1", TMIDValue1.Value)
        TIMS.SetMyValue(s_SearchStr1, "OCIDValue1", OCIDValue1.Value)
        TIMS.SetMyValue(s_SearchStr1, "STDate1", STDate1.Text)
        TIMS.SetMyValue(s_SearchStr1, "STDate2", STDate2.Text)
        TIMS.SetMyValue(s_SearchStr1, "FTDate1", FTDate1.Text)
        TIMS.SetMyValue(s_SearchStr1, "FTDate2", FTDate2.Text)
        TIMS.SetMyValue(s_SearchStr1, "PageIndex", (DataGrid1.CurrentPageIndex + 1))
        TIMS.SetMyValue(s_SearchStr1, "Query", If(DataGrid1.Visible, "true", "false"))
        Session("SearchStr") = s_SearchStr1

    End Sub

    Function Get_Record_Url(ByVal OCID As String) As String
        Dim sql As String = ""
        Dim sqlstr As String = ""
        Dim revalue As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        '**by Milor 20080429--將原本用ApplyDate與SeqNO組合URL名稱，改為同一個ApplyDate名稱堆疊----start
        sql = ""
        sql &= " select OCID,ApplyDate,SeqNo,'CV' as Type" & vbCrLf
        sql &= " ,convert(varchar, ApplyDate, 112) as AppDate" & vbCrLf
        sql &= " ,'CP_01_006_add9' as url "
        'sql += "'<SPAN class=newlink><A class=newlink target=_blank href=""CP_01_006_add.aspx?ID=" & Request("ID") & "&OCID=" & OCID & "&State=View&Type=CV"
        'sql += "&SeqNo=@SeqNo""><font color=blue>'+ CONVERT(varchar(10) ,ApplyDate,112)+'-'+convert(varchar,SeqNo) "
        'sql += "+'</font></A></SPAN>' "
        'sql += "as url "
        sql &= " FROM dbo.CLASS_UNEXPECTVISITOR" & vbCrLf
        sql &= " where OCID='" & OCID & "' " & vbCrLf
        sql &= " union all " & vbCrLf
        sql &= " select OCID,ApplyDate,SeqNo,'CT' as Type" & vbCrLf
        sql &= " ,convert(varchar, ApplyDate, 112) as AppDate" & vbCrLf
        sql &= " ,'CP_01_007_add' as url "
        'sql += "'<SPAN class=newlink><A class=newlink target=_blank href=""CP_01_007_add.aspx?ID=" & Request("ID") & "&OCID=" & OCID & "&State=View&Type=CT"
        'sql += "&SeqNo=@SeqNo""><font color=blue>'+ CONVERT(varchar(10) ,ApplyDate,112)+'-'+convert(varchar,SeqNo) "
        'sql += "+'</font></A></SPAN>' "
        'sql += "as url "
        sql &= " FROM dbo.CLASS_UNEXPECTTEL" & vbCrLf
        sql &= " where OCID='" & OCID & "' " & vbCrLf
        sql &= " ORDER BY ApplyDate,SeqNo "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Return ""
        'If dt IsNot Nothing Then End If

        Dim urlwd As String = ""
        Dim fword As String = ""
        Dim cnt As Integer = 0
        Dim appDate As String = ""
        For i As Integer = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            urlwd = ""
            If Not dt.Rows(i).Item("AppDate") Is Nothing Then   '沒有調查過就步做任何處理，雖說這種情況應該不會存在
                appDate = dt.Rows(i).Item("AppDate").ToString
                appDate = (CInt(appDate.Substring(0, 4)) - 1911).ToString + appDate.Substring(4)

                If dt.Rows(i).Item("AppDate").ToString = fword Then '調查日期沒有重複的話，就不用加上-區分
                    urlwd = "<SPAN class=newlink><A class=newlink target=_self href=""" + dt.Rows(i).Item("url").ToString + ".aspx?ID=" & Request("ID") & "&OCID=" & OCID & "&State=View&Type=" + dt.Rows(i).Item("Type").ToString + "&SeqNo=@SeqNo""><font color=blue>" + appDate + "-" + CStr(cnt + 1) + "</font></A></SPAN>"
                    cnt += 1
                Else
                    urlwd = "<SPAN class=newlink><A class=newlink target=_self href=""" + dt.Rows(i).Item("url").ToString + ".aspx?ID=" & Request("ID") & "&OCID=" & OCID & "&State=View&Type=" + dt.Rows(i).Item("Type").ToString + "&SeqNo=@SeqNo""><font color=blue>" + appDate + "</font></A></SPAN>"
                    cnt = 0
                End If
            End If
            urlwd = Replace(urlwd, "@SeqNo", dr("SeqNo"))
            revalue &= urlwd & "<br>" & vbCrLf
            fword = dt.Rows(i).Item("AppDate").ToString
        Next

        '**by Milor 20080429----end
        Return revalue
    End Function

    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Dim Relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT a.OCID,a.DISTID,a.RID" & vbCrLf
        sql &= " ,a.DISTNAME,a.ORGNAME,a.ORGTYPENAME" & vbCrLf
        sql &= " ,a.CLASSCNAME,a.CYCLTYPE,a.STDATE,a.FTDATE" & vbCrLf
        sql &= " ,a.IsBusiness,a.PointYN" & vbCrLf
        sql &= " ,a.PlanID,a.ComIDNO,a.SeqNo" & vbCrLf
        sql &= " ,a.TaddressZip ,a.TaddressZip6W,a.TAddress" & vbCrLf
        sql &= " FROM VIEW2 a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND a.NotOpen='N'" & vbCrLf
        'sql &= " AND a.IsSuccess='Y'" & vbCrLf
        'sql &= " and a.TPLANID='28'" & vbCrLf
        'sql &= " and a.YEARS='2021'" & vbCrLf
        'sql &= " AND a.DISTID ='001'" & vbCrLf
        sql &= " AND a.NotOpen='N' " & vbCrLf
        sql &= " AND a.IsSuccess='Y'" & vbCrLf
        If OCIDValue1.Value <> "" Then
            sql &= " AND a.OCID = @OCID " & vbCrLf
            parms.Add("OCID", OCIDValue1.Value)
        End If
        If RIDValue.Value <> "" Then
            sql &= " AND a.RID LIKE @RID " & vbCrLf
            parms.Add("RID", RIDValue.Value & "%")
        Else
            sql &= " AND a.RID = @RID " & vbCrLf
            parms.Add("RID", sm.UserInfo.RID)
        End If
        If sm.UserInfo.DistID <> "000" Then
            sql &= " AND a.PlanID = @PlanID " & vbCrLf
            parms.Add("PlanID", sm.UserInfo.PlanID)
        Else
            sql &= " AND a.TPlanID = @TPlanID " & vbCrLf
            sql &= " AND a.Years = @Years " & vbCrLf
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("Years", sm.UserInfo.Years)
        End If
        If STDate1.Text <> "" Then
            sql &= " AND a.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", If(flag_ROC, TIMS.Cdate18(STDate1.Text), STDate1.Text))  'edit，by:20181018
        End If
        If STDate2.Text <> "" Then
            sql &= " AND a.STDate <= @STDate2 " & vbCrLf
            parms.Add("STDate2", If(flag_ROC, TIMS.Cdate18(STDate2.Text), STDate2.Text))  'edit，by:20181018
        End If
        If FTDate1.Text <> "" Then
            sql &= " AND a.FTDate >= @FTDate1 " & vbCrLf
            parms.Add("FTDate1", If(flag_ROC, TIMS.Cdate18(FTDate1.Text), FTDate1.Text))  'edit，by:20181018
        End If
        If FTDate2.Text <> "" Then
            sql &= " AND a.FTDate <= @FTDate2 " & vbCrLf
            parms.Add("FTDate2", If(flag_ROC, TIMS.Cdate18(FTDate2.Text), FTDate2.Text))  'edit，by:20181018
        End If
        sql &= " )" & vbCrLf

        sql &= " ,WV2 AS (" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,MAX(b.SeqNo) SEQNO" & vbCrLf
        sql &= " ,COUNT(CASE WHEN b.LItem2='1' THEN 1 END) cTIMES" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf
        sql &= " JOIN CLASS_UNEXPECTVISITOR b on b.OCID =a.OCID" & vbCrLf
        sql &= " GROUP BY a.OCID )" & vbCrLf

        sql &= " ,WT2 AS (" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,MAX(b.SeqNo) SEQNO" & vbCrLf
        sql &= " ,COUNT(CASE WHEN b.Item10='2' THEN 1 END) cTIMES" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf
        sql &= " JOIN CLASS_UNEXPECTTEL b on b.OCID =a.OCID" & vbCrLf
        sql &= " GROUP BY a.OCID )" & vbCrLf

        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,a.RID,a.PLANID" & vbCrLf
        sql &= " ,a.DISTNAME" & vbCrLf
        sql &= " ,a.ORGNAME" & vbCrLf
        sql &= " ,a.ORGTYPENAME" & vbCrLf
        sql &= " ,a.CYCLTYPE" & vbCrLf
        sql &= " ,a.STDATE,a.FTDATE" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,dbo.DECODE(a.IsBusiness,'Y','是','否') ISBUSINESS" & vbCrLf
        sql &= " ,dbo.DECODE(a.PointYN,'Y','是','否') POINTYN" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_ONCLASS(a.PlanID,a.ComIDNO,a.SeqNo,'WEEKTIME') WEEKS" & vbCrLf
        sql &= " ,concat(dbo.FN_GET_ZIPCODE(a.TaddressZip ,a.TaddressZip6W),a.TAddress) ADDRESS" & vbCrLf
        sql &= " ,ISNULL(cv.SeqNo,0)+ISNULL(ct.SeqNo,0) SEQNO" & vbCrLf
        sql &= " ,ISNULL(cv.cTIMES,0)+ISNULL(ct.cTIMES,0) UNEXPECTTIMES" & vbCrLf
        sql &= " ,ce.VISITORNAME" & vbCrLf
        sql &= " ,ce.EXPECTDATE" & vbCrLf
        sql &= " ,ce.SeqNo ceSEQNO" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf
        sql &= " LEFT JOIN WV2 cv on cv.OCID=a.OCID" & vbCrLf
        sql &= " LEFT JOIN WT2 ct on ct.OCID=a.OCID" & vbCrLf
        sql &= " LEFT JOIN Class_Expect ce on ce.OCID=a.OCID" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料!"
        DataGridTable.Visible = False
        Save.Visible = False
        Print.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True
            Save.Visible = True
            Print.Visible = True

            'PageControler1.SqlDataCreate(sql, "RID,OCID,CyclType")
            PageControler1.PageDataTable = dt
            PageControler1.Sort = "RID,OCID,CyclType"
            PageControler1.ControlerLoad()

            Dim MyValue As String = ""
            MyValue = "b=156"
            MyValue &= "&Years=" & CStr(CInt(sm.UserInfo.Years) - 1911)
            MyValue &= "&OrgName=" & Convert.ToString(center.Text)
            MyValue &= "&ClassCName=" & Convert.ToString(OCID1.Text)
            MyValue &= "&Relship=" & Relship
            If sm.UserInfo.DistID <> "000" Then MyValue &= "&PlanID=" & sm.UserInfo.PlanID
            MyValue &= "&OCID=" & OCIDValue1.Value
            If flag_ROC Then
                MyValue &= "&STDate1=" & TIMS.Cdate18(STDate1.Text)  'edit，by:20181018
                MyValue &= "&STDate2=" & TIMS.Cdate18(STDate2.Text)  'edit，by:20181018
                MyValue &= "&FTDate1=" & TIMS.Cdate18(FTDate1.Text)  'edit，by:20181018
                MyValue &= "&FTDate2=" & TIMS.Cdate18(FTDate2.Text)  'edit，by:20181018
            Else
                MyValue &= "&STDate1=" & STDate1.Text  'edit，by:20181018
                MyValue &= "&STDate2=" & STDate2.Text  'edit，by:20181018
                MyValue &= "&FTDate1=" & FTDate1.Text  'edit，by:20181018
                MyValue &= "&FTDate2=" & FTDate2.Text  'edit，by:20181018
            End If
            Print.Attributes("onclick") = ReportQuery.ReportScript(Me, cst_printFN3, MyValue)
        End If

    End Sub

    Private Sub Query_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Query.Click
        Call sSearch1()
        Call GetSearchStr()
    End Sub

    Sub SaveData1(ByRef s_MSG As String)
        s_MSG = ""

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Integer = 0
        'i = 0'j = 0'k = 0
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vRIDValue As String = If(RIDValue.Value = "", sm.UserInfo.RID, RIDValue.Value)

        Dim Relship As String = TIMS.GET_RelshipforRID(vRIDValue, objconn)

        '查詢
        Dim sql As String = " SELECT * FROM CLASS_EXPECT WHERE OCID=@OCID AND SEQNO=@SEQNO "
        Dim sCmd As New SqlCommand(sql, objconn)

        '新增
        sql = " INSERT INTO CLASS_EXPECT (OCID,SEQNO,EXPECTDATE,VISITORNAME,MODIFYACCT,MODIFYDATE) " & vbCrLf
        sql += " VALUES (@OCID,@SEQNO,@EXPECTDATE,@VISITORNAME,@MODIFYACCT,@MODIFYDATE) "
        Dim iCmd As New SqlCommand(sql, objconn)

        '修改
        sql = " UPDATE CLASS_EXPECT "
        sql += " SET EXPECTDATE=@EXPECTDATE "
        sql += " ,VISITORNAME=@VISITORNAME "
        sql += " ,MODIFYACCT=@MODIFYACCT "
        sql += " ,MODIFYDATE=@MODIFYDATE "
        sql += " WHERE OCID=@OCID "
        sql += " AND SEQNO=@SEQNO "
        Dim uCmd As New SqlCommand(sql, objconn)

        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing

        '**by Milor 20080429--將原本的全部儲存，改為單一儲存，僅將預計抽訪人員與預計抽訪日期都不為空的資料進行儲存----start
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim labSEQNO As Label = eItem.FindControl("labSEQNO")
            Dim VisitorName As TextBox = eItem.FindControl("VisitorName")
            Dim ExpectDate As TextBox = eItem.FindControl("ExpectDate")
            Dim hid_OCID As HtmlInputHidden = eItem.FindControl("hid_OCID")
            Dim hid_ceSEQNO As HtmlInputHidden = eItem.FindControl("hid_ceSEQNO")

            If VisitorName.Text <> "" AndAlso ExpectDate.Text = "" Then
                j += 1
            ElseIf VisitorName.Text = "" AndAlso ExpectDate.Text <> "" Then
                k += 1
            ElseIf VisitorName.Text <> "" AndAlso ExpectDate.Text <> "" Then

                Dim sqlstr As String = ""
                '查詢資料
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = hid_OCID.Value
                    .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = hid_ceSEQNO.Value
                    dt = DbAccess.GetDataTable(sCmd.CommandText, objconn, sCmd.Parameters)
                End With

                If dt.Rows.Count = 0 Then
                    '新增
                    With iCmd
                        .Parameters.Clear()
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = hid_OCID.Value
                        .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = Get_iExpectSeqNoMax1(hid_ceSEQNO.Value)
                        .Parameters.Add("EXPECTDATE", SqlDbType.VarChar).Value = If(ExpectDate.Text = "", Convert.DBNull, If(flag_ROC, TIMS.Cdate18(ExpectDate.Text), ExpectDate.Text))  'edit，by:20181018
                        .Parameters.Add("VISITORNAME", SqlDbType.NVarChar).Value = IIf(VisitorName.Text = "", Convert.DBNull, VisitorName.Text)
                        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .Parameters.Add("MODIFYDATE", SqlDbType.DateTime).Value = Now
                        DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
                    End With
                    i += 1
                Else
                    With uCmd
                        .Parameters.Clear()
                        .Parameters.Add("EXPECTDATE", SqlDbType.VarChar).Value = If(ExpectDate.Text = "", Convert.DBNull, If(flag_ROC, TIMS.Cdate18(ExpectDate.Text), ExpectDate.Text))  'edit，by:20181018
                        .Parameters.Add("VISITORNAME", SqlDbType.NVarChar).Value = IIf(VisitorName.Text = "", Convert.DBNull, VisitorName.Text)
                        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .Parameters.Add("MODIFYDATE", SqlDbType.DateTime).Value = Now
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = hid_OCID.Value
                        .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = hid_ceSEQNO.Value
                        DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
                    End With
                    i += 1
                End If
            End If
        Next

        'Dim msg As String = ""
        If j > 0 Then s_MSG = "已填入預計抽訪人員時，預計抽訪日期不可空白。" & vbCrLf
        If k > 0 Then s_MSG += "已填入預計抽訪日期時，預計抽訪人員不可空白" & vbCrLf
        If i > 0 Then s_MSG += i & "筆資料儲存成功"
    End Sub

    Private Sub Save_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Save.Click
        Dim msg As String = ""
        Call SaveData1(msg)
        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        Query_Click(sender, e)
    End Sub

    Function Get_iExpectSeqNoMax1(ByVal OCID As String) As Integer
        Dim rst As Integer = 1
        Dim sql As String = "SELECT ISNULL(max(SeqNo),0)+1 MAXSEQNO FROM Class_Expect WHERE OCID =@OCID"
        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = CInt(OCID)
            rst = .ExecuteScalar
        End With
        Return rst
    End Function


    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        Save.Visible = False
        Print.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
        Save.Visible = False
        Print.Visible = False
    End Sub

    'Protected Sub Print_Click(sender As Object, e As EventArgs) Handles Print.Click
    'End Sub
End Class