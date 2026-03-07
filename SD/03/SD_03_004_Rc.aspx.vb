Partial Class SD_03_004_Rc
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在---------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在---------------------------End

        If Not IsPostBack Then
            'sql = ""
            'sql += " SELECT NORID,OTHERREASON,ISBUSINESS,APPLIEDRESULTR,APPLIEDRESULTM,PNUM,   " & vbCrLf
            'sql += " ISCONT,QAYSDATE,QAYFDATE,LASTSTATE,TADDRESSZIP2W,EVTA_NOSHOW,   " & vbCrLf
            'sql += " ETRAIN_SHOW,ECOMMENT,COMPANYNAME,NOTICE,CJOB_UNKEY,EXAMPERIOD,   " & vbCrLf
            'sql += " OCID,CLSID,PLANID,YEARS,CYCLTYPE,LEVELTYPE,RID,CLASSCNAME,CLASSENGNAME,   " & vbCrLf
            'sql += " dbo.SUBSTR(CONTENT, 1, 4000) CONTENT,dbo.SUBSTR(PURPOSE, 1, 4000) PURPOSE,   " & vbCrLf
            'sql += " TPROPERTYID,TMID,CLID,SENTERDATE,FENTERDATE,CHECKINDATE,EXAMDATE,STDATE,   " & vbCrLf
            'sql += " FTDATE,TADDRESSZIP,TADDRESS,THOURS,TNUM,TDEADLINE,TPERIOD,NOTOPEN,   " & vbCrLf
            'sql += " ISAPPLIC,RELSHIP,COMIDNO,SEQNO,ISCALCULATE,ISSUCCESS,CTNAME,MODIFYACCT,   " & vbCrLf
            'sql += " MODIFYDATE,CLASSNUM,LEVELCOUNT,ISFULLDATE,CLASS_UNIT,ISCLOSED,BGTIME   " & vbCrLf
            Dim ReqTMID As String = Request("TMID")
            Dim Reqyears As String = Right(Request("years"), 2)
            ReqTMID = TIMS.ClearSQM(ReqTMID)
            Reqyears = TIMS.ClearSQM(Reqyears)
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql += " SELECT OCID" & vbCrLf
            sql += " ,ClassCName" & vbCrLf
            sql += " ,CyclType" & vbCrLf
            sql += " ,LevelType" & vbCrLf
            sql += " FROM Class_ClassInfo" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            sql += " AND TMID='" & ReqTMID & "'" & vbCrLf
            sql += " and RID='" & sm.UserInfo.RID & "'" & vbCrLf
            sql += " and PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql += " and Years='" & Reqyears & "'"
            Dim dt As DataTable = DbAccess.GetDataTable(sql)

            If dt.Rows.Count = 0 Then
                DataGrid1.Visible = False
                Common.RespWrite(Me, "<script>")
                Common.RespWrite(Me, "alert('查不到此職類的班級');")
                Common.RespWrite(Me, "window.close();")
                Common.RespWrite(Me, "</script>")
            Else
                DataGrid1.Visible = True
                DataGrid1.DataSource = dt
                DataGrid1.DataKeyField = "OCID"
                DataGrid1.DataBind()
            End If
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            'Dim myradio As HtmlInputRadioButton
            'myradio = e.Item.Cells(0).FindControl("Radio1")
            Dim myradio As HtmlInputRadioButton = e.Item.FindControl("Radio1")
            myradio.Value = DataGrid1.DataKeys(e.Item.ItemIndex)
            e.Item.Cells(1).Text += "第" & TIMS.GetChtNum(CInt(e.Item.Cells(2).Text)) & "期"
            If CInt(e.Item.Cells(3).Text) <> 0 Then
                e.Item.Cells(1).Text += "第" & TIMS.GetChtNum(CInt(e.Item.Cells(3).Text)) & "階段"
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dr As DataGridItem
        For Each dr In DataGrid1.Items
            Dim myradio As HtmlInputRadioButton
            myradio = dr.Cells(0).FindControl("Radio1")
            If myradio.Checked = True Then
                Common.RespWrite(Me, "<script language='javascript'>")
                Common.RespWrite(Me, "window.opener.document.form1.OCIDName.value='" & dr.Cells(1).Text & "';")
                Common.RespWrite(Me, "window.opener.document.form1.OCIDValue.value='" & myradio.Value & "';")
                Common.RespWrite(Me, "window.close();")
                Common.RespWrite(Me, "</script>")
            End If
        Next
    End Sub
End Class
