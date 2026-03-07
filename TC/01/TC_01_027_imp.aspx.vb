Partial Class TC_01_027_imp
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            btnImpYear.Attributes("onclick") = "return chkdata();"
            Fromyear = TIMS.GetSyear(Fromyear, 0, sm.UserInfo.Years, True)
            'Toyear = TIMS.GetSyear(Toyear)
            hidToyear.Value = sm.UserInfo.Years
            DistID = TIMS.Get_DistID(DistID, Nothing, objconn)
#Region "(No Use)"

            'Dim sql As String
            'Dim dt As DataTable
            'sql = "SELECT * FROM ID_District"
            'dt = DbAccess.GetDataTable(sql, objconn)
            'Me.ViewState("DistID") = dt

            ''要是署(局)的身分，要產生所有的轄區代碼
            'If sm.UserInfo.LID = 0 Then
            '    With DistID
            '        .DataSource = dt
            '        .DataTextField = "Name"
            '        .DataValueField = "DistID"
            '        .DataBind()
            '        .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            '    End With
            'End If

            'Tplan = TIMS.Get_TPlan(Tplan, , 1)
            'sql = "SELECT * FROM Key_Plan"
            'dt = DbAccess.GetDataTable(sql, objconn)
            'Me.ViewState("TPlan") = dt

#End Region
        End If
#Region "(No Use)"

        'If sm.UserInfo.LID = 0 Then
        '    Table3.Style.Item("display") = "inline"
        'Else
        '    Table3.Style.Item("display") = "none"
        'End If
        'If Not Me.ViewState("dt") Is Nothing Then PageControler1.PageDataTable = Me.ViewState("dt")
        'Button1.Attributes("onclick") = "return chkdata();"
        'Button2.Attributes("onclick") = "window.close();"

#End Region
    End Sub

    ''' <summary>
    ''' true:檢核後可新增 false:異常-不可新增
    ''' </summary>
    ''' <param name="RID"></param>
    ''' <param name="TeacherID"></param>
    ''' <returns></returns>
    Function chk1(ByVal RID As String, ByVal TeacherID As String) As Boolean
        Dim rst As Boolean = True 'true:檢核後可新增 false:異常-不可新增
        Dim sql As String = ""
        sql = " SELECT 'x' FROM TEACH_TEACHERINFO WHERE RID = @RID AND TeacherID = @TeacherID "
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("RID", RID)
        parms.Add("TeacherID", TeacherID)
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then rst = False '已有資料
        Return rst
    End Function

    Sub sSavedata1(ByRef dr1 As DataRow)
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim iTECHID As Integer = 0
        iTECHID = DbAccess.GetNewId(objconn, "TEACH_TEACHERINFO_TECHID_SEQ,TEACH_TEACHERINFO,TECHID")
        sql = " SELECT * FROM TEACH_TEACHERINFO WHERE 1<>1 "
        dt = DbAccess.GetDataTable(sql, da, objconn)
        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("TECHID") = iTECHID 'TEACH_TEACHERINFO_TECHID_SEQ
        dr("RID") = sm.UserInfo.RID
        dr("TeacherID") = dr1("TeacherID")
        dr("TeachCName") = dr1("TeachCName")
        dr("TeachEName") = dr1("TeachEName")
        dr("Birthday") = dr1("Birthday")
        dr("IDNO") = TIMS.ChangeIDNO(dr1("IDNO"))
        dr("TMID") = dr1("TMID")
        dr("DegreeID") = dr1("DegreeID")
        dr("SchoolName") = dr1("SchoolName")
        dr("Department") = dr1("Department")
        dr("GraduateStatus") = dr1("GraduateStatus")
        dr("Phone") = dr1("Phone")
        dr("AddressZip") = dr1("AddressZip")
        dr("AddressZIP6W") = dr1("AddressZIP6W")
        dr("Address") = dr1("Address")

        dr("WorkOrg") = dr1("WorkOrg")
        dr("IVID") = dr1("IVID")
        dr("Invest") = dr1("Invest")
        dr("TechType1") = dr1("TechType1")
        dr("TechType2") = dr1("TechType2")
        dr("ExpYears") = dr1("ExpYears")

        dr("WorkZip") = dr1("WorkZip")
        dr("WorkZIP6W") = dr1("WorkZIP6W")
        dr("WorkAddr") = dr1("WorkAddr")

        dr("WorkPhone") = dr1("WorkPhone") ' IIf(WorkPhone.Text = "", Convert.DBNull, WorkPhone.Text)
        dr("ExpUnit1") = dr1("ExpUnit1") ' IIf(ExpUnit1.Text = "", Convert.DBNull, ExpUnit1.Text)
        dr("ExpSDate1") = dr1("ExpSDate1") 'IIf(ExpSDate1.Text = "", Convert.DBNull, ExpSDate1.Text)
        dr("ExpEDate1") = dr1("ExpEDate1") 'IIf(ExpEDate1.Text = "", Convert.DBNull, ExpEDate1.Text)
        dr("ExpYears1") = dr1("ExpYears1") 'IIf(ExpYears1.Text = "", Convert.DBNull, ExpYears1.Text)
        dr("ExpUnit2") = dr1("ExpUnit2") 'IIf(ExpUnit2.Text = "", Convert.DBNull, ExpUnit2.Text)
        dr("ExpSDate2") = dr1("ExpSDate2") 'IIf(ExpSDate2.Text = "", Convert.DBNull, ExpSDate2.Text)
        dr("ExpEDate2") = dr1("ExpEDate2") 'IIf(ExpEDate2.Text = "", Convert.DBNull, ExpEDate2.Text)
        dr("ExpYears2") = dr1("ExpYears2") 'IIf(ExpYears2.Text = "", Convert.DBNull, ExpYears2.Text)
        dr("ExpUnit3") = dr1("ExpUnit3") 'IIf(ExpUnit3.Text = "", Convert.DBNull, ExpUnit3.Text)
        dr("ExpSDate3") = dr1("ExpSDate3") 'IIf(ExpSDate3.Text = "", Convert.DBNull, ExpSDate3.Text)
        dr("ExpEDate3") = dr1("ExpEDate3") 'IIf(ExpEDate3.Text = "", Convert.DBNull, ExpEDate3.Text)
        dr("ExpYears3") = dr1("ExpYears3") 'IIf(ExpYears3.Text = "", Convert.DBNull, ExpYears3.Text)
        dr("INV1") = dr1("INV1") 'IIf(tINV1.Text = "", Convert.DBNull, tINV1.Text)
        dr("INV2") = dr1("INV2") 'IIf(tINV2.Text = "", Convert.DBNull, tINV2.Text)
        dr("INV3") = dr1("INV3") 'IIf(tINV3.Text = "", Convert.DBNull, tINV3.Text)
        dr("ExpMonths") = dr1("ExpMonths") 'IIf(ExpMonths.SelectedIndex = 0, Convert.DBNull, ExpMonths.SelectedValue)
        dr("ExpMonths1") = dr1("ExpMonths1") 'IIf(ExpMonths1.SelectedIndex = 0, Convert.DBNull, ExpMonths1.SelectedValue)
        dr("ExpMonths2") = dr1("ExpMonths2") 'IIf(ExpMonths2.SelectedIndex = 0, Convert.DBNull, ExpMonths2.SelectedValue)
        dr("ExpMonths3") = dr1("ExpMonths3") 'IIf(ExpMonths3.SelectedIndex = 0, Convert.DBNull, ExpMonths3.SelectedValue)
        dr("Specialty1") = dr1("Specialty1") 'IIf(Specialty1.Text = "", Convert.DBNull, Specialty1.Text)
        dr("Specialty2") = dr1("Specialty2") ' IIf(Specialty2.Text = "", Convert.DBNull, Specialty2.Text)
        dr("Specialty3") = dr1("Specialty3") 'IIf(Specialty3.Text = "", Convert.DBNull, Specialty3.Text)
        dr("Specialty4") = dr1("Specialty4") 'IIf(Specialty4.Text = "", Convert.DBNull, Specialty4.Text)
        dr("Specialty5") = dr1("Specialty5") 'IIf(Specialty5.Text = "", Convert.DBNull, Specialty5.Text)
        dr("KindID") = dr1("KindID") ' 130
        dr("KindEngage") = dr1("KindEngage") 'KindEngage.SelectedValue
        dr("WorkStatus") = dr1("WorkStatus") 'WorkStatus.SelectedValue '排課使用
        dr("Sex") = dr1("Sex") 'Sex.SelectedValue
        dr("Mobile") = dr1("Mobile") 'IIf(Mobile.Text = "", Convert.DBNull, Mobile.Text)
        dr("Email") = dr1("Email") 'IIf(Email.Text = "", Convert.DBNull, Email.Text)
        dr("ServDept") = dr1("ServDept") 'IIf(ServDept.Text = "", Convert.DBNull, ServDept.Text)
        dr("Fax") = dr1("Fax") 'IIf(Fax.Text = "", Convert.DBNull, Fax.Text)
        dr("TransBook") = dr1("TransBook") 'IIf(TransBook.Text = "", Convert.DBNull, TransBook.Text)
        '未輸入有效資訊!!
        dr("ProLicense1") = dr1("ProLicense1")
        dr("ProLicense2") = dr1("ProLicense2")
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        dr("PassPortNO") = dr1("PassPortNO")
        DbAccess.UpdateDataTable(dt, da)
    End Sub

    Sub xYearImport1()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT tt.* " & vbCrLf
        sql &= " FROM TEACH_TEACHERINFO tt " & vbCrLf
        sql &= " JOIN view_ridname rr ON rr.rid= tt.rid " & vbCrLf
        sql &= " JOIN view_plan ip ON ip.planid = rr.planid " & vbCrLf
        sql &= " JOIN org_orginfo oo ON oo.orgid = rr.orgid " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND ip.years = @Years " & vbCrLf
        sql &= " AND ip.distid = @DistID " & vbCrLf
        sql &= " AND ip.tplanid = @TPlanID " & vbCrLf
        sql &= " AND oo.orgid = @OrgID " & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("Years", Fromyear.SelectedValue)
        parms.Add("DistID", DistID.SelectedValue)
        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        parms.Add("OrgID", sm.UserInfo.OrgID)
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt1.Rows.Count = 0 Then
            sm.LastResultMessage = "該年度無師資資料!"
            Exit Sub
        End If
        For Each dr1 As DataRow In dt1.Rows
            Dim flag_save As Boolean = chk1(sm.UserInfo.RID, Convert.ToString(dr1("TeacherID")))
            If flag_save Then Call sSavedata1(dr1)
        Next
        'Dim rqMID As String = TIMS.Get_MRqID(Me)
        'Dim url1 As String = "./TC_01_027.aspx?ID=" & rqMID
        'sm.LastResultMessage = "年度複製成功!"  'edit，by:20181024
        'sm.RedirectUrlAfterBlock = ResolveUrl(url1)

        Dim rqMID As String = TIMS.Get_MRqID(Me)
        sm.LastResultMessage = "年度複製成功!"  'edit，by:20181024
        Dim url1 As String = "TC_01_027.aspx?ID=" & rqMID
        TIMS.Utl_Redirect(Me, objconn, url1)
#Region "(No Use)"

        'Common.MessageBox(Me, "年度複製成功!")
        'Page.RegisterStartupScript("", "<script>window.close();</script>")

#End Region
    End Sub

    Protected Sub btnImpYear_Click(sender As Object, e As EventArgs) Handles btnImpYear.Click
        xYearImport1()
    End Sub

    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Dim url1 As String = "TC_01_027.aspx?ID=" & rqMID
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class