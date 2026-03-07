Public Class CR_01_006
    Inherits AuthBasePage

    'OJT-23020102：產投 - 新增【年度-主責分署設定】功能
    '年度-主責分署設定/SYS_LIABILITY
    'SELECT TOP 10 * FROM SYS_LIABILITY
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        msg1.Text = ""
        tbDataGrid1.Visible = False
        PanelEdit1.Visible = False
        'tbPanelEdit1.Visible = False

        '【年度】：由112年開始~當年度+1   (因半年前會開始審隔年的班級)
        Dim iSYears As Integer = 2023
        Dim iEYears As Integer = Now.Year + 1

        ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH, iSYears, iEYears, False)
        Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)
        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        ddlYEARS = TIMS.GetSyear(ddlYEARS, iSYears, iEYears, True)
        ddlAPPSTAGE = TIMS.Get_APPSTAGE2(ddlAPPSTAGE)
        rblDISTMAIN = TIMS.Get_VDISTID2(rblDISTMAIN, objconn)
    End Sub

    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Call sSearch1()
    End Sub

    Sub sSearch1()
        msg1.Text = TIMS.cst_NODATAMsg1
        tbDataGrid1.Visible = False

        'Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH)
        Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH)
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        If v_YEARS_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return
        End If
        Dim v_YEARS_SCH_ROC As String = TIMS.GET_YEARS_ROC(v_YEARS_SCH)

        Dim parms As New Hashtable
        parms.Add("YEARS", v_YEARS_SCH)
        If v_APPSTAGE_SCH <> "" Then parms.Add("APPSTAGE", v_APPSTAGE_SCH)

        Dim sql As String = ""
        sql &= " SELECT a.SLBID,a.YEARS,a.APPSTAGE,a.DISTID" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,dbo.FN_GET_APPSTAGE(a.APPSTAGE) APPSTAGE_N" & vbCrLf
        sql &= " ,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sql &= " FROM SYS_LIABILITY a" & vbCrLf
        sql &= " WHERE a.YEARS=@YEARS" & vbCrLf
        If v_APPSTAGE_SCH <> "" Then sql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            msg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        msg1.Text = ""
        tbDataGrid1.Visible = True
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim Hid_YEARS As HiddenField = e.Item.FindControl("Hid_YEARS")
                Dim Hid_APPSTAGE As HiddenField = e.Item.FindControl("Hid_APPSTAGE")
                Dim Hid_DISTID As HiddenField = e.Item.FindControl("Hid_DISTID")
                'RadioButtonList ID="rbl_DISTNM
                Dim rbl_DISTNM As RadioButtonList = e.Item.FindControl("rbl_DISTNM")

                Hid_YEARS.Value = Convert.ToString(drv("YEARS"))
                Hid_APPSTAGE.Value = Convert.ToString(drv("APPSTAGE"))
                Hid_DISTID.Value = Convert.ToString(drv("DISTID"))
                rbl_DISTNM = TIMS.Get_VDISTID2(rbl_DISTNM, objconn)
                Common.SetListItem(rbl_DISTNM, Hid_DISTID.Value)
                rbl_DISTNM.Enabled = False
                TIMS.Tooltip(rbl_DISTNM, "暫不提供修改選項")
        End Select
    End Sub

    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Call sSaveData1()
    End Sub

    ''' <summary> 檢核有誤／有資料：false  </summary>
    ''' <param name="v_YEARS"></param>
    ''' <param name="v_APPSTAGE"></param>
    ''' <returns></returns>
    Function CHK_LIABILITY(ByRef v_YEARS As String, ByRef v_APPSTAGE As String) As Boolean
        If v_YEARS = "" Then Return False
        If v_APPSTAGE = "" Then Return False

        Dim parms As New Hashtable
        parms.Add("YEARS", v_YEARS)
        parms.Add("APPSTAGE", v_APPSTAGE)
        Dim sql As String = ""
        sql &= " SELECT a.SLBID,a.YEARS,a.APPSTAGE,a.DISTID" & vbCrLf
        sql &= " ,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sql &= " FROM SYS_LIABILITY a" & vbCrLf
        sql &= " WHERE a.YEARS=@YEARS" & vbCrLf
        sql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then Return False
        Return True
    End Function

    ''' <summary>儲存</summary>
    Sub sSaveData1()
        Dim iRst As Integer = 0

        Dim eErrMsg1 As String = ""
        'tbPanelEdit1
        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_ddlAPPSTAGE As String = TIMS.GetListValue(ddlAPPSTAGE)
        Dim v_rblDISTMAIN As String = TIMS.GetListValue(rblDISTMAIN)

        If v_ddlYEARS = "" Then eErrMsg1 &= "年度資料不可為空" & vbCrLf
        If v_ddlAPPSTAGE = "" Then eErrMsg1 &= "申請階段資料不可為空" & vbCrLf
        If v_rblDISTMAIN = "" Then eErrMsg1 &= "主責分署資料不可為空" & vbCrLf
        If eErrMsg1 = "" AndAlso Not CHK_LIABILITY(v_ddlYEARS, v_ddlAPPSTAGE) Then eErrMsg1 &= "該年度／申請階段已有資料不可再次新增" & vbCrLf

        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Return
        End If

        Dim i_sql As String = ""
        i_sql &= " INSERT INTO SYS_LIABILITY(SLBID ,YEARS,APPSTAGE ,DISTID ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@SLBID ,@YEARS ,@APPSTAGE ,@DISTID ,@MODIFYACCT ,GETDATE())" & vbCrLf
        Dim iSLBID As Integer = DbAccess.GetNewId(objconn, "SYS_LIABILITY_SLBID_SEQ,SYS_LIABILITY,SLBID")
        Dim i_parms As New Hashtable
        i_parms.Add("SLBID", iSLBID)
        i_parms.Add("YEARS", v_ddlYEARS)
        i_parms.Add("APPSTAGE", v_ddlAPPSTAGE)
        i_parms.Add("DISTID", v_rblDISTMAIN)
        i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        iRst += DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)

        Common.SetListItem(ddlYEARS, "")
        Common.SetListItem(ddlAPPSTAGE, "")
        Common.SetListItem(rblDISTMAIN, "")
        PanelEdit1.Visible = False
        panelSch.Visible = True
    End Sub

    Protected Sub BtnAddNew1_Click(sender As Object, e As EventArgs) Handles BtnAddNew1.Click
        Common.SetListItem(ddlYEARS, "")
        Common.SetListItem(ddlAPPSTAGE, "")
        Common.SetListItem(rblDISTMAIN, "")

        PanelEdit1.Visible = True
        panelSch.Visible = False
        'tbPanelEdit1.Visible = True
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        Common.SetListItem(ddlYEARS, "")
        Common.SetListItem(ddlAPPSTAGE, "")
        Common.SetListItem(rblDISTMAIN, "")

        PanelEdit1.Visible = False
        panelSch.Visible = True
    End Sub
End Class