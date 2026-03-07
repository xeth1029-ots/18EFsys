Public Class CR_01_002
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'SYS_GCODEREVIE/SYS_GRADEQUOTA
    'select top 10 * from SYS_GCODEREVIE
    'Const cst_SCORELEVEL_A As String = "A"
    'Const cst_SCORELEVEL_B As String = "B"
    'Const cst_SCORELEVEL_C As String = "C"
    'Const cst_SCORELEVEL_D As String = "D"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            CCreate1()
        End If
    End Sub

    Sub CCreate1()
        msg1.Text = ""
        tbDataGrid1.Visible = False

        'ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, TIMS.Get_DISTIDT2(objconn))
        'Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID)

        ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")
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
        If v_YEARS_SCH = "" OrElse v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return
        End If
        Dim v_YEARS_SCH_ROC As String = TIMS.GET_YEARS_ROC(v_YEARS_SCH)

        Dim parms As New Hashtable
        parms.Add("YEARS", v_YEARS_SCH)
        parms.Add("APPSTAGE", v_APPSTAGE_SCH)
        'If (v_ddlDISTID_SCH <> "") Then parms.Add("DISTID", v_ddlDISTID_SCH)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (SELECT GCODE,CNAME FROM ID_GOVCLASSCAST3 WHERE PARENTS IS NULL)" & vbCrLf
        'sql &= " SELECT  TOP 500" & vbCrLf sql &= " a.SGRID  /*PK*/" & vbCrLf
        sql &= " SELECT a.SGRID" & vbCrLf
        sql &= " ,c.GCODE" & vbCrLf
        sql &= " ,concat('(',c.GCODE,')',c.CNAME) GCODE_CNAME" & vbCrLf
        sql &= String.Format(" ,ISNULL(a.YEARS,'{0}') YEARS", v_YEARS_SCH) & vbCrLf
        'sql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sql &= String.Format(" ,ISNULL(dbo.FN_CYEAR2(a.YEARS),'{0}') YEARS_ROC", v_YEARS_SCH_ROC) & vbCrLf
        sql &= String.Format(" ,ISNULL(a.APPSTAGE,'{0}') APPSTAGE", v_APPSTAGE_SCH) & vbCrLf
        sql &= String.Format(" ,dbo.FN_GET_APPSTAGE(ISNULL(a.APPSTAGE,'{0}')) APPSTAGE_N", v_APPSTAGE_SCH) & vbCrLf
        sql &= " ,a.DISTID" & vbCrLf
        'sql &= " --,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM WC1 c" & vbCrLf
        sql &= " LEFT JOIN SYS_GCODEREVIE a on a.GCODE=c.GCODE" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.YEARS=@YEARS" & vbCrLf
        sql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        'If (v_ddlDISTID_SCH <> "") Then sql &= " AND a.DISTID=@DISTID" & vbCrLf

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
                Dim Hid_GCODE As HiddenField = e.Item.FindControl("Hid_GCODE")
                Dim Hid_DISTID As HiddenField = e.Item.FindControl("Hid_DISTID")
                'RadioButtonList ID="rbl_DISTNM
                Dim rbl_DISTNM As RadioButtonList = e.Item.FindControl("rbl_DISTNM")

                'Dim sCmdArg As String = ""
                'TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                'TIMS.SetMyValue(sCmdArg, "APPSTAGE", Convert.ToString(drv("APPSTAGE")))
                Hid_YEARS.Value = Convert.ToString(drv("YEARS"))
                Hid_APPSTAGE.Value = Convert.ToString(drv("APPSTAGE"))
                Hid_GCODE.Value = Convert.ToString(drv("GCODE"))
                Hid_DISTID.Value = Convert.ToString(drv("DISTID"))
                rbl_DISTNM = TIMS.Get_VDISTID2(rbl_DISTNM, objconn)
                Common.SetListItem(rbl_DISTNM, Hid_DISTID.Value)
        End Select
    End Sub

    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Call sSaveData1()
    End Sub

    ''' <summary>儲存</summary>
    Sub sSaveData1()
        Dim iRst As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim Hid_YEARS As HiddenField = eItem.FindControl("Hid_YEARS")
            Dim Hid_APPSTAGE As HiddenField = eItem.FindControl("Hid_APPSTAGE")
            Dim Hid_GCODE As HiddenField = eItem.FindControl("Hid_GCODE")
            Dim Hid_DISTID As HiddenField = eItem.FindControl("Hid_DISTID")
            'RadioButtonList ID="rbl_DISTNM
            Dim rbl_DISTNM As RadioButtonList = eItem.FindControl("rbl_DISTNM")
            Dim v_DISTID As String = TIMS.GetListValue(rbl_DISTNM)

            '立即將控制權轉移到迴圈的下一個反復專案。
            Dim flag_Continue As Boolean = False
            If v_DISTID = "" Then flag_Continue = True '都沒有值，不儲存
            '由於有值，且沒有異動，所以不儲存
            If Hid_DISTID.Value <> "" AndAlso Hid_DISTID.Value = v_DISTID Then flag_Continue = True
            '立即將控制權轉移到迴圈的下一個反復專案。
            If flag_Continue Then Continue For

            iRst += UPDATE1_ROW1(eItem)
        Next
        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
    End Sub

    ''' <summary>'更新／新增</summary>
    ''' <param name="eItem"></param>
    ''' <returns></returns>
    Function UPDATE1_ROW1(ByRef eItem As DataGridItem) As Integer
        Dim iRst As Integer = 0
        Dim s_sql As String
        s_sql = "" & vbCrLf
        s_sql &= " SELECT SGRID FROM SYS_GCODEREVIE" & vbCrLf
        s_sql &= " WHERE YEARS=@YEARS AND APPSTAGE=@APPSTAGE AND GCODE=@GCODE" & vbCrLf

        Dim i_sql As String = ""
        i_sql = "" & vbCrLf
        i_sql &= " INSERT INTO SYS_GCODEREVIE(SGRID ,YEARS ,APPSTAGE ,GCODE ,DISTID ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@SGRID ,@YEARS ,@APPSTAGE ,@GCODE ,@DISTID ,@MODIFYACCT ,GETDATE())" & vbCrLf

        Dim u_sql As String = ""
        u_sql = ""
        u_sql &= " UPDATE SYS_GCODEREVIE" & vbCrLf
        u_sql &= " SET DISTID=@DISTID" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE 1=1" & vbCrLf
        u_sql &= " AND SGRID=@SGRID" & vbCrLf
        u_sql &= " AND GCODE=@GCODE" & vbCrLf

        Dim Hid_YEARS As HiddenField = eItem.FindControl("Hid_YEARS")
        Dim Hid_APPSTAGE As HiddenField = eItem.FindControl("Hid_APPSTAGE")
        Dim Hid_GCODE As HiddenField = eItem.FindControl("Hid_GCODE")
        Dim Hid_DISTID As HiddenField = eItem.FindControl("Hid_DISTID")
        'RadioButtonList ID="rbl_DISTNM
        Dim rbl_DISTNM As RadioButtonList = eItem.FindControl("rbl_DISTNM")
        Dim v_DISTID As String = TIMS.GetListValue(rbl_DISTNM)
        '有值有異動才更新
        Dim flag_modify1 As Boolean = If(v_DISTID <> "" AndAlso Hid_DISTID.Value <> v_DISTID, True, False)
        If Not flag_modify1 Then Return iRst

        Dim parms As New Hashtable
        parms.Add("YEARS", Hid_YEARS.Value)
        parms.Add("APPSTAGE", Hid_APPSTAGE.Value)
        parms.Add("GCODE", Hid_GCODE.Value)
        Dim dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, parms)

        If dt.Rows.Count = 0 Then
            Dim iSGRID As Integer = DbAccess.GetNewId(objconn, "SYS_GCODEREVIE_SGRID_SEQ,SYS_GCODEREVIE,SGRID")
            Dim i_parms As New Hashtable
            i_parms.Add("SGRID", iSGRID)
            i_parms.Add("YEARS", Hid_YEARS.Value)
            i_parms.Add("APPSTAGE", Hid_APPSTAGE.Value)
            i_parms.Add("GCODE", Hid_GCODE.Value)
            i_parms.Add("DISTID", v_DISTID)
            i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            iRst += DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
        Else
            Dim dr1 As DataRow = dt.Rows(0)
            Dim iSGRID As Integer = Val(dr1("SGRID"))
            Dim u_parms As New Hashtable
            u_parms.Add("DISTID", v_DISTID)
            u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            u_parms.Add("SGRID", iSGRID)
            u_parms.Add("GCODE", Hid_GCODE.Value)
            iRst += DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        End If
        Return iRst
    End Function

End Class