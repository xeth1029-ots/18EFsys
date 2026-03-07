Public Class TC_01_022
    Inherits AuthBasePage 'System.Web.UI.Page

    'Const cst_mmo1 As String = "※確認送出後即鎖定不可再修改!"
    'ORG_TTQSLOCK / ORG_TTQS2

    Dim objconn As SqlConnection = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(MRqID, TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call cCreate1()
        End If
    End Sub

    Sub cCreate1()
        lab_msg1.Text = ""
        If Not CHECK_TTQSQUERY() Then
            Dim v_msg_1 As String = "目前無開放單位 最近一次審查計分等級查詢!"
            lab_msg1.Text = v_msg_1
            Common.MessageBox(Me, v_msg_1)
            Return
        End If

        Call SHOW_ORGSCORING2()
    End Sub

    Private Sub SHOW_ORGSCORING2()
        HID_SCORINGID.Value = TIMS.ClearSQM(HID_SCORINGID.Value)
        If HID_SCORINGID.Value = "" Then Return

        Dim sPMS As New Hashtable
        sPMS.Add("TPLANID", sm.UserInfo.TPlanID)
        sPMS.Add("ORGID", sm.UserInfo.OrgID)
        sPMS.Add("SCORINGID", HID_SCORINGID.Value)

        Dim sSql As String = ""
        sSql &= " SELECT a.TPLANID,oo.ORGID,oo.COMIDNO" & vbCrLf
        sSql &= " ,a.YEARS ,a.YEARS1" & vbCrLf
        sSql &= " ,a.FIRSTCHK ,a.SECONDCHK,a.RLEVEL_2" & vbCrLf
        sSql &= " ,oo.ORGNAME ,oo.COMIDNO" & vbCrLf
        sSql &= " ,a.DISTID,kd.NAME DISTNAME" & vbCrLf
        sSql &= " ,oo.ORGKIND,k1.NAME ORGKIND_N" & vbCrLf
        sSql &= " ,v1.VNAME SENDVER_N" & vbCrLf
        sSql &= " ,v2.VNAME RESULT_N" & vbCrLf
        sSql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
        sSql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID = a.OrgID" & vbCrLf
        sSql &= " JOIN dbo.ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID = a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sSql &= " LEFT JOIN KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID = oo.ORGKIND" & vbCrLf
        sSql &= " LEFT JOIN dbo.ORG_TTQS2 b On concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
        sSql &= " LEFT JOIN dbo.V_SENDVER v1 On v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sSql &= " LEFT JOIN dbo.V_RESULT v2 On v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
        sSql &= " WHERE a.FIRSTCHK='Y' AND a.SECONDCHK='Y'" & vbCrLf
        sSql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        sSql &= " AND a.ORGID=@ORGID" & vbCrLf
        sSql &= " AND CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2)=@SCORINGID " & vbCrLf
        sSql &= " ORDER BY kd.DISTID,a.OSID2" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
        'PageControler1.Visible = False 'DataGridTable.Visible = False
        DataGrid1.Visible = False
        lab_msg1.Text = "查無資料"
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        DataGrid1.Visible = True
        lab_msg1.Text = ""
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    ''' <summary>檢查是否有資料，有資料為TRUE, 沒資料為FALSE</summary>
    ''' <returns></returns>
    Private Function CHECK_TTQSQUERY() As Boolean
        Dim sSql As String = ""
        sSql &= " WITH WQ1 AS (" & " SELECT CONCAT(dbo.FN_CYEAR2(YEARS),'年',MONTHS,'月'" & vbCrLf
        sSql &= " ,'(',dbo.FN_CYEAR2(YEARS1) ,'年',case when HALFYEAR1=1 then '上半年' else '下半年' end ,'~'" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(YEARS2) ,'年',case when HALFYEAR2=1 then '上半年' else '下半年' end ,')') TEXTFD" & vbCrLf
        sSql &= " ,CONCAT(YEARS ,'-',MONTHS,'-',YEARS1 ,'-',HALFYEAR1,'-',YEARS2 ,'-',HALFYEAR2) VALUEFD" & vbCrLf
        sSql &= " ,OTLID FROM ORG_TTQSLOCK )" & vbCrLf

        sSql &= " SELECT a.OTQID" & vbCrLf
        sSql &= " ,format(a.QCDATE,'yyyy/MM/dd') QCDATE" & vbCrLf
        sSql &= " ,format(a.QSDATE,'yyyy/MM/dd HH:mm') QSDATE" & vbCrLf
        sSql &= " ,format(a.QFDATE,'yyyy/MM/dd HH:mm') QFDATE" & vbCrLf
        sSql &= " ,a.QEXPLAIN" & vbCrLf
        sSql &= " ,a.YEARS ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,a.APPSTAGE ,dbo.FN_GET_APPSTAGE(a.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,a.OTLID,dbo.FN_GET_TTQSLOCK_N(a.OTLID) TTQSLOCK_N" & vbCrLf
        sSql &= " ,a.ISDELETE" & vbCrLf
        sSql &= " ,q.TEXTFD ,q.VALUEFD" & vbCrLf
        sSql &= " FROM ORG_TTQSQUERY a" & vbCrLf
        sSql &= " JOIN WQ1 q ON q.OTLID=a.OTLID" & vbCrLf
        sSql &= " WHERE a.ISDELETE IS NULL" & vbCrLf
        sSql &= " AND (a.QSDATE <= GETDATE() AND a.QFDATE >= GETDATE())" & vbCrLf
        sSql &= " ORDER BY a.QCDATE DESC,a.MODIFYDATE DESC" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql, objconn)
        If dr1 IsNot Nothing Then
            YEARS_ROC.Text = Convert.ToString(dr1("YEARS_ROC")) '年度N
            APPSTAGE_N.Text = Convert.ToString(dr1("APPSTAGE_N")) '申請階段

            Hid_OTQID.Value = Convert.ToString(dr1("OTQID"))
            HID_SCORINGID.Value = Convert.ToString(dr1("VALUEFD"))
            Hid_YEARS.Value = Convert.ToString(dr1("YEARS"))
            Hid_APPSTAGE.Value = Convert.ToString(dr1("APPSTAGE"))
        End If
        Return If(dr1 IsNot Nothing, True, False)
    End Function
End Class