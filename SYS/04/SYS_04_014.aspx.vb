Public Class SYS_04_014
    'Inherits System.Web.UI.Page
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            TIMS.Display_None(txt_localareanet)
            Call search1()
        End If

    End Sub

    Sub search1()
        Dim s_ROWNUM As String = If(TIMS.IsNumeric1(txtROWNUM.Text), CInt(txtROWNUM.Text), 6)
        txtROWNUM.Text = s_ROWNUM

        Dim parms As New Hashtable
        'parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("Years", sm.UserInfo.Years)

        '查詢沒有座標資訊的課程
        Dim sql As String = ""
        sql &= String.Format(" select top {0} cc.OCID,cc.CLASSCNAME", s_ROWNUM) & vbCrLf
        sql &= " ,cc.CLASSCNAME" & vbCrLf
        sql &= " ,sj.CJOBNAME2 CJOBNAME" & vbCrLf
        sql &= " ,oo.ORGNAME" & vbCrLf
        sql &= " ,cc.TADDRESSZIP" & vbCrLf
        sql &= " ,iz.CTNAME,iz.ZNAME,iz.ZIPNAME,cc.TADDRESS" & vbCrLf
        sql &= " ,cc.TWD97_X,cc.TWD97_Y" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=cc.PLANID and pp.COMIDNO=cc.COMIDNO and pp.SEQNO=cc.SEQNO" & vbCrLf
        sql &= " JOIN V_SHARECJOB sj on sj.CJOB_UNKEY=cc.CJOB_UNKEY" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno=cc.comidno" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid=cc.planid" & vbCrLf
        sql &= " JOIN VIEW_ZIPNAME iz on iz.zipcode=cc.taddresszip" & vbCrLf
        sql &= " WHERE CC.ISSUCCESS='Y'" & vbCrLf
        sql &= " AND CC.NOTOPEN='N'" & vbCrLf
        sql &= " AND CC.EVTA_NOSHOW IS NULL" & vbCrLf
        'sql &= " AND CC.TNUM>0" & vbCrLf
        'sql &= " AND CC.ISBUSINESS='N'" & vbCrLf
        'sql &= " AND ip.TPLANID='28'" & vbCrLf
        'sql &= " AND CC.ONSHELLDATE <= GETDATE()" & vbCrLf
        'sql &= " AND CC.FTDATE > GETDATE()" & vbCrLf
        'sql &= " AND CC.FENTERDATE >= GETDATE()" & vbCrLf
        'sql &= " AND CC.ISSUCCESS='Y'" & vbCrLf
        'sql &= " AND CC.NOTOPEN='N'" & vbCrLf
        'sql &= " and ip.tplanid ='28'" & vbCrLf
        'sql &= " and ip.years >='2021'" & vbCrLf
        '地址座標為空
        sql &= " and (cc.TWD97_X is null or cc.TWD97_Y is null)" & vbCrLf
        'sql &= " and ip.TPLANID =@TPLANID" & vbCrLf
        sql &= " and ip.years =@Years" & vbCrLf
        sql &= " ORDER BY NEWID()" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        litMsg.Text = String.Format("共 {0} 筆待地址定位", dt.Rows.Count)

        lv_classQueryResult.DataSource = dt ';//結果資料集
        lv_classQueryResult.DataBind()

    End Sub

    ''' <summary>
    ''' 組合機構地址 但地址可能有重複字，則排除
    ''' </summary>
    ''' <param name="drv"></param>
    ''' <returns></returns>
    Function get_TrainAddress(ByRef drv As DataRowView) As String
        Dim rst As String = ""
        Dim ctname As String = drv("CTNAME").ToString()
        Dim zname As String = drv("ZNAME").ToString()
        Dim zipname As String = drv("ZIPNAME").ToString()
        Dim TADDRESS As String = drv("TADDRESS").ToString()
        If TADDRESS.StartsWith(ctname) OrElse TADDRESS.Contains(zname) OrElse TADDRESS.StartsWith(zipname) Then
            rst = TADDRESS
            Return rst
        End If
        rst = String.Format("{0}{1}", drv("ZIPNAME").ToString(), drv("TADDRESS").ToString())
        Return rst
    End Function

    Private Sub lv_classQueryResult_ItemDataBound(sender As Object, e As ListViewItemEventArgs) Handles lv_classQueryResult.ItemDataBound

        Select Case e.Item.ItemType
            Case ListViewItemType.DataItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim litOCID As Literal = e.Item.FindControl("litOCID")
                Dim litClassCName As Literal = e.Item.FindControl("litClassCName")
                Dim litTAddress As Literal = e.Item.FindControl("litTAddress")
                'Dim hOCCU_DESC As String = drv("OCCU_DESC").ToString()
                'Dim hCLASSCNAME As String = drv("CLASSCNAME").ToString()

                Dim classItemUL As HtmlGenericControl = e.Item.FindControl("classItemUL")
                If (classItemUL IsNot Nothing) Then classItemUL.Attributes.Add("data-id", drv("ocid").ToString())
                If (litOCID IsNot Nothing) Then litOCID.Text = drv("ocid").ToString()

                Dim s_ClassCName As String = String.Format("{0}_{1}({2})", drv("CLASSCNAME").ToString(), drv("CJOBNAME").ToString(), drv("ORGNAME").ToString())
                If (litClassCName IsNot Nothing) Then litClassCName.Text = s_ClassCName

                Dim s_TAddress As String = get_TrainAddress(drv)
                If txtaddress2.Text <> "" Then s_TAddress = txtaddress2.Text
                If (litTAddress IsNot Nothing) Then litTAddress.Text = s_TAddress

                'function SubOCIDAddress(ocid1, address1)
                Dim s_scriptFMT1 As String = String.Format("SubOCIDAddress('{0}','{1}');", drv("ocid"), s_TAddress)
                Dim btnselect1 As HtmlButton = e.Item.FindControl("btnselect1")
                btnselect1.Attributes("onclick") = s_scriptFMT1
        End Select

    End Sub

    Sub search2()
        txtocid2.Text = TIMS.ClearSQM(txtocid2.Text)
        txtaddress2.Text = TIMS.ClearSQM(txtaddress2.Text)
        Dim parms As New Hashtable
        parms.Add("OCID", txtocid2.Text)
        'parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("Years", sm.UserInfo.Years)

        '查詢沒有座標資訊的課程
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select top 6 cc.OCID,cc.CLASSCNAME" & vbCrLf
        sql &= " ,sj.CJOBNAME2 CJOBNAME" & vbCrLf
        sql &= " ,oo.ORGNAME" & vbCrLf
        sql &= " ,cc.TADDRESSZIP" & vbCrLf
        sql &= " ,iz.CTNAME,iz.ZNAME,iz.ZIPNAME,cc.TADDRESS" & vbCrLf
        sql &= " ,cc.TWD97_X,cc.TWD97_Y" & vbCrLf
        sql &= " from CLASS_CLASSINFO cc" & vbCrLf
        sql &= " join PLAN_PLANINFO pp on pp.PLANID=cc.PLANID and pp.COMIDNO=cc.COMIDNO and pp.SEQNO=cc.SEQNO" & vbCrLf
        sql &= " join v_SHARECJOB sj on sj.CJOB_UNKEY=cc.CJOB_UNKEY" & vbCrLf
        sql &= " join org_orginfo oo on oo.comidno=cc.comidno" & vbCrLf
        sql &= " join view_plan ip on ip.planid=cc.planid" & vbCrLf
        sql &= " join view_zipname iz on iz.zipcode=cc.taddresszip" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND CC.ISSUCCESS='Y'" & vbCrLf
        sql &= " AND CC.NOTOPEN='N'" & vbCrLf
        '地址座標為空
        'sql &= " and (cc.TWD97_X is null or cc.TWD97_Y is null)" & vbCrLf
        sql &= " and cc.OCID=@OCID" & vbCrLf
        'sql &= " and ip.tplanid =@TPLANID" & vbCrLf
        sql &= " and ip.years =@Years" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        litMsg.Text = String.Format("共 {0} 筆待地址定位", dt.Rows.Count)

        lv_classQueryResult.DataSource = dt ';//結果資料集
        lv_classQueryResult.DataBind()

    End Sub

    Protected Sub btnSch2_Click(sender As Object, e As EventArgs) Handles btnSch2.Click
        Call search2()
    End Sub

    Protected Sub btnSch3_Click(sender As Object, e As EventArgs) Handles btnSch3.Click
        txtocid2.Text = ""
        txtaddress2.Text = ""
        Call search1()
    End Sub
End Class