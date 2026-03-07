Public Class SYS_04_015
    'Inherits System.Web.UI.Page
    Inherits AuthBasePage

    Const cst_RB_LID_1 As String = "1"
    Const cst_RB_LID_2 As String = "2"

    Dim giRow As Integer = 0

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

        'RB_LID 1:分署/2:機構
        Dim v_RB_LID As String = TIMS.GetListValue(RB_LID)
        v_RB_LID = If(v_RB_LID <> "", v_RB_LID, "2")
        Dim s_PLAN_s1 As String = ""
        Dim s_PLAN_w2 As String = ""
        Select Case v_RB_LID
            Case cst_RB_LID_1 '"1"
                s_PLAN_s1 = " LEFT JOIN dbo.VIEW_PLAN ip on ip.planid =rr.planid"
                s_PLAN_w2 = " and ip.PLANID is null"
            Case Else '"2"
                s_PLAN_s1 = " JOIN dbo.VIEW_PLAN ip on ip.planid =rr.planid"
                s_PLAN_w2 = ""
        End Select

        Dim parms As New Hashtable
        '查詢沒有座標資訊的機構
        Dim sql As String = ""
        sql &= String.Format(" SELECT TOP {0} op.RSID", s_ROWNUM) & vbCrLf
        sql &= " ,oo.ORGNAME" & vbCrLf
        sql &= " ,ip.PLANNAME" & vbCrLf
        sql &= " ,op.ZIPCODE,op.ZIPCODE6W" & vbCrLf
        sql &= " ,iz.CTNAME,iz.ZNAME" & vbCrLf
        sql &= " ,iz.ZIPNAME,op.ADDRESS OADDRESS" & vbCrLf
        sql &= " ,op.TWD97_X,op.TWD97_Y" & vbCrLf
        sql &= " FROM dbo.ORG_ORGPLANINFO op" & vbCrLf
        sql &= " JOIN dbo.VIEW_ZIPNAME iz on iz.zipcode=op.ZIPCODE" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP rr on rr.rsid=op.rsid" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo on oo.orgid=rr.orgid" & vbCrLf
        sql &= s_PLAN_s1 & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= s_PLAN_w2 & vbCrLf
        '地址座標為空
        sql &= " and (op.TWD97_X is null or op.TWD97_Y is null)" & vbCrLf
        Select Case v_RB_LID
            Case cst_RB_LID_1
            Case Else
                'parms.Add("TPLANID", sm.UserInfo.TPlanID)
                parms.Add("Years", sm.UserInfo.Years)
                'sql &= " and ip.tplanid =@TPLANID" & vbCrLf
                sql &= " and ip.years =@Years" & vbCrLf
        End Select
        sql &= " ORDER BY op.RSID DESC" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        litMsg.Text = String.Format("共 {0} 筆待地址定位", dt.Rows.Count)

        giRow = 0
        lv_orgQueryResult.DataSource = dt ';//結果資料集
        lv_orgQueryResult.DataBind()

    End Sub

    Private Sub lv_classQueryResult_ItemDataBound(sender As Object, e As ListViewItemEventArgs) Handles lv_orgQueryResult.ItemDataBound
        Select Case e.Item.ItemType
            Case ListViewItemType.DataItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim litRSID As Literal = e.Item.FindControl("litRSID")
                Dim litOrgName As Literal = e.Item.FindControl("litOrgName")
                Dim litTAddress As Literal = e.Item.FindControl("litTAddress")

                Dim classItemUL As HtmlGenericControl = e.Item.FindControl("classItemUL")
                If (classItemUL IsNot Nothing) Then classItemUL.Attributes.Add("data-id", drv("RSID").ToString())
                If (litRSID IsNot Nothing) Then litRSID.Text = drv("RSID").ToString()

                giRow += 1
                Dim s_PLANNAME As String = If(drv("PLANNAME").ToString() <> "", drv("PLANNAME").ToString(), "全計畫")
                Dim s_OrgName1 As String = String.Format("{0}.{1}({2})", giRow, drv("ORGNAME").ToString(), s_PLANNAME)
                If (litOrgName IsNot Nothing) Then litOrgName.Text = s_OrgName1

                Dim s_OAddress As String = get_orgAddress(drv)
                If txtaddress2.Text <> "" Then s_OAddress = txtaddress2.Text
                If (litTAddress IsNot Nothing) Then litTAddress.Text = s_OAddress

                ' function SubRSIDAddress(rsid1, address1)
                Dim s_scriptFMT1 As String = String.Format("SubRSIDAddress('{0}','{1}');", drv("RSID"), s_OAddress)
                Dim btnselect1 As HtmlButton = e.Item.FindControl("btnselect1")
                btnselect1.Attributes("onclick") = s_scriptFMT1
        End Select

    End Sub

    ''' <summary>
    ''' 組合機構地址 但地址可能有重複字，則排除
    ''' </summary>
    ''' <param name="drv"></param>
    ''' <returns></returns>
    Function get_orgAddress(ByRef drv As DataRowView) As String
        Dim rst As String = ""
        Dim ctname As String = drv("CTNAME").ToString()
        Dim zname As String = drv("ZNAME").ToString()
        Dim zipname As String = drv("ZIPNAME").ToString()
        Dim OADDRESS As String = drv("OADDRESS").ToString()
        If OADDRESS.StartsWith(ctname) OrElse OADDRESS.Contains(zname) OrElse OADDRESS.StartsWith(zipname) Then
            rst = OADDRESS
            Return rst
        End If
        rst = String.Format("{0}{1}", drv("ZIPNAME").ToString(), drv("OADDRESS").ToString())
        Return rst
    End Function

    Sub search2()
        txtRSID2.Text = TIMS.ClearSQM(txtRSID2.Text)
        txtaddress2.Text = TIMS.ClearSQM(txtaddress2.Text)

        'RB_LID 1:分署/2:機構
        Dim v_RB_LID As String = TIMS.GetListValue(RB_LID)
        v_RB_LID = If(v_RB_LID <> "", v_RB_LID, "2")
        Dim s_PLAN_s1 As String = ""
        Dim s_PLAN_w2 As String = ""
        Select Case v_RB_LID
            Case cst_RB_LID_1 '"1"
                s_PLAN_s1 = " LEFT JOIN dbo.VIEW_PLAN ip on ip.planid =rr.planid"
                s_PLAN_w2 = " and ip.PLANID is null"
            Case Else '"2"
                s_PLAN_s1 = " JOIN dbo.VIEW_PLAN ip on ip.planid =rr.planid"
                s_PLAN_w2 = ""
        End Select

        Dim parms As New Hashtable
        parms.Add("RSID", txtRSID2.Text)
        '查詢沒有座標資訊的課程
        Dim sql As String = ""
        sql &= " SELECT TOP 6 op.RSID" & vbCrLf
        sql &= " ,oo.ORGNAME" & vbCrLf
        sql &= " ,ip.PLANNAME" & vbCrLf
        sql &= " ,op.ZIPCODE,op.ZIPCODE6W" & vbCrLf
        sql &= " ,iz.CTNAME,iz.ZNAME" & vbCrLf
        sql &= " ,iz.ZIPNAME,op.ADDRESS OADDRESS" & vbCrLf
        sql &= " ,op.TWD97_X,op.TWD97_Y" & vbCrLf
        sql &= " FROM dbo.ORG_ORGPLANINFO op" & vbCrLf
        sql &= " JOIN dbo.VIEW_ZIPNAME iz on iz.zipcode=op.ZIPCODE" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP rr on rr.rsid=op.rsid" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo on oo.orgid=rr.orgid" & vbCrLf
        sql &= s_PLAN_s1 & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= s_PLAN_w2 & vbCrLf
        '地址座標為空
        sql &= " and (op.TWD97_X is null or op.TWD97_Y is null)" & vbCrLf
        sql &= " and op.RSID =@RSID" & vbCrLf
        'Select Case v_RB_LID
        '    Case cst_RB_LID_1
        '    Case Else
        '        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        '        parms.Add("Years", sm.UserInfo.Years)
        '        sql &= " and ip.tplanid =@TPLANID" & vbCrLf
        '        sql &= " and ip.years =@Years" & vbCrLf
        'End Select
        sql &= " ORDER BY op.RSID DESC" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        litMsg.Text = String.Format("共 {0} 筆待地址定位", dt.Rows.Count)

        giRow = 0
        lv_orgQueryResult.DataSource = dt ';//結果資料集
        lv_orgQueryResult.DataBind()

    End Sub

    Protected Sub btnSch2_Click(sender As Object, e As EventArgs) Handles btnSch2.Click
        Call search2()
    End Sub

    Protected Sub btnSch3_Click(sender As Object, e As EventArgs) Handles btnSch3.Click
        txtRSID2.Text = ""
        txtaddress2.Text = ""
        Call search1()
    End Sub
End Class