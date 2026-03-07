Public Class SD_15_021
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Call sUtl_CloseAllTab()
            divSearch1.Visible = True

            yearlist1 = TIMS.GetSyear(yearlist1)
            Common.SetListItem(yearlist1, sm.UserInfo.Years)
            yearlist2 = TIMS.GetSyear(yearlist2)
            Common.SetListItem(yearlist2, sm.UserInfo.Years)

            Call TIMS.Get_DISTCBL(Distid, objconn)
            'Distid.Items.Insert(0, New ListItem("全部", 0))
            Distid.Attributes("onclick") = "SelectAll('Distid','DistHidden');"

            Distid.Enabled = True
            If sm.UserInfo.DistID <> "000" Then '若登入者非署(局)署，鎖定轄區
                Common.SetListItem(Distid, sm.UserInfo.DistID)
                'Distid.Enabled = False
            End If
        End If

    End Sub


    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call sUtl_CloseAllTab()
        Div1.Visible = True
        divSearch2.Visible = True

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ip.Years 年度" & vbCrLf
        sql &= " ,ip.distname 分署" & vbCrLf
        sql &= " ,a.CNAME 設備名稱" & vbCrLf
        sql &= " ,a.Price 單價" & vbCrLf
        sql &= " /*,SYS.DBMS_RANDOM.RANDOM 數量*/" & vbCrLf
        sql &= " ,trunc(DBMS_RANDOM.value(3,10)) 數量" & vbCrLf
        sql &= " ,trunc(a.Price*DBMS_RANDOM.value(5,10)*0.8) 總價" & vbCrLf
        sql &= " ,oo.orgname  委訓採購廠商" & vbCrLf
        sql &= " ,CONVERT(varchar, CC.STDATE, 111) 採購日期" & vbCrLf
        sql &= " ,a.purpose 備註" & vbCrLf
        sql &= " FROM PLAN_PERSONCOST a" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp on pp.planid=a.planid and pp.comidno=a.comidno and pp.seqno=a.seqno" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid=pp.planid" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno =pp.comidno" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        sql &= " JOIN VIEW_TRAINTYPE tt on tt.tmid =pp.tmid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and ip.TPLANID ='28'" & vbCrLf
        sql &= " AND PP.ISAPPRPAPER='Y'" & vbCrLf
        sql &= " and rownum <=100" & vbCrLf
        sql &= " order by    dbms_random.value" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        msg.Text = ""
        'DataGrid1.DataSource = dt
        'DataGrid1.DataBind()
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        Call sUtl_CloseAllTab()
        divSearch1.Visible = True 'False
    End Sub

    Sub sUtl_CloseAllTab()
        divSearch1.Visible = False 'True 'False
        Div1.Visible = False
        divSearch2.Visible = False
        divAdd1.Visible = False
    End Sub

    Protected Sub btnAdd1_Click(sender As Object, e As EventArgs) Handles btnAdd1.Click
        Call sUtl_CloseAllTab()
        divAdd1.Visible = True 'False

        ddlYears1 = TIMS.GetSyear(ddlYears1)
        Common.SetListItem(ddlYears1, sm.UserInfo.Years)

        Call TIMS.Show_DistList(ddlDistID1, objconn)
        Common.SetListItem(ddlDistID1, sm.UserInfo.DistID)
        ddlDistID1.Enabled = False
    End Sub

    Protected Sub btnSave2_Click(sender As Object, e As EventArgs) Handles btnSave2.Click
        Call sUtl_CloseAllTab()
        divSearch1.Visible = True 'False
    End Sub
End Class