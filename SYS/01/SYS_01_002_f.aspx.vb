Partial Class SYS_01_002_f
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            create1()
        End If

        'Lis_acc.Enabled = False
        btnClose.Attributes("onclick") = "return closeMe();"
    End Sub

    Sub create1()
        Lis_acc.Enabled = False
        'Request("RID")
        'Request("AN")
        'sm.UserInfo.RoleID
        Try
            Dim dt As DataTable
            Dim objstr As String = ""
            '取出帳號
            If Convert.ToString(Request("RID")) <> "" AndAlso Convert.ToString(Request("AN")) <> "" Then
                objstr = "" & vbCrLf
                objstr += " select a.Account,a.RoleID" & vbCrLf
                objstr += " ,a.Name+'('+b.Name+') ['+a.Account+']' Name" & vbCrLf
                objstr += " from Auth_Account a" & vbCrLf
                objstr += " join ID_Role b on a.RoleID=b.RoleID" & vbCrLf
                objstr += " join Auth_Relship c on c.OrgID=a.OrgID" & vbCrLf
                objstr += " where c.RID='" & Request("RID") & "'" & vbCrLf
                objstr += " and a.Account ='" & Request("AN") & "'" & vbCrLf
                'objstr += "and a.RoleID >=" & sm.UserInfo.RoleID & vbCrLf
                dt = DbAccess.GetDataTable(objstr, objconn)
                dt.DefaultView.Sort = "RoleID,Name"

                Me.Lis_acc.DataSource = dt
                Me.Lis_acc.DataTextField = "name"
                Me.Lis_acc.DataValueField = "account"
                Me.Lis_acc.DataBind()

                Common.SetListItem(Lis_acc, Request("AN"))

                objstr = "select RoleID from auth_account where account = '" & Me.Lis_acc.SelectedValue & "'"
                Dim TmpRoleID As String = Convert.ToString(DbAccess.ExecuteScalar(objstr, objconn))

                If sm.UserInfo.RoleID <= 1 Then       '假如登入者為系統管理者
                    objstr = "SELECT Distinct Years FROM view_LoginPlan WHERE DistID=(SELECT DistID FROM Auth_Relship WHERE RID='" & Convert.ToString(Request("RID")) & "') ORDER BY 1 DESC"
                Else
                    If TmpRoleID = 1 Then
                        objstr = "SELECT Distinct Years FROM view_LoginPlan WHERE DistID=(SELECT DistID FROM Auth_Relship WHERE RID='" & Convert.ToString(Request("RID")) & "') ORDER BY 1 DESC"
                    Else
                        objstr = "SELECT Distinct Years FROM view_LoginPlan WHERE PlanID='" & sm.UserInfo.PlanID & "' ORDER BY 1 DESC"
                    End If
                End If
                dt = DbAccess.GetDataTable(objstr, objconn)
                Years.DataSource = dt
                Years.DataTextField = "Years"
                Years.DataValueField = "Years"
                Years.DataBind()

                If Convert.ToString(Request("Years")) <> "" Then
                    Common.SetListItem(Years, Request("Years"))

                    getDataGrid1(Convert.ToString(Request("Years")), Convert.ToString(Request("AN")))

                End If
            End If

        Catch ex As Exception
            Dim strErrmsg As String = ""
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(Me, ex, strErrmsg)
            'Common.MessageBox(Me, ex.ToString)
            'Exit Sub
        End Try

    End Sub

    Private Sub Years_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Years.SelectedIndexChanged
        If Convert.ToString(Request("RID")) <> "" AndAlso Convert.ToString(Request("AN")) <> "" Then
            getDataGrid1(Years.SelectedValue, Lis_acc.SelectedValue)
        End If
    End Sub

    '取得所選取帳號目前已賦予之計畫
    '建置訓練計畫清單
    Private Sub getDataGrid1(ByVal YearsVal As String, ByVal AccountVal As String)
        Dim sql As String
        Dim dt As DataTable
        Try
            'sql = "" & vbCrLf
            'sql += " SELECT a.RID,b.TPlanID, b.PlanID,a.CreateByAcc" & vbCrLf
            'sql += " ,CASE WHEN o2.OrgName is null then b.Years+e.Name+f.PlanName+b.Seq" & vbCrLf
            'sql += " ELSE b.Years+e.Name+f.PlanName+b.Seq +'　('+o2.OrgName+')'" & vbCrLf
            'sql += " END AS PlanName" & vbCrLf
            'sql += " ,d.OrgID, d.OrgName, ac.isused, c.DistID, c.OrgLevel " & vbCrLf
            'sql += " ,a2.RID CRID,o2.OrgID COrgID, o2.OrgName COrgName" & vbCrLf
            'sql += " FROM " & vbCrLf
            'sql += " (SELECT * FROM Auth_AccRWPlan WHERE account='" & AccountVal & "') a" & vbCrLf
            'sql += "   JOIN ID_Plan b on a.PlanID=b.PlanID" & vbCrLf
            'sql += "   JOIN ( select * , CASE " & vbCrLf
            'sql += "        WHEN len(Relship)-4 >0 " & vbCrLf
            'sql += "        THEN replace(replace(substring(Relship,5,len(Relship)-4),RID,''),'/','') " & vbCrLf
            'sql += "        END AS CRID from Auth_Relship) c on a.RID=c.RID  " & vbCrLf
            'sql += "   JOIN Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
            'sql += "   JOIN ID_District e on b.DistID=e.DistID" & vbCrLf
            'sql += "   JOIN Key_Plan f on b.TPlanID=f.TPlanID" & vbCrLf
            'sql += "   JOIN auth_account ac on a.account=ac.account" & vbCrLf
            'sql += "   left join Auth_Relship a2" & vbCrLf
            'sql += "  	on a2.RID=c.CRID" & vbCrLf
            'sql += "   left join Org_OrgInfo o2 on a2.OrgID=o2.OrgID" & vbCrLf
            'sql += " WHERE b.Years='" & YearsVal & "'" & vbCrLf

            sql = "" & vbCrLf
            sql += " SELECT a.RID,b.TPlanID, b.PlanID,a.CreateByAcc" & vbCrLf
            sql += " ,CASE WHEN o2.OrgName is null then b.Years+e.Name+f.PlanName+b.Seq" & vbCrLf
            sql += "   ELSE b.Years+e.Name+f.PlanName+b.Seq +'　('+o2.OrgName+')' END AS PlanName" & vbCrLf
            sql += " ,d.OrgID, d.OrgName, ac.isused" & vbCrLf
            sql += " , c.DistID, c.OrgLevel" & vbCrLf
            sql += " ,a2.RID2 CRID" & vbCrLf
            sql += " ,o2.OrgID COrgID" & vbCrLf
            sql += " ,o2.OrgName COrgName" & vbCrLf
            sql += " FROM Auth_AccRWPlan a" & vbCrLf
            sql += " join ID_Plan b on a.PlanID=b.PlanID and a.account='" & AccountVal & "'" & vbCrLf
            sql += " join Auth_Relship c on a.RID=c.RID" & vbCrLf
            sql += " join Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
            sql += " join ID_District e on b.DistID=e.DistID" & vbCrLf
            sql += " join Key_Plan f on b.TPlanID=f.TPlanID" & vbCrLf
            sql += " join auth_account ac on a.account=ac.account" & vbCrLf
            sql += " left join MVIEW_RELSHIP23 a2 on a2.RID3=c.RID" & vbCrLf
            sql += " left join Org_OrgInfo o2 on a2.OrgID2=o2.OrgID" & vbCrLf
            sql += " WHERE b.Years='" & YearsVal & "'" & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)

            msg.Text = "查無資料"
            DataGrid1.Visible = False
            If dt.Rows.Count > 0 Then
                msg.Text = ""
                DataGrid1.Visible = True
                DataGrid1.DataSource = dt
                DataGrid1.DataBind()
            End If

        Catch ex As Exception
            msg.Text = "查詢資料有誤！！"
            DataGrid1.Visible = False

            Dim strErrmsg As String = ""
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(Me, ex, strErrmsg)
            'Common.MessageBox(Me, ex.ToString)
        End Try

    End Sub
End Class
