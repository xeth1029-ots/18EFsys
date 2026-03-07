Partial Class CM_03_004
    Inherits AuthBasePage

    'Mode.SelectedIndex
    '--失業週數:6(Mode.SelectedIndex)
    '上課時數:7(Mode.SelectedIndex)
    Const cst_ModeT1_身分別 As String = "1"
    Const cst_ModeT1_年齡 As String = "2"
    Const cst_ModeT1_訓練職類 As String = "3"
    Const cst_ModeT1_教育程度 As String = "4"
    Const cst_ModeT1_通俗職類 As String = "5"
    Const cst_ModeT1_開班訓練地點 As String = "6"
    'Const cst_ModeT1_失業週數 As String="6"
    Const cst_ModeT1_上課時數 As String = "8"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            'msg.Text=""
            CreateItem()

            FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
            FTDate2.Text = TIMS.Cdate3(Now.Date)
            OCID.Style("display") = "none"
            msg.Text = TIMS.cst_NODATAMsg11

            'DistID.Attributes("onclick")="ClearData();"
            'TPlanID.Attributes("onclick")="ClearData();"
            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        End If

        'Button2.Attributes("onclick")="GetOrg();"
        Button3.Style("display") = "none"
        Button4.Visible = False
        'DistID.Enabled=True
        If sm.UserInfo.DistID <> "000" Then
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If

    End Sub

    Sub CreateItem()
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))
        'DistID.Items(0).Selected=True
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        '預算來源
        BudgetList = TIMS.Get_Budget(BudgetList, 3)
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trim(STDate1.Text) <> "" Then
            STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            STDate1.Text = ""
        End If

        If Trim(STDate2.Text) <> "" Then
            STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            STDate2.Text = ""
        End If

        If Errmsg = "" Then
            If STDate1.Text.ToString <> "" AndAlso STDate2.Text.ToString <> "" Then
                If CDate(STDate1.Text) > CDate(STDate2.Text) Then
                    Errmsg += "【開訓區間】的起日不得大於【開訓區間】的迄日!!" & vbCrLf
                End If
            End If
        End If

        If Trim(FTDate1.Text) <> "" Then
            FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate1.Text = ""
        End If

        If Trim(FTDate2.Text) <> "" Then
            FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate2.Text = ""
        End If

        If Errmsg = "" Then
            If FTDate1.Text.ToString <> "" AndAlso FTDate2.Text.ToString <> "" Then
                If CDate(FTDate1.Text) > CDate(FTDate2.Text) Then
                    Errmsg += "【結訓區間】的起日不得大於【結訓區間】的迄日!!" & vbCrLf
                End If
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim dt As DataTable
        Dim sql As String = ""
        Dim SearchStr3 As String = ""
        Dim itemstr As String = ""
        '選擇轄區
        itemstr = ""
        For Each objitem As ListItem In Me.DistID.Items
            If objitem.Selected = True Then
                If itemstr <> "" Then itemstr += ","
                itemstr += "'" & objitem.Value & "'"
            End If
        Next
        If itemstr <> "" Then
            SearchStr3 += " and ip.DistID IN (" & itemstr & ")" & vbCrLf
        End If

        '選擇計畫
        itemstr = ""
        For Each objitem As ListItem In Me.TPlanID.Items
            If objitem.Selected = True Then
                If itemstr <> "" Then itemstr += ","
                itemstr += "'" & objitem.Value & "'"
            End If
        Next
        If itemstr <> "" Then
            SearchStr3 += " and ip.TPlanID IN (" & itemstr & ")" & vbCrLf
        End If

        '業務ID
        If RIDValue.Value <> "" Then
            SearchStr3 += " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
        End If
        '計畫ID
        If PlanID.Value <> "" Then
            SearchStr3 += " and ip.PlanID='" & PlanID.Value & "'" & vbCrLf
        End If

        '選擇班級
        Dim SelOCIDflag As Boolean = False '是否有選擇班級 false:沒有 true:有
        itemstr = ""
        For Each objitem As ListItem In OCID.Items
            If objitem.Selected = True Then
                If objitem.Value = "%" Then
                    itemstr = ""
                    Exit For
                Else
                    If itemstr <> "" Then itemstr += ","
                    itemstr += "" & objitem.Value & ""
                End If
            End If
        Next
        If itemstr <> "" Then
            SelOCIDflag = True
        End If

        '是否有選擇班級 false:沒有 true:有
        If SelOCIDflag Then
            SearchStr3 += " and cc.OCID IN (" & itemstr & ")" & vbCrLf
        End If
        '是否有選擇班級 false:沒有 true:有
        If Not SelOCIDflag Then
            '開訓區間
            If STDate1.Text <> "" Then
                SearchStr3 += " and cc.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf '" & STDate1.Text & "'" & vbCrLf
            End If
            If STDate2.Text <> "" Then
                SearchStr3 += " and cc.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
            End If
            '結訓區間
            If FTDate1.Text <> "" Then
                SearchStr3 += " and cc.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
            End If
            If FTDate2.Text <> "" Then
                SearchStr3 += " and cc.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
            End If
        End If

        '選擇預算來源
        'itemstr=""
        itemstr = TIMS.CombiSQM2IN(TIMS.GetCblValue(BudgetList))
        'For Each objitem As ListItem In Me.BudgetList.Items
        '    If objitem.Selected=True Then
        '        If itemstr <> "" Then itemstr += ","
        '        itemstr += "'" & objitem.Value & "'"
        '    End If
        'Next
        If itemstr <> "" Then
            SearchStr3 += " and cs.BudgetID in (" & itemstr & ")"
        End If

        Dim v_Modet1 As String = TIMS.GetListValue(rbl_ModeT1)
        Select Case v_Modet1 'rbl_ModeT1.SelectedIndex
            Case cst_ModeT1_身分別 '身分別
                sql = "" & vbCrLf
                sql &= " SELECT a.IdentityID" & vbCrLf
                sql &= " ,a.Name Title" & vbCrLf
                sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
                sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
                sql &= " FROM (" & vbCrLf
                sql &= "   SELECT IDENTITYID ,NAME FROM dbo.KEY_IDENTITY WITH(NOLOCK)" & vbCrLf
                sql &= "   WHERE 1=1" & vbCrLf
                If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    sql &= "   AND IDENTITYID IN (" & TIMS.Cst_Identity06_2019_11 & ")" & vbCrLf
                Else
                    sql &= "   AND IDENTITYID IN (" & TIMS.Cst_Identity28_2019_11 & ")" & vbCrLf
                End If
                'sql &= "   union select '05'+KNID IdentityID, '　'+Name name from Key_Native" & vbCrLf
                'sql &= "   union select '05'+'20' IdentityID, '　'+N'族別未填' name " & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " left join (" & vbCrLf
                sql &= " select cs.MIdentityID" & vbCrLf
                sql &= " ,ISNULL(sum(case when ss.Sex='M' then 1 end),0) as MStudent" & vbCrLf
                sql &= " ,ISNULL(sum(case when ss.Sex='F' then 1 end),0) as FStudent" & vbCrLf
                sql &= " FROM dbo.id_plan ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= SearchStr3 & vbCrLf
                sql &= " Group By cs.MIdentityID" & vbCrLf
                'sql &= " union " & vbCrLf
                'sql &= " select cs.MIdentityID+ISNULL(cs.Native,'20') MIdentityID " & vbCrLf
                'sql &= " ,ISNULL(sum(case when ss.Sex='M' then 1 end),0) as MStudent" & vbCrLf
                'sql &= " ,ISNULL(sum(case when ss.Sex='F' then 1 end),0) as FStudent" & vbCrLf
                'sql &= " FROM dbo.id_plan ip" & vbCrLf
                'sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                'sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                'sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                'sql &= " and cs.MIdentityID='05' " & vbCrLf
                'sql &= " where 1=1" & vbCrLf
                'sql &= SearchStr3 & vbCrLf
                'sql &= " Group By cs.MIdentityID+ISNULL(cs.Native,'20')" & vbCrLf
                sql &= " ) b on a.IdentityID=b.MIdentityID" & vbCrLf
                sql &= " order by a.IdentityID" & vbCrLf

            Case cst_ModeT1_年齡 '年齡
                Dim sSql_YearsOld As String = ""
                sSql_YearsOld = " SELECT '15-19歲' AS Title,1 as Sort"
                sSql_YearsOld &= " UNION SELECT '20-24歲' AS Title,2 as Sort"
                sSql_YearsOld &= " UNION SELECT '25-29歲' AS Title,3 as Sort"
                sSql_YearsOld &= " UNION SELECT '30-34歲' AS Title,4 as Sort"
                sSql_YearsOld &= " UNION SELECT '35-39歲' AS Title,5 as Sort"
                sSql_YearsOld &= " UNION SELECT '40-44歲' AS Title,6 as Sort"
                sSql_YearsOld &= " UNION SELECT '45-49歲' AS Title,7 as Sort"
                sSql_YearsOld &= " UNION SELECT '50-54歲' AS Title,8 as Sort"
                sSql_YearsOld &= " UNION SELECT '55-59歲' AS Title,9 as Sort"
                sSql_YearsOld &= " UNION SELECT '60-64歲' AS Title,10 as Sort"
                sSql_YearsOld &= " UNION SELECT '65歲以上' AS Title,11 as Sort"

                sql = " SELECT a.Title" & vbCrLf
                sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
                sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
                sql &= String.Concat(" from (", sSql_YearsOld, ") a", vbCrLf)
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= "  select g.Sort" & vbCrLf
                sql &= "  ,sum(case when g.Sex='M' then 1 end) MStudent" & vbCrLf
                sql &= "  ,sum(case when g.Sex='F' then 1 end) FStudent" & vbCrLf
                sql &= "  from (" & vbCrLf
                sql &= "    SELECT ss.Sex" & vbCrLf
                sql &= "    ,dbo.FN_YEARSOLD3D(cc.FTDate,ss.Birthday) Sort" & vbCrLf
                sql &= " FROM dbo.id_plan ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= SearchStr3 & vbCrLf
                sql &= "   ) g" & vbCrLf
                sql &= "   GROUP BY g.Sort" & vbCrLf
                sql &= " ) b on b.Sort=a.Sort" & vbCrLf
                sql &= " order by a.Sort" & vbCrLf

            Case cst_ModeT1_訓練職類 '訓練職類
                sql = "" & vbCrLf
                sql &= " select a.Title" & vbCrLf
                sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
                sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select max(b3.tmkey) sort,b3.BusName+'-'+b3.JobName Title " & vbCrLf
                sql &= "   from VIEW_TRAINTYPE b3" & vbCrLf
                sql &= "   group by b3.BusName+'-'+b3.JobName" & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " left join (" & vbCrLf
                sql &= " select g.Title" & vbCrLf
                sql &= " ,sum(case when g.Sex='M' then 1 end) MStudent" & vbCrLf
                sql &= " ,sum(case when g.Sex='F' then 1 end) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select ss.Sex" & vbCrLf
                sql &= "   ,b3.BusName+'-'+b3.JobName Title" & vbCrLf

                sql &= " FROM dbo.id_plan ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_TRAINTYPE b3 on b3.tmid=cc.tmid" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= SearchStr3 & vbCrLf
                sql &= " ) g" & vbCrLf
                sql &= " Group By g.Title" & vbCrLf
                sql &= " ) b on a.Title=b.Title" & vbCrLf
                sql &= " order by a.Sort" & vbCrLf

            Case cst_ModeT1_教育程度 '教育程度
                sql = "" & vbCrLf
                sql &= " select a.Title" & vbCrLf
                sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
                sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select b3.Name Title" & vbCrLf
                sql &= "    ,b3.DegreeID+b3.Name sort " & vbCrLf
                sql &= "   from Key_Degree b3" & vbCrLf
                sql &= "   where b3.degreetype=1" & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " left join (" & vbCrLf
                sql &= " select g.sort" & vbCrLf
                sql &= " ,sum(case when g.Sex='M' then 1 end) MStudent" & vbCrLf
                sql &= " ,sum(case when g.Sex='F' then 1 end) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select ss.Sex ,b3.DegreeID+b3.Name sort " & vbCrLf
                sql &= " FROM dbo.id_plan ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                sql &= " LEFT JOIN dbo.KEY_DEGREE b3 on b3.DegreeID=ss.DegreeID" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= SearchStr3 & vbCrLf
                sql &= " ) g" & vbCrLf
                sql &= " Group By g.sort" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.Sort" & vbCrLf

            Case cst_ModeT1_通俗職類 '通俗職類
                '啟用2016年通俗職類
                Dim flag_Cjob2016 As Boolean = TIMS.Get_sCjob2016_USE(Page)
                Dim str_SHARECJOB_YEAR As String = ""
                If flag_Cjob2016 Then str_SHARECJOB_YEAR = TIMS.cst_SHARE_CJOB_2016

                sql = "" & vbCrLf
                sql &= " SELECT a.Title" & vbCrLf
                sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
                sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
                sql &= " FROM (" & vbCrLf

                Select Case str_SHARECJOB_YEAR
                    Case TIMS.cst_SHARE_CJOB_2016
                        sql &= "  SELECT b3.CJOB_NAME Title " & vbCrLf
                        sql &= "  ,b3.CJOB_UNKEY SORT " & vbCrLf '依此JOIN 
                        sql &= "  ,ISNULL(b3.CJOB_NO,b3.CJOB_TYPE)+b3.JOB_NO SORT2" & vbCrLf '依此排序
                        sql &= "  FROM SHARE_CJOB b3" & vbCrLf
                        sql &= "  WHERE b3.CYEARS ='2019'" & vbCrLf
                    Case Else
                        sql &= "  select b3.CJOB_NAME Title " & vbCrLf
                        sql &= "  ,b3.CJOB_UNKEY SORT " & vbCrLf '依此JOIN 
                        sql &= "  ,NVL2(b3.CJOB_NO , LPAD(b3.CJOB_NO, 4, '0'), LPAD(b3.CJOB_TYPE, 2, '0')) SORT2" & vbCrLf '依此排序
                        sql &= "  FROM SHARE_CJOB b3" & vbCrLf
                        sql &= "  WHERE b3.CYEARS ='2014'" & vbCrLf
                End Select

                sql &= " ) a" & vbCrLf
                sql &= " left join (" & vbCrLf
                sql &= " select g.sort" & vbCrLf
                sql &= " ,sum(case when g.Sex='M' then 1 end) MStudent" & vbCrLf
                sql &= " ,sum(case when g.Sex='F' then 1 end) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select ss.Sex" & vbCrLf
                'sql += "   ,b3.CJOB_NO+b3.CJOB_Name sort " & vbCrLf
                sql &= "  ,b3.CJOB_UNKEY SORT " & vbCrLf '依此JOIN 
                sql &= " FROM dbo.id_plan ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                sql &= " JOIN dbo.SHARE_CJOB b3 on b3.CJOB_UNKEY=cc.CJOB_UNKEY" & vbCrLf
                sql &= "   where 1=1" & vbCrLf
                sql &= SearchStr3 & vbCrLf
                sql &= "   ) g" & vbCrLf
                sql &= " Group By g.sort" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.SORT2" & vbCrLf

            Case cst_ModeT1_開班訓練地點 '開班訓練地點
                sql = "" & vbCrLf
                sql &= " select a.Title" & vbCrLf
                sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
                sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select b3.CTID sort,b3.CTName Title" & vbCrLf
                sql &= "   from dbo.ID_City b3" & vbCrLf
                sql &= "   where b3.CTID <> 999" & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " left join (" & vbCrLf
                sql &= " select g.sort" & vbCrLf
                sql &= " ,sum(case when g.Sex='M' then 1 end) MStudent" & vbCrLf
                sql &= " ,sum(case when g.Sex='F' then 1 end) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select ss.Sex" & vbCrLf
                sql &= "   ,b3.CTID sort " & vbCrLf
                sql &= " FROM dbo.id_plan ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_ZIPNAME b3 on b3.zipcode=cc.TaddressZip" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= SearchStr3 & vbCrLf
                sql &= " ) g" & vbCrLf
                sql &= " Group By g.sort" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.Sort" & vbCrLf

            'Case cst_ModeT1_失業週數 '6 '失業週數:6(Mode.SelectedIndex)
            '    sql="" & vbCrLf
            '    sql &= " select a.Title" & vbCrLf
            '    sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
            '    sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
            '    sql &= " from (" & vbCrLf
            '    sql &= "  select 1 sort, '(失業25週以下)' Title " & vbCrLf
            '    sql &= "  union select 2 sort, '(失業26週以上)' Title " & vbCrLf
            '    sql &= " ) a" & vbCrLf
            '    sql &= " left join (" & vbCrLf
            '    sql &= " select g.sort" & vbCrLf
            '    sql &= " ,sum(case when g.Sex='M' then 1 end) MStudent" & vbCrLf
            '    sql &= " ,sum(case when g.Sex='F' then 1 end) FStudent" & vbCrLf
            '    sql &= " from (" & vbCrLf
            '    sql &= "   select ss.Sex" & vbCrLf
            '    sql &= "   ,ISNULL(st.LostJobWeek,1) sort " & vbCrLf
            '    sql &= " FROM dbo.id_plan ip" & vbCrLf
            '    sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
            '    sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
            '    sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
            '    sql &= " JOIN dbo.VIEW_ZIPNAME b3 on b3.zipcode=cc.TaddressZip" & vbCrLf
            '    sql &= " LEFT JOIN dbo.VIEW_LOSTJOBWEEK st on cs.socid=st.socid" & vbCrLf
            '    sql &= " where 1=1" & vbCrLf
            '    sql &= SearchStr3 & vbCrLf
            '    sql &= " ) g" & vbCrLf
            '    sql &= " Group By g.sort" & vbCrLf
            '    sql &= " ) b on a.sort=b.sort" & vbCrLf
            '    sql &= " order by a.Sort" & vbCrLf

            Case cst_ModeT1_上課時數 ' '上課時數:7(Mode.SelectedIndex)
                sql = "" & vbCrLf
                sql &= " select a.Title" & vbCrLf
                sql &= " ,ISNULL(b.MStudent,0) MStudent" & vbCrLf
                sql &= " ,ISNULL(b.FStudent,0) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "  select TrainHours1 sort" & vbCrLf
                sql &= "  ,TrainHours Title" & vbCrLf
                sql &= "  FROM dbo.V_TRAINHOURS" & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " left join (" & vbCrLf
                sql &= " select g.sort" & vbCrLf
                sql &= " ,sum(case when g.Sex='M' then 1 end) MStudent" & vbCrLf
                sql &= " ,sum(case when g.Sex='F' then 1 end) FStudent" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= "   select ss.Sex ,th.TrainHours1 sort " & vbCrLf
                sql &= " FROM dbo.id_plan ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_ZIPNAME b3 on b3.zipcode=cc.TaddressZip" & vbCrLf
                sql &= " LEFT JOIN dbo.V_CLASSTRAINHOURS th on th.OCID=cc.OCID" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= SearchStr3 & vbCrLf
                sql &= " ) g" & vbCrLf
                sql &= " Group By g.sort" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.Sort" & vbCrLf

        End Select

        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, TIMS.cst_ErrorMsg9) '"資料庫效能異常，請重新查詢")
            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg += "/*  sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End Try

        Me.Button4.Visible = False
        Me.DataGrid1.Visible = False
        Me.Datagrid2.Visible = False

        If dt.Rows.Count > 0 Then
            FrameTable3.Visible = False

            Me.Button4.Visible = True
            Me.DataGrid1.Visible = True
            Me.Datagrid2.Visible = True

            DataGrid1.PageSize = dt.Rows.Count
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If

        'CHANGE BY NICK 060525 搞啥! 功能沒寫出來就不要寫...不要寫了又半調子...後面的人改會很累的

        'Dim sql As String=""
        sql = "" & vbCrLf
        sql &= " select a.Title 年齡區間" & vbCrLf
        sql &= " ,ISNULL(b.MStudent,0) 人數" & vbCrLf
        sql &= " from (" & vbCrLf
        sql &= " SELECT '25歲以下' Title,1 Sort" & vbCrLf
        sql &= " UNION SELECT '26歲-35歲' Title,2 Sort" & vbCrLf
        sql &= " UNION SELECT '36歲以上' Title,3 Sort" & vbCrLf '主TB
        sql &= " ) a" & vbCrLf

        sql &= " left join (" & vbCrLf
        sql &= " select g.Title" & vbCrLf
        sql &= " ,count(1) MStudent" & vbCrLf
        sql &= " from (" & vbCrLf
        sql &= " select ss.Sex" & vbCrLf
        sql &= " ,case when dbo.FN_YEARSOLD(cc.FTDate,ss.Birthday) <=25 then '25歲以下'" & vbCrLf
        sql &= "  when dbo.FN_YEARSOLD(cc.FTDate,ss.Birthday) <=35 then '26歲-35歲'" & vbCrLf
        sql &= "  else '36歲以上' end Title" & vbCrLf '資料TB 

        sql &= " FROM dbo.id_plan ip" & vbCrLf
        sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y' and cc.FTDate<GETDATE()" & vbCrLf
        sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID and cs.StudStatus Not IN (2,3) and cs.MakeSOCID is null" & vbCrLf
        sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
        'sql &= " JOIN dbo.VIEW_ZIPNAME b3 on b3.zipcode=cc.TaddressZip" & vbCrLf
        'sql &= " LEFT JOIN dbo.V_CLASSTRAINHOURS th on th.OCID=cc.OCID" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= SearchStr3 & vbCrLf
        sql &= " ) g" & vbCrLf
        sql &= " Group By g.Title" & vbCrLf
        sql &= " ) b on a.Title=b.Title" & vbCrLf

        sql &= " order by a.Sort" & vbCrLf

        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, TIMS.cst_ErrorMsg9) '"資料庫效能異常，請重新查詢")
            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg += "/*  sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End Try

        If dt.Rows.Count > 0 Then
            'Dim dr As DataRow
            Dim iSum As Integer = 0
            For Each dr As DataRow In dt.Rows
                iSum += dr("人數")
            Next
            Dim dr2 As DataRow = dt.NewRow()
            dr2("年齡區間") = "合計"
            dr2("人數") = iSum
            dt.Rows.Add(dr2)
            dt.AcceptChanges()
        End If

        Datagrid2.PageSize = dt.Rows.Count
        Datagrid2.DataSource = dt
        Datagrid2.DataBind()

        Page.RegisterStartupScript("load", "<script>ReStart();</script>")
    End Sub

    Private Sub DataGrid1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemCreated
        Dim MyCell As TableCell
        Select Case e.Item.ItemType
            Case ListItemType.Pager
                e.Item.Cells.Clear()
                e.Item.CssClass = "CM_TR1"

                MyCell = New TableCell
                MyCell.RowSpan = 2
                MyCell.Text = "統計項目"
                e.Item.Cells.Add(MyCell)

                MyCell = New TableCell
                MyCell.ColumnSpan = 4
                MyCell.Text = "性別"
                MyCell.Width = Unit.Pixel(320)
                e.Item.Cells.Add(MyCell)
            Case ListItemType.Header
                e.Item.Cells.Clear()
                e.Item.CssClass = "CM_TR1"

                MyCell = New TableCell
                MyCell.Text = "男"
                MyCell.Width = Unit.Pixel(80)
                e.Item.Cells.Add(MyCell)
                MyCell = New TableCell
                MyCell.Text = "女"
                MyCell.Width = Unit.Pixel(80)
                e.Item.Cells.Add(MyCell)
                MyCell = New TableCell
                MyCell.Text = "總計"
                MyCell.Width = Unit.Pixel(80)
                e.Item.Cells.Add(MyCell)
                MyCell = New TableCell
                MyCell.Text = "百分比"
                MyCell.Width = Unit.Pixel(80)
                e.Item.Cells.Add(MyCell)
            Case ListItemType.Footer
                e.Item.HorizontalAlign = HorizontalAlign.Center
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim dt As DataTable = drv.Row.Table
                Dim dr As DataRow
                Dim total As Integer
                For Each dr In dt.Rows
                    If Left(dr("Title"), 1) <> "　" Then
                        total += Int(dr("MStudent")) + Int(dr("FStudent"))
                    End If
                Next
                e.Item.Cells(3).Text = Int(e.Item.Cells(1).Text) + Int(e.Item.Cells(2).Text)
                If total = 0 Then
                    e.Item.Cells(4).Text = "0%"
                Else
                    If Left(e.Item.Cells(0).Text, 1) <> "　" Then
                        e.Item.Cells(4).Text = Math.Round((Int(e.Item.Cells(1).Text) + Int(e.Item.Cells(2).Text)) / total * 100, 2) & "%"
                    Else
                        e.Item.Cells(4).Text = Math.Round((Int(e.Item.Cells(1).Text) + Int(e.Item.Cells(2).Text)) / total * 100, 2) & "%"
                        e.Item.Cells(0).ForeColor = Color.SlateBlue
                        e.Item.Cells(1).ForeColor = Color.SlateBlue
                        e.Item.Cells(2).ForeColor = Color.SlateBlue
                        e.Item.Cells(3).ForeColor = Color.SlateBlue
                        e.Item.Cells(4).ForeColor = Color.SlateBlue
                    End If
                End If
            Case ListItemType.Footer
                e.Item.Cells(0).HorizontalAlign = HorizontalAlign.Left
                e.Item.Cells(1).Text = 0
                e.Item.Cells(2).Text = 0
                e.Item.Cells(3).Text = 0
                For Each item As DataGridItem In DataGrid1.Items
                    If Left(item.Cells(0).Text, 1) <> "　" Then
                        e.Item.Cells(1).Text += Int(item.Cells(1).Text)
                        e.Item.Cells(2).Text += Int(item.Cells(2).Text)
                        e.Item.Cells(3).Text += Int(item.Cells(3).Text)

                    End If
                Next
                If e.Item.Cells(3).Text = "0" Then
                    e.Item.Cells(4).Text = "0%"
                Else
                    e.Item.Cells(4).Text = "100%"
                End If
        End Select
    End Sub

    '查詢班級。
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow

        msg.Text = ""
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        sql = "SELECT * FROM CLASS_CLASSINFO WHERE PlanID='" & PlanID.Value & "' and RID='" & RIDValue.Value & "' and NotOpen='N' and IsSuccess='Y'"
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            msg.Text = "查無此機構底下的班級"
            OCID.Style("display") = "none"
        Else
            OCID.Items.Clear()
            OCID.Items.Add(New ListItem("全選", "%"))
            For Each dr In dt.Rows
                Dim ClassName As String

                ClassName = dr("ClassCName").ToString
                If IsNumeric(dr("CyclType")) Then
                    If Int(dr("CyclType")) <> 0 Then
                        ClassName += "第" & Int(dr("CyclType")) & "期"
                    End If
                End If

                OCID.Items.Add(New ListItem(ClassName, dr("OCID")))
            Next
            OCID.Style("display") = "inline"
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Button4.Visible = False
        Me.DataGrid1.Visible = False
        Me.Datagrid2.Visible = False
        FrameTable3.Visible = True
    End Sub

    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick

        Dim DistID1 As String
        Dim TPlanID1 As String
        Dim N As Integer
        Dim N1 As Integer

        DistID1 = ""
        N = 0   '預設 N =0 表示沒有勾選轄區選項
        For i As Integer = 1 To Me.DistID.Items.Count - 1

            If Me.DistID.Items(i).Selected Then '假如有勾選
                N = N + 1  '計算轄區勾選選項的數目
                If N = 1 Then '如果是勾選一個選項
                    DistID1 = Convert.ToString(Me.DistID.Items(i).Value) '取得選項的值
                End If
                If N = 2 Then '如果轄區勾選選項的數目=2
                    Common.MessageBox(Me, "只能選擇一個轄區")
                    DistID1 = ""
                    Exit For
                End If
            End If
        Next

        If N = 0 Then '如果轄區選項沒有選
            Common.MessageBox(Me, "請選擇轄區")
        End If

        TPlanID1 = ""
        N1 = 0 '預設 N1 =0 表示沒有勾選計畫選項
        For j As Integer = 1 To Me.TPlanID.Items.Count - 1

            If Me.TPlanID.Items(j).Selected Then '假如有勾選
                N1 = N1 + 1 '計算計畫勾選選項的數目
                If N1 = 1 Then '如果是勾選一個選項
                    TPlanID1 = Convert.ToString(Me.TPlanID.Items(j).Value) '取得選項的值
                End If
                If N1 = 2 Then '如果計畫勾選選項的數目=2
                    Common.MessageBox(Me, "只能選擇一個計畫")
                    TPlanID1 = ""
                    Exit For
                End If

            End If
        Next

        If N = 0 Then '如果計畫選項沒有選
            Common.MessageBox(Me, "請選擇計畫")
        End If

        If DistID1 <> "" And TPlanID1 <> "" Then
            Dim strScript1 As String

            strScript1 = "<script language=""javascript"">" + vbCrLf
            strScript1 += "wopen('../../Common/MainOrg.aspx?DistID=' + '" & DistID1 & "' + '&TPlanID=' + '" & TPlanID1 & "'  + '&BtnName=Button3','查詢機構',400,400,1);"
            strScript1 += "</script>"
            Page.RegisterStartupScript("", strScript1)

        End If
    End Sub

End Class