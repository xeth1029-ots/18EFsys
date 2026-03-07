Partial Class TC_03_003
    Inherits AuthBasePage

    '使用此程式計畫有 TPlanID : 28.54.(其餘計畫暫不提供使用。)
    '訓練業別 \TIMS.NET40o\js\OpenWin\openwin.js : openTrain
    'wopen('../../Common/TrainJob.aspx?field=TB_career_id&amp;TMID='+TMID,'TrainJob',420,200,0);

    'Imports System
    'Imports System.Web
    'Imports System.Drawing
    'Imports System.IO
    'Imports System.Drawing.Imaging

    '--UPDATE
    'SELECT * FROM PLAN_ONCLASS where rownum <=10
    'SELECT * FROM PLAN_TRAINDESC where rownum <=10
    'Plan_PlanInfo '計畫主檔
    'Plan_TrainDesc '課程資料(課程大綱)
    'x Teach_TeacherInfo '師資資料檔
    'PLAN_COSTITEM '經費資料
    'SELECT COSTID+','+ITEMCOSTNAME COSTID,COSTNAME FROM KEY_COSTITEM2 ORDER BY SORT

    'Plan_VerReport '訓練計劃開班總表(產學訓)
    'Plan_Teacher '班級申請老師檔(產學訓)
    'Plan_OnClass '計畫上課時間檔(產學訓)
    'Plan_TrainPlace '開班計畫場地檔(產學訓)
    'Plan_BusPackage '計畫包班事業單位(產學訓)
    'Plan_Material (停用)計畫材料品名項目檔
    'Plan_PersonCost '一人份材料明細(產學訓)
    'Plan_CommonCost '共同材料明細(產學訓)
    'Plan_SheetCost '教材費用 (產學訓)
    'Plan_OtherCost '其他費用 (產學訓)

    'org_orginfo
    'SELECT * FROM PLAN_ONCLASS where PlanID='1785'
    'select * from Plan_BusPackage where PlanID='1785'
    '--SELECT
    'Class_ClassInfo
    'Sys_GlobalVar
    'Key_TrainType
    'ID_GovClassCast

    'select * from Plan_PersonCost where planid =1871
    'select * from Plan_CommonCost where planid =1871
    'select * from plan_planinfo where planid =1871
    'select * from class_classinfo where planid =1871
    'select * from id_class  m
    'where exists (
    '	select 'x' from class_classinfo x where x.planid =1871
    '	and x.clsid=m.clsid
    ')
    'select * from view_plan where planid =1871
    'select * from org_orginfo where comidno ='49502521'
    'select * from plan_planinfo where comidno ='49502521'  and planid =1871 and seqno=9
    'SELECT * FROM PLAN_ONCLASS where PlanID='1871'
    'select * from Plan_BusPackage where PlanID='1871'
    'SELECT * FROM Plan_CostItem where PlanID='1871'
    'SELECT * FROM Plan_VerReport where PlanID='1871'
    'SELECT * FROM PLAN_TRAINDESC where PlanID='1871'
    'SELECT * FROM Plan_Teacher where PlanID='1871'
    'SELECT * FROM Teach_TeacherInfo  t
    'where exists (
    'SELECT 'x' FROM Plan_Teacher x where x.PlanID='1871'
    'and x.techid =t.techid
    ')
    'select * from Plan_TrainPlace m where exists (
    '		select 'x' from plan_planinfo x1 where x1.planid ='1871' and x1.addresssciptid=m.ptid
    '		union 
    '		select 'x' from plan_planinfo x2 where x2.planid ='1871' and x2.addresstechptid=m.ptid
    ')

    'Dim dr, dr1 As DataRow
    'Dim rqPlanID As String = ""
    'Dim sql As String = ""
    'Request("PlanID") /TIMS.ClearSQM(Request("ComIDNO") /TIMS.ClearSQM(Request("SeqNO") 有空值或異常:true
    Dim g_flagNG As Boolean = False

    Dim iPlanKind As Integer = 0
    Dim TPlanID As String = ""
    Dim iCostMode As Integer = 0
    Dim PlanID_value As String = ""
    Dim ComIDNO_value As String = ""
    Dim SeqNO_value As String = ""
    Dim dbld2TempTotal As Double = 0
    Dim tmpNoteDt As DataTable '暫存使用(有多個表格連續使用所以使用共用參數)
    Dim tmpPCS As String = "" '有儲存資料過了

    'Const cst_小計 = 3

    '2011 功能按鈕權限控管參數 ---------------------Start
    'Dim FunDr As DataRow = Nothing
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印
    '2011 功能按鈕權限控管參數 ---------------------End

    Const cst_titlemsg1 As String = "早上：7:00-13:00、下午：13:00-18:00、晚上：18:00-22:00"

    Const cst_errmsg1 As String = "程式出現例外狀況，請聯絡TIMS系統駐點人員!"
    Const cst_errmsg2 As String = "產生 課程申請流水號 有誤!!"
    Const cst_errmsg3 As String = "傳入參數異常，請重新查詢!!"
    Const cst_errmsg4 As String = "查詢時發生錯誤，請重新輸入查詢值!!"
    Const cst_errmsg5 As String = "查無資料，請重新確認查詢值!!"
    Const cst_errmsg6 As String = "儲存資料有誤!!!"
    Const cst_errmsg7 As String = "課程大綱 儲存資料有誤!!!"
    Const cst_errmsg8 As String = "計畫經費項目檔 儲存資料有誤!!!"
    Const cst_errmsg9 As String = "計畫材料品名項目檔 儲存資料有誤!!!"
    Const cst_errmsg10 As String = "一人份材料明細 儲存資料有誤!!!"
    Const cst_errmsg11 As String = "共同材料明細 儲存資料有誤!!!"
    Const cst_errmsg12 As String = "教材費用 儲存資料有誤!!!"
    Const cst_errmsg13 As String = "其他費用 儲存資料有誤!!!"
    Const cst_errmsg14 As String = "上課時間 儲存資料有誤!!!"
    Const cst_errmsg15 As String = "計畫包班事業單位 儲存資料有誤!!!"
    Const cst_errmsg16 As String = "傳入表格資訊有誤，刪除失敗!!"
    Const cst_errmsg17 As String = "找不到對應的場地代碼"
    Const cst_errmsg18 As String = "請勿嘗試在頁面輸入具有危險性的字元!"
    Const cst_errmsg19 As String = "未建立正式儲存資料，不可按「匯出EXCEL」!!!"
    Const cst_errmsg20 As String = "檢驗有誤!!"
    Const cst_errmsg21 As String = "該功能，不提供該登入計畫使用，若有需要，請先與系統管理者聯繫!!謝謝!!"
    Const cst_errmsg22 As String = "登入者無正確的業務權限，不提供儲存服務!(請勿在同一瀏覽器開不同視窗，同時登入不同計畫進行資料處理)"
    'Const cst_errmsg23 As String = "尚未基本儲存，開班計劃表資料維護無法儲存或輸入!"
    Const cst_errmsg24 As String = "課程大綱，為必填資料"
    Const cst_errmsg25 As String = "課程大綱內容資料，請重新確認!!"

    Const cst_errmsg35 As String = "班別資料「優先排序」欄位 同一個[申請階段]內，不可重複填寫相同數字!!"

    Const Cst_msgother1 As String = "( 學員資格* 請到 開班計畫表資料維護作業 )"
    Const Cst_msgother3 As String = "※請先確認有【一人份材料明細】或【共同材料明細】資料後，先按「正式儲存」，再按「匯出EXCEL」!! <br>　更新資料訓練費用編列說明，請按「匯出EXCEL」!!"
    Const Cst_msgother3b As String = "※更新資料訓練費用編列說明，請按其他說明「修改」!!"

    Const cst_CostItemTable As String = "CostItemTable" 'KEY_COSTITEM2 產投用經費項目
    Const cst_TrainDescTable As String = "TrainDescTable" '產學訓課程大綱
    Const cst_Plan_OnClass As String = "Plan_OnClass"
    Const cst_Plan_BusPackage As String = "Plan_BusPackage"

    Const Cst_MaterialTable As String = "MaterialTable"
    Const Cst_PMID As String = "PMID"
    'CreatePersonCost()
    Const Cst_PersonCostTable As String = "PersonCostTable" 'PersonCost: Plan_PersonCost–一人份材料明細
    Const Cst_PersonCostpkName As String = "ppcID"
    'CreateCommonCost()
    Const Cst_CommonCostTable As String = "CommonCostTable" 'CommonCost: Plan_CommonCost–共同材料明細
    Const Cst_CommonCostpkName As String = "pcmID"
    'CommonCost
    Const Cst_SheetCostTable As String = "SheetCostTable" 'SHEETCOST: PLAN_SHEETCOST–教材費用 (產學訓)
    Const Cst_SheetCostpkName As String = "pshID"
    'OtherCost
    Const Cst_OtherCostTable As String = "OtherCostTable" 'OTHERCOST: PLAN_OTHERCOST–其他費用 (產學訓)
    Const Cst_OtherCostpkName As String = "potID"

    Const cst_學分班 As String = "Y"
    Const cst_非學分班 As String = "N"
    '不管什麼都是「年滿15歲以上」。
    Const cst_AgeOtherDef As Integer = 16 'other Years Start

    Const cst_ccopy As String = "ccopy" 'Request(cst_ccopy)

    Dim strYears As String = "" '2014 / 2015'(經費分類代碼。)

#Region "NO USE"
    'Private Function GetGUID()
    '    Dim NEWGUID As System.Guid = System.Guid.NewGuid
    '    Return NEWGUID.ToString
    'End Function

    'Function CreatePICDT()
    '    Dim dt As DataTable
    '    If Not Session("dtPIC") Is Nothing Then
    '        dt = Session("dtPIC")
    '        Dim i As Integer
    '        Dim tmpstr As String
    '        Dim tmpary As Array
    '        Dim tmpflag101, tmpflag102, tmpflag201, tmpflag202 As Boolean
    '        tmpflag101 = True
    '        tmpflag102 = True
    '        tmpflag201 = True
    '        tmpflag202 = True
    '        For i = 0 To dt.Rows.Count - 1
    '            tmpstr += "," & dt.Rows(i).Item("depid")
    '        Next
    '        tmpary = Split(tmpstr, ",")
    '        With depID
    '            .Items.Clear()
    '            .Items.Add(New ListItem("==請選擇==", ""))
    '            For i = LBound(tmpary) To UBound(tmpary)
    '                If tmpary(i) = "101" Then tmpflag101 = False
    '                If tmpary(i) = "102" Then tmpflag102 = False
    '                If tmpary(i) = "201" Then tmpflag201 = False
    '                If tmpary(i) = "202" Then tmpflag202 = False
    '            Next
    '            If tmpflag101 Then .Items.Add(New ListItem("學科教室1", "101"))
    '            If tmpflag102 Then .Items.Add(New ListItem("學科教室2", "102"))
    '            If tmpflag201 Then .Items.Add(New ListItem("術科教室1", "201"))
    '            If tmpflag202 Then .Items.Add(New ListItem("術科教室2", "202"))
    '        End With
    '        If dt.Rows.Count = 0 Then
    '            DataGrid3Table.Visible = False
    '        Else
    '            DataGrid3Table.Visible = True
    '            DataGrid3.DataSource = dt
    '            DataGrid3.DataBind()
    '        End If
    '    End If
    'End Function

    'Session("saveok") 增加詢問
    'If Session("saveok") = True Then
    '    Session("saveok") = Nothing
    '    Page.RegisterStartupScript("計畫申請成功!!", "<SCRIPT>if(confirm('是否要繼續新增「開班計劃表資料維護」'))location.href='../01/TC_01_014.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "' ;else location.href='TC_03_003.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</SCRIPT>")
    '    'Page()
    'End If

    '檢查功能權限----------------------------Start
    'If sm.UserInfo.FunDt Is Nothing Then
    '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
    'Else
    '    Dim FunDt As DataTable = sm.UserInfo.FunDt
    '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & TIMS.ClearSQM(Request("ID")) & "'")
    '    FunDr = FunDrArray(0)
    '    If FunDr("Adds") = 1 Then
    '        btnAdd.Enabled = True
    '        Button8.Enabled = True
    '    Else
    '        btnAdd.Enabled = False
    '        Button8.Enabled = False
    '    End If
    'End If

    'Private Sub GCode1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Page.RegisterStartupScript("Londing", "<script>Layer_change(7);</script>")
    '    If GovClass.SelectedValue <> "" And GCode1.SelectedValue <> "" Then
    '        GCode2 = TIMS.Get_GCode2(GCode2, GovClass.SelectedValue, GCode1.SelectedValue)
    '    Else
    '        GCode2.Items.Clear()
    '        Common.MessageBox(Me, "請選擇有效資料")
    '    End If
    'End Sub

    'Private Sub Classification1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Classification1.SelectedIndexChanged
    '    Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    '    Select Case Classification1.SelectedValue
    '        Case "1" '學科
    '            If TIMS.ClearSQM(Request("ComIDNO") Is Nothing Then
    '                PTID = TIMS.Get_SciPTID(PTID, TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID.ToString))
    '            Else
    '                PTID = TIMS.Get_SciPTID(PTID, TIMS.ClearSQM(Request("ComIDNO"))
    '            End If
    '        Case "2" '術科
    '            If TIMS.ClearSQM(Request("ComIDNO") Is Nothing Then
    '                PTID = TIMS.Get_TechPTID(PTID, TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID.ToString))
    '            Else
    '                PTID = TIMS.Get_TechPTID(PTID, TIMS.ClearSQM(Request("ComIDNO"))
    '            End If
    '    End Select
    '    TIMS.Tooltip(PTID1, "上課地點以登入者的機構為準")
    '    TIMS.Tooltip(PTID2, "上課地點以登入者的機構為準")
    'End Sub

#End Region

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        Call sUtl_PageInit1()
        '檢查Session是否存在 End
        iPYNum = TIMS.sUtl_GetPYNum(Me)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso iPYNum >= 3 Then
            Dim rqID As String = TIMS.Get_MRqID(Me) 'Request(" 
            Dim grqPlanID As String = TIMS.sUtl_GetRqValue(Me, "PlanID")
            Dim grqComIDNO As String = TIMS.sUtl_GetRqValue(Me, "ComIDNO")
            Dim grqSeqNO As String = TIMS.sUtl_GetRqValue(Me, "SeqNO")
            Dim str6 As String = ""
            TIMS.SetMyValue(str6, "PlanID", grqPlanID)
            TIMS.SetMyValue(str6, "ComIDNO", grqComIDNO)
            TIMS.SetMyValue(str6, "SeqNO", grqSeqNO)
            Dim url1 As String = "TC_03_006.aspx?ID=" & rqID & str6
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        '(經費分類代碼。)
        strYears = "2014" '2014年  顯示層級。
        If sm.UserInfo.Years >= "2015" Then
            strYears = "2015" '2015年 不顯示層級。
        End If

        TableCost5.Visible = True '材料品名
        TableCost6.Visible = False '一人份材料明細
        TableCost7.Visible = False '共同材料明細
        If TIMS.Utl_GetConfigSet("work2013x01") = "Y" Then
            Labmsg3.Text = Cst_msgother3
            Note.ReadOnly = True
            Note.Style.Item("background-color") = "#BDBDBD"
            TableCost5.Visible = False '材料品名
            TableCost6.Visible = True
            TableCost7.Visible = True
        End If
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            hTPlanID54.Value = cst_errmsg21
            'Common.RespWrite(Me, "<script>alert('" & hTPlanID54.Value & "');</script>")
            'Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            'Response.End()
            'Exit Sub
            Dim sScript1 As String = ""
            sScript1 &= "<script>alert('" & hTPlanID54.Value & "');</script>"
            sScript1 &= "<script>location.href='../../main2.aspx';</script>"
            Call TIMS.Utl_RespWriteEnd(Me, objconn, "")
            Exit Sub
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            LabTMID.Text = "訓練業別"
        End If
        hTPlanID54.Value = ""
        '
        Datagrid4headTable.Visible = False
        Datagrid4Table.Visible = False
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '充電起飛計畫
            hTPlanID54.Value = "1"
            Datagrid4headTable.Visible = True
            Datagrid4Table.Visible = True
        End If

        '增加 申請階段 說明
        Select Case sm.UserInfo.TPlanID
            Case "28"
                Me.labAppStageMsg.Text = ""
                Me.labAppStageMsg.Text += "申請上半年課程：1<br>" & vbCrLf
                Me.labAppStageMsg.Text += "申請下半年課程：2<br>" & vbCrLf
            Case "54"
                Me.labAppStageMsg.Text = ""
                Me.labAppStageMsg.Text += "1月1~15日：1、1月16~31日：2    <br>" & vbCrLf
                Me.labAppStageMsg.Text += "2月1~15日：3、2月16~29日：4    <br>" & vbCrLf
                Me.labAppStageMsg.Text += "3月1~15日：5、3月16~31日：6    <br>" & vbCrLf
                Me.labAppStageMsg.Text += "4月1~15日：7、4月16~30日：8    <br>" & vbCrLf
                Me.labAppStageMsg.Text += "5月1~15日：9、5月16~31日：10   <br>" & vbCrLf
                Me.labAppStageMsg.Text += "6月1~15日：11、6月16~30日：12  <br>" & vbCrLf
                Me.labAppStageMsg.Text += "7月1~15日：13、7月16~31日：14  <br>" & vbCrLf
                Me.labAppStageMsg.Text += "8月1~15日：15、8月16~31日：16  <br>" & vbCrLf
                Me.labAppStageMsg.Text += "9月1~15日：17、9月16~30日：18  <br>" & vbCrLf
                Me.labAppStageMsg.Text += "10月1~15日：19、10月16~31日：20<br>" & vbCrLf
                Me.labAppStageMsg.Text += "11月1~15日：21、11月16~30日：22<br>" & vbCrLf
                Me.labAppStageMsg.Text += "12月1~15日：23、12月16~31日：24<br>" & vbCrLf

        End Select

        RoomName.ReadOnly = True
        FactModeOther.ReadOnly = True

        RoomName.Enabled = False
        RoomName.Visible = False
        FactMode.Enabled = False
        FactMode.Visible = False
        FactModeOther.Enabled = False
        FactModeOther.Visible = False
        FactModeTR.Visible = False
        RoomNameTD.Visible = False
        ContentTR.Visible = False
        'trainTR.Visible = False

        '20090319--加入判斷只能輸入整數的script，對應在新增的動作再檢查一次，以避免script運作無效。
        'PHour.Attributes.Add("onBlur", "if(this.value!=''){ if(!isInt(this.value)){alert('時數只能輸入整數。'); this.focus();} }")
        'Dim PHour_js_onBlur As String = "if(this.value !='') {var msg = ''; if(!isInt(this.value)){msg+='時數只能輸入整數。\n';} if(this.value <= 0){msg+='時數必須大於0\n';} if(msg !=''){alert(msg);this.focus();}}"
        'PHour.Attributes.Add("onBlur", PHour_js_onBlur)

        'Dim dt As DataTable
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))

        '放置畫面上的Dropdonwlist
        If Not IsPostBack Then
            'ClassAdd_TR.style("display") = "none"

            Call GetNewtable()  '建立空白表格

            Session(cst_TrainDescTable) = Nothing
            Session(cst_CostItemTable) = Nothing
            Session(Cst_MaterialTable) = Nothing
            Session(cst_Plan_OnClass) = Nothing
            Session(cst_Plan_BusPackage) = Nothing
            Session(Cst_PersonCostTable) = Nothing
            Session(Cst_CommonCostTable) = Nothing
            Session(Cst_SheetCostTable) = Nothing
            Session(Cst_OtherCostTable) = Nothing

            If Not Session("search") Is Nothing Then
                ViewState("search") = Session("search")
                Session("search") = Nothing
                'Session.Remove("search")
            End If

            '建立物件----Start
            Call CreateItem()
            '建立物件----End
            ViewState("GUID1") = TIMS.GetGUID() : Session("GUID1") = ViewState("GUID1")
            DataGrid2Table.Style("display") = "none"
            Page.RegisterStartupScript("window_onload", "<script language=""javascript"">Layer_change(1);</script>")
        End If

        If Not sUtl_PageLoad2() Then
            Exit Sub
        End If

        btnAdd.Enabled = True
        Button8.Enabled = True
        'If Not au.blnCanAdds Then
        '    Button8.Enabled = False
        '    TIMS.Tooltip(Button8, "您無權限使用該功能")
        '    btnAdd.Enabled = False
        '    TIMS.Tooltip(btnAdd, "您無權限使用該功能")
        'End If

        'State: View
        If TIMS.ClearSQM(Request("State")) <> "" Then
            btnAdd.Enabled = False
            Button8.Enabled = False
            TIMS.Tooltip(btnAdd, "狀態不可儲存。")
            TIMS.Tooltip(Button8, "狀態不可儲存。")
        End If
        '檢查功能權限----------------------------End

        'Request("PlanID") /TIMS.ClearSQM(Request("ComIDNO") /TIMS.ClearSQM(Request("SeqNO") 有空值或異常:true
        g_flagNG = Get_GflagNG1()

        If Not IsPostBack Then
            Label3.Text = sm.UserInfo.Years
            '(加強操作便利性)2005/4/1-Melody
            If TIMS.ClearSQM(Request("PlanID")) = "" Then              '新增狀態、帶入預設值
                '如果是自辦計劃，或者是委外並且是委訓登入，則帶入預設值
                If iPlanKind = 1 Or sm.UserInfo.LID = 2 Then
                    Dim sql As String = ""
                    sql = ""
                    sql &= " Select b.orgname, b.ComIDNO, c.ContactEmail, c.ZipCode, c.Address, b.OrgKind2"
                    sql &= " from Auth_Relship a"
                    sql &= " join Org_orginfo b on a.orgid=b.orgid"
                    sql &= " join Org_OrgPlanInfo c on a.RSID=c.RSID "
                    sql &= " where a.RID='" & sm.UserInfo.RID & "'"
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                    If dr IsNot Nothing Then
                        RIDValue.Value = sm.UserInfo.RID
                        ComidValue.Value = dr("ComIDNO")
                        center.Text = dr("orgname")
                        EMail.Text = TIMS.ChangeEmail(dr("ContactEmail").ToString)

                        EnterSupplyStyle.Enabled = False
                        Common.SetListItem(EnterSupplyStyle, "1")
                        Select Case Convert.ToString(dr("OrgKind2"))
                            Case "G" '非勞工團體
                            Case "W" '勞工團體
                                EnterSupplyStyle.Enabled = True
                                Common.SetListItem(EnterSupplyStyle, "2")
                        End Select

                    End If
                End If

                '建立訓練的DATATABLE
                Org.Disabled = False
                Button24.Visible = False

                '啟動年度 2013
                'If sm.UserInfo.Years >= 2013 Then
                'End If
                'Common.SetListItem(Solder, "00")
                'Solder.Enabled = False
                'TIMS.Tooltip(Solder, "「受訓資格」統一選擇為「不限」")
            Else
                Org.Disabled = True
                Button24.Visible = True
                '加入登入年度之RID by nick 

                '顯示該計畫資料，應該是修改或檢視
                Call Show_PlanPlanInfo()

                CreateClassTime()
                CreateCostItem()
                CreateMaterial()
                CreateTrainDesc()
                CreateBusPackage()
                Call CreatePersonCost()
                Call CreateCommonCost()
                Call CreateSheetCost()
                Call CreateOtherCost()
            End If
        End If

        '2004/12/7---------------------------------------前端增加javascript屬性---------Start
        'Me.rblAge.Attributes("onclick") = "set_Agelu();"
        date1.Attributes("onclick") = "javascript:show_calendar('STDate','','','CY/MM/DD');"
        date2.Attributes("onclick") = "javascript:show_calendar('FDDate','','','CY/MM/DD');"
        date3.Attributes("onclick") = "return chkTrainDate('STrainDate');"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?btnName=Button28');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        '增加快速點選機構清單
        If Org.Disabled = False Then
            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button28")
            If HistoryRID.Rows.Count <> 0 Then
                center.Attributes("onclick") = "showObj('HistoryList2');"
                center.Style("CURSOR") = "hand"
            End If
        End If

        'Modify by Kevin 2007/08/01 不適用於產業人才投資方案
        '//計算學科時數
        'GenSciHours.Attributes("onblur") = "set_SciHours();"
        'ProSciHours.Attributes("onblur") = "set_SciHours();"
        'ProTechHours.Attributes("onblur") = "set_SciHours();"
        'OtherHours.Attributes("onblur") = "set_SciHours();"

        '//班別資料訓練時數
        'THours.Attributes("onblur") = "set_THours();"

        '自辦申請計畫
        Button9.Attributes("onclick") = "return check_Cost2();"
        Button8.Attributes("onclick") = "return Check_Temp();" '草稿儲存

        Button1.Attributes("onclick") = "return CheckTrainDescTable();" '課程大綱檢查
        Button29.Attributes("onclick") = "return CheckAddTime();" ''上課時間檢查
        btnAddBusPackage.Attributes("onclick") = "return CheckAddBusPackage();" '包班事業單位資料檢查

        'Button1.Attributes("onclick") = "return CheckAddPIC();"
        'CostSort.Attributes("onclick") = " return CostModeChange()"
        'CostMode2.Attributes("onclick") = " return CostModeChange()"
        'CostMode3.Attributes("onclick") = " return CostModeChange()"
        'CostMode4.Attributes("onclick") = " return CostModeChange()"

        '計算經費來源的加總
        TNum.Attributes("onblur") = "CountCostSource();"
        box6.Attributes("onclick") += "javascript:CountCostSource();"
        box7.Attributes("onclick") += "javascript:CountCostSource();"

        DefGovCost.Attributes("onblur") = "CountCostSource();"
        DefUnitCost.Attributes("onblur") = "CountCostSource();"
        DefStdCost.Attributes("onblur") = "CountCostSource();"

        'btn_GCID.Attributes("onclick") = "GETvalue();Get_GovClass('GCIDName');"
        'btn_GCID.Attributes("onclick") = "GETvalue();"
        GCIDName.Attributes.Add("onDblClick", "javascript:Get_GovClass('GCIDName');")
        GCIDName.Style("CURSOR") = "hand"
        'CostID2.Attributes.Add("onchange", "if(this.value =='03'){if(document.getbyTNum.value !=''){Itemage.value =TNum.value;}}")
        CostID2.Attributes("onchange") = "GetItemage();"

        IsBusiness.Enabled = False
        EnterpriseName.Enabled = False
        IsBusiness.ToolTip = "本年度暫不開放此功能"
        EnterpriseName.ToolTip = "本年度暫不開放此功能"

        Dim str_script As String = ""
        str_script += "<script>" & vbCrLf
        str_script += "document.getElementById('Labsave').style.display=""none"";" & vbCrLf
        str_script += "CountCostSource();" & vbCrLf
        str_script += "showPTID('Classification1','PTID1','PTID2');" & vbCrLf
        str_script += "showCostType('" & RadioButtonList1.ClientID & "','" & DataGrid2Table.ClientID & "','" & Table6.ClientID & "');" & vbCrLf
        str_script += "</script>" & vbCrLf
        Page.RegisterStartupScript("CountCostSource", str_script)

        '確認機構是否為黑名單
        Dim vsMsg2 As String = ""
        vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            btnAdd.Visible = False
            Button8.Visible = False
            TIMS.Tooltip(btnAdd, vsMsg2)
            TIMS.Tooltip(Button8, vsMsg2)

            Page.RegisterStartupScript("", String.Concat("<script>alert('", vsMsg2, "');</script>"))
        End If

        PointType.Attributes("onclick") = "GetPointName();"

        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'PackageType.Items(0).Attributes.Add("disabled", "disabled")
            'TIMS.Tooltip(PackageType.Items(0), "充電起飛計畫不可選擇非包班!!")
            TIMS.Tooltip(PackageType, "充電起飛計畫不可選擇非包班!!")
            PackageType.Attributes("onclick") = "GetPackageName54();"
        Else
            PackageType.Attributes("onclick") = "GetPackageName();"
        End If
        'If Radiobuttonlist1.SelectedValue = "Y" Then '判斷學分數的*是否顯示
        '    S1.Visible = True
        'Else
        '    S1.Visible = False
        'End If

        '2011 功能按鈕權限控管--Start
        Dim strSechObjID As String = "" '查詢按鈕物件ID
        Dim strAddsObjID As String = "" '維護按鈕物件ID
        Dim strPrntObjID As String = "" '列印按鈕物件ID

        strAddsObjID = String.Concat(Button1.ClientID, ",", Button29.ClientID, ",", btnAddBusPackage.ClientID, ",", Button9.ClientID, ",", Button8.ClientID, ",", btnAdd.ClientID)
        Call TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
        '2011 功能按鈕權限控管--End
    End Sub

    'Request("PlanID") /TIMS.ClearSQM(Request("ComIDNO") /TIMS.ClearSQM(Request("SeqNO") 有空值或異常:true
    Function Get_GflagNG1() As Boolean
        Dim rst As Boolean = False
        If TIMS.ClearSQM(Request("PlanID")) = "" Then rst = True
        If TIMS.ClearSQM(Request("ComIDNO")) = "" Then rst = True 'Exit Sub
        If TIMS.ClearSQM(Request("SeqNO")) = "" Then rst = True 'Exit Sub
        If Val(Request("PlanID")) = 0 Then rst = True
        If Val(Request("ComIDNO")) = 0 Then rst = True 'Exit Sub
        If Val(Request("SeqNO")) = 0 Then rst = True 'Exit Sub
        Return rst
    End Function

    Function sUtl_PageLoad2() As Boolean
        Dim rst As Boolean = True
        Dim rqPlanID As String = "" '外部傳入 copy可能
        rqPlanID = TIMS.ClearSQM(Request("PlanID"))
        If rqPlanID = "" Then rqPlanID = sm.UserInfo.PlanID
        '判斷計畫種類，選擇要顯示的經費項目
        'rqPlanID = Convert.ToString(Request("PlanID"))
        Dim sql As String = ""
        sql = "SELECT TPLANID,PLANKIND,YEARS FROM ID_PLAN WHERE PlanID=@PlanID"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.VarChar).Value = rqPlanID
            dt.Load(.ExecuteReader())
        End With
        '顯示E-Mail欄位給予填寫
        Table1_Email.Visible = True
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, cst_errmsg1)
            rst = False
            'Exit Sub
        End If
        Dim dr As DataRow = dt.Rows(0)
        iPlanKind = dr("PlanKind")
        TPlanID = Convert.ToString(dr("TPlanID"))
        '顯示E-Mail欄位給予填寫
        Table1_Email.Visible = True
        If iPlanKind = 1 Then '自辦
            Table1_Email.Visible = False
        End If
        Return rst
    End Function

    '如果是複製狀態, 則RID還是為原登入計畫之RID by nick
    Function sUtl_GetRIDn() As String
        Dim rst As String = ""
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT RID as RIDn FROM Auth_Relship WHERE PlanID = '" & sm.UserInfo.PlanID & "' and DistID = '" & sm.UserInfo.DistID & "'"
        sql &= " and orgid in (select orgid from org_orginfo where ComIDNO ='" & TIMS.ClearSQM(Request("ComIDNO")) & "')"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If sm.UserInfo.LID = 0 Then
                'sm.UserInfo.TPlanID = "28" Gloria 同意，加入署(局)可檢視、修改班級查詢資料 by AMU 20070913
                Button8.Visible = False
                btnAdd.Visible = True
                Button24.ToolTip = "審核通過 or 審核後修正者,檢視班級"
                sql = ""
                sql &= "SELECT RID as RIDn FROM Auth_Relship WHERE RID = '" & sm.UserInfo.RID & "' and DistID = '" & sm.UserInfo.DistID & "'"
            End If
        End If
        Dim dr As DataRow
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            rst = Convert.ToString(dr("RIDn"))
        End If
        Return rst
    End Function

    '顯示該計畫資料
    Sub Show_PlanPlanInfo()
        If g_flagNG Then
            Common.MessageBox(Me, cst_errmsg3)
            Exit Sub
        End If

        Dim sRIDn As String
        sRIDn = sUtl_GetRIDn()

        Dim dr As DataRow
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.*" & vbCrLf
        'sql &= " ,s.CJOB_UNKEY" & vbCrLf
        'sql &= " ,s.CJOB_NO" & vbCrLf
        'sql &= " ,s.CJOB_Name" & vbCrLf
        sql &= " ,b.OrgName" & vbCrLf
        sql &= " ,c.RID RIDValue" & vbCrLf
        sql &= " ,b.OrgKind2" & vbCrLf
        sql &= " ,ISNULL(d.JobID,d.TrainID) JobID" & vbCrLf
        sql &= " ,ISNULL(d.JobName,d.TrainName) JobName" & vbCrLf
        sql &= " ,ISNULL(d.JobID,d.TrainID) TrainID" & vbCrLf
        sql &= " ,ISNULL(d.JobName,d.TrainName) TrainName" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN Org_OrgInfo b ON a.ComIDNO=b.ComIDNO" & vbCrLf
        sql &= " JOIN Auth_Relship c ON c.RID=a.RID And c.OrgID=b.OrgID And c.Planid=a.Planid " & vbCrLf
        sql &= " LEFT JOIN Key_TrainType d ON a.TMID=d.TMID" & vbCrLf
        sql &= " LEFT JOIN SHARE_CJOB s on s.CJOB_UNKEY = a.CJOB_UNKEY" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.PlanID='" & TIMS.ClearSQM(Request("PlanID")) & "' " & vbCrLf
        sql &= " and a.ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO")) & "' " & vbCrLf
        sql &= " and a.SeqNO='" & TIMS.ClearSQM(Request("SeqNO")) & "'" & vbCrLf
        Try
            dr = DbAccess.GetOneRow(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Page, cst_errmsg4)
            'Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = ""
            strErrmsg &= "/* sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf

            strErrmsg += TIMS.GetErrorMsg(Page) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            'Throw ex
            Exit Sub
        End Try
        If dr Is Nothing Then
            Common.MessageBox(Page, cst_errmsg5)
            Exit Sub
        End If
        'If dr Is Nothing Then Exit Sub
        Dim dtSCJOB As DataTable = TIMS.Get_SHARECJOBdt(Me, objconn)

        'If Not dr Is Nothing Then
        'End If

        'If dr("ProcID").ToString <> "" Then
        '    Common.SetListItem(ClassChar, dr("ProcID").ToString)
        'End If
        Call GetNewtable() '建立空白表格Taddress2下拉選單用

        If dr("PointYN").ToString <> "" Then '是否為學分班
            'Radiobuttonlist1.SelectedValue = dr("PointYN").ToString
            Common.SetListItem(RadioButtonList1, dr("PointYN").ToString)
            Select Case Convert.ToString(dr("PointYN"))
                Case cst_學分班
                    Labmsg3.Text = Cst_msgother3b
                Case cst_非學分班
                    Labmsg3.Text = Cst_msgother3
            End Select
        End If
        tNote2b.Text = Convert.ToString(dr("Note2")) 'cst_學分班
        tNote2.Text = Convert.ToString(dr("Note2")) 'cst_非學分班

        Dim sTMP_PlaceID As String = ""
        sTMP_PlaceID = TIMS.ClearSQM(Convert.ToString(dr("SciPlaceID")))
        If sTMP_PlaceID <> "" Then
            Common.SetListItem(SciPlaceID, sTMP_PlaceID)
            If SciPlaceID.SelectedIndex = 0 Then
                SciPlaceID = TIMS.Get_SciPlaceID(SciPlaceID, dr("ComIDNO"), 4, sTMP_PlaceID, objconn)
                Common.SetListItem(SciPlaceID, sTMP_PlaceID)
            End If
            Taddress2 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress2, sTMP_PlaceID, 1, 1, objconn)
        End If

        sTMP_PlaceID = TIMS.ClearSQM(Convert.ToString(dr("TechPlaceID")))
        If sTMP_PlaceID <> "" Then
            Common.SetListItem(TechPlaceID, sTMP_PlaceID)
            If TechPlaceID.SelectedIndex = 0 Then
                TechPlaceID = TIMS.Get_TechPlaceID(TechPlaceID, dr("ComIDNO"), 4, sTMP_PlaceID, objconn)
                Common.SetListItem(TechPlaceID, sTMP_PlaceID)
            End If
            Taddress3 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress3, sTMP_PlaceID, 2, 2, objconn)
        End If

        sTMP_PlaceID = TIMS.ClearSQM(Convert.ToString(dr("SciPlaceID2")))
        If sTMP_PlaceID <> "" Then
            Common.SetListItem(SciPlaceID2, sTMP_PlaceID)
            If SciPlaceID2.SelectedIndex = 0 Then
                SciPlaceID2 = TIMS.Get_SciPlaceID(SciPlaceID2, dr("ComIDNO"), 4, sTMP_PlaceID, objconn)
                Common.SetListItem(SciPlaceID2, sTMP_PlaceID)
            End If
            Taddress2 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress2, sTMP_PlaceID, 3, 1, objconn)
        End If

        sTMP_PlaceID = TIMS.ClearSQM(Convert.ToString(dr("TechPlaceID2")))
        If sTMP_PlaceID <> "" Then
            Common.SetListItem(TechPlaceID2, sTMP_PlaceID)
            If TechPlaceID2.SelectedIndex = 0 Then
                TechPlaceID2 = TIMS.Get_TechPlaceID(TechPlaceID2, dr("ComIDNO"), 4, sTMP_PlaceID, objconn)
                Common.SetListItem(TechPlaceID2, sTMP_PlaceID)
            End If
            Taddress3 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress3, sTMP_PlaceID, 4, 2, objconn)
        End If

        If Convert.ToString(dr("AddressSciPTID")) <> "" Then
            'Taddress2.SelectedValue = dr("AddressSciPTID")
            Common.SetListItem(Taddress2, dr("AddressSciPTID"))
        End If
        If Convert.ToString(dr("AddressTechPTID")) <> "" Then
            'Taddress3.SelectedValue = dr("AddressTechPTID")
            Common.SetListItem(Taddress3, dr("AddressTechPTID"))
        End If


        If dr("SciPlaceID").ToString <> "" _
            OrElse dr("TechPlaceID").ToString <> "" _
            OrElse dr("SciPlaceID2").ToString <> "" _
            OrElse dr("TechPlaceID2").ToString <> "" Then
            RoomName.Enabled = False
            FactMode.Enabled = False
            FactModeOther.Enabled = False
        Else
            RoomName.Enabled = False
            FactMode.Enabled = False
            FactModeOther.Enabled = False
        End If

        RIDValue.Value = dr("RID")
        ComidValue.Value = dr("ComIDNO")
        center.Text = dr("orgname")

        'PackageType欄位是2011/5/12才加進去的,如果是舊資料才需帶IsBusiness
        IsBusiness.Checked = False
        If Not IsDBNull(dr("IsBusiness")) Then
            IsBusiness.Checked = True
            If Convert.ToString(dr("IsBusiness")).ToString = "N" Then
                IsBusiness.Checked = False
            End If
            'If Convert.ToString(dr("IsBusiness")).ToString = "N" Then IsBusiness.Checked = False Else IsBusiness.Checked = True
        End If
        EnterpriseName.Text = dr("EnterpriseName").ToString
        FirstSort.Text = dr("FirstSort").ToString

        If dr("EnterSupplyStyle").ToString <> "" Then
            Common.SetListItem(EnterSupplyStyle, dr("EnterSupplyStyle").ToString)
            Select Case Convert.ToString(dr("OrgKind2"))
                Case "G" '非勞工團體
                    EnterSupplyStyle.Enabled = False
                Case "W" '勞工團體
                    EnterSupplyStyle.Enabled = True
            End Select
        End If

        'Select Case strYears
        '    Case "2014"
        '    Case "2015"
        'End Select

        '設定 GCID1Value.Value  '取得要比對的業別資料。
        Select Case strYears
            Case "2014"
                If dr("GCID").ToString <> "" Then
                    Me.GCIDValue.Value = dr("GCID").ToString
                    Me.GCIDName.Text = TIMS.Get_GCIDName(dr("GCID").ToString, strYears, objconn)
                    '------------------------------------
                    Dim sql99 As String = "SELECT GCODE1 FROM ID_GOVCLASSCAST WHERE GCID =" & Convert.ToString(dr("GCID"))
                    Dim dr99 As DataRow = DbAccess.GetOneRow(sql99, objconn)
                    If Not dr99 Is Nothing Then GCID1Value.Value = Convert.ToString(dr99("GCode1"))
                    '------------------------------------
                End If
            Case "2015"
                If dr("GCID2").ToString <> "" Then
                    Me.GCIDValue.Value = dr("GCID2").ToString
                    Me.GCIDName.Text = TIMS.Get_GCIDName(dr("GCID2").ToString, strYears, objconn)

                    Dim sql99 As String = "SELECT GCODE1 FROM V_GOVCLASSCAST2 WHERE GCID2 =" & Convert.ToString(dr("GCID2"))
                    Dim dr99 As DataRow = DbAccess.GetOneRow(sql99, objconn)
                    If Not dr99 Is Nothing Then GCID1Value.Value = Convert.ToString(dr99("GCODE1"))
                End If
        End Select

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            jobValue.Value = dr("TMID").ToString
            TB_career_id.Text = "[" & dr("jobID").ToString & "]" & dr("jobName").ToString
        Else
            trainValue.Value = dr("TMID").ToString
            TB_career_id.Text = "[" & dr("TrainID").ToString & "]" & dr("TrainName").ToString
        End If

        cjobValue.Value = dr("CJOB_UNKEY").ToString
        txtCJOB_NAME.Text = TIMS.Get_CJOBNAME(dtSCJOB, cjobValue.Value)

        'If dr("CJOB_UNKEY").ToString <> "" Then
        '    txtCJOB_NAME.Text = "[" & dr("CJOB_NO").ToString & "]" & dr("CJOB_NAME").ToString
        'End If

        If Request(cst_ccopy) = "1" Then
            Label3.Text = sm.UserInfo.Years
        Else
            Label3.Text = dr("PlanYear").ToString
        End If
        PlanCause.Text = dr("PlanCause").ToString
        PurScience.Text = dr("PurScience").ToString
        PurTech.Text = dr("PurTech").ToString
        PurMoral.Text = dr("PurMoral").ToString
        Common.SetListItem(Degree, dr("CapDegree").ToString)
        Common.SetListItem(AppStage, dr("AppStage").ToString)

        '不管什麼都是「年滿15歲以上」。
        'Const cst_ageoDef As Integer = 16 'other Years Start
        rdoAge1.Checked = True
        rdoAge2.Checked = False
        txtAge1.Text = "" 'cst_AgeOtherDef
        If Convert.ToString(dr("CapAge1")) <> "" Then
            If Val(dr("CapAge1")) >= cst_AgeOtherDef Then
                '若不是 年滿15歲以上 選擇顯示 目前所輸入的年齡。
                txtAge1.Text = Convert.ToString(dr("CapAge1"))
                rdoAge1.Checked = False
                rdoAge2.Checked = True
            End If
        End If

        'Common.SetListItem(Sex, dr("CapSex").ToString)
        'Common.SetListItem(Solder, dr("CapMilitary").ToString)
        ''啟動年度 2013
        'If sm.UserInfo.Years >= 2013 Then
        '    Common.SetListItem(Solder, "00")
        '    Solder.Enabled = False
        '    TIMS.Tooltip(Solder, "「受訓資格」統一選擇為「不限」")
        'End If
        If sm.UserInfo.Years >= "2015" OrElse Convert.ToString(Request(cst_ccopy)) = "1" Then
            '該欄位2015年後 暫不使用。
            Other1.Text = "" ' dr("CapOther1").ToString
            Other2.Text = "" ' dr("CapOther2").ToString
            Other3.Text = "" ' dr("CapOther3").ToString
        Else
            Other1.Text = dr("CapOther1").ToString
            Other2.Text = dr("CapOther2").ToString
            Other3.Text = dr("CapOther3").ToString
        End If
        If Other1.Text = "" Then Other1.Text = Cst_msgother1
        If Other2.Text = "" Then Other2.Text = Cst_msgother1
        If Other3.Text = "" Then Other3.Text = Cst_msgother1
        TIMS.Tooltip(Other1, Cst_msgother1)
        TIMS.Tooltip(Other2, Cst_msgother1)
        TIMS.Tooltip(Other3, Cst_msgother1)

        Other1.Enabled = False
        Other2.Enabled = False
        Other3.Enabled = False

        TMScience.Text = dr("TMScience").ToString
        GenSciHours.Text = dr("GenSciHours").ToString
        ProSciHours.Text = dr("ProSciHours").ToString
        SciHours.Text = Int(IIf(dr("GenSciHours").ToString = "", 0, dr("GenSciHours"))) + Int(IIf(dr("ProSciHours").ToString = "", 0, dr("ProSciHours").ToString))
        ProTechHours.Text = dr("ProTechHours").ToString
        OtherHours.Text = dr("OtherHours").ToString
        TotalHours.Text = dr("TotalHours").ToString

        EMail.Text = dr("PlanEMail").ToString
        CredPoint.Text = dr("CredPoint").ToString
        RoomName.Text = dr("RoomName").ToString
        Common.SetListItem(FactMode, dr("FactMode").ToString)
        FactModeOther.Text = dr("FactModeOther").ToString
        ConNum.Text = dr("ConNum").ToString
        ContactName.Text = dr("ContactName").ToString
        ContactPhone.Text = dr("ContactPhone").ToString
        ContactEmail.Text = dr("ContactEmail").ToString
        ContactFax.Text = dr("ContactFax").ToString
        If Convert.ToString(dr("ClassCate")) <> "" Then
            Common.SetListItem(ClassCate, dr("ClassCate").ToString)
        End If
        '授課時段'早上'下午'晚上
        'TPERIOD28_C1.Checked = False '早上
        'TPERIOD28_C2.Checked = False '下午
        'TPERIOD28_C3.Checked = False '晚上
        'If Len(Convert.ToString(dr("TPERIOD28"))) >= 3 Then
        '    If Convert.ToString(dr("TPERIOD28")).Substring(0, 1) = "Y" Then
        '        TPERIOD28_C1.Checked = True
        '    End If
        '    If Convert.ToString(dr("TPERIOD28")).Substring(1, 1) = "Y" Then
        '        TPERIOD28_C2.Checked = True
        '    End If
        '    If Convert.ToString(dr("TPERIOD28")).Substring(2, 1) = "Y" Then
        '        TPERIOD28_C3.Checked = True
        '    End If
        'End If
        Content.Text = dr("Content").ToString

        If Convert.ToString(dr("TotalCost")) <> "" Then
            TotalCost3.Text = CInt(dr("TotalCost")) '學分班(總價)
            'TotalCost2.Text = TotalCost3.Text '非學分班(總價)
        End If
        'Note.Text = dr("Note").ToString '這個是用產生的

        '複製狀態下,有些資料不複製--------------Start
        If Request(cst_ccopy) = "1" Then
            '如果是複製狀態, 則RID還是為原登入計畫之RID by nick
            RIDValue.Value = sRIDn 'dr("RIDn")
        Else
            RIDValue.Value = dr("RID")
            PointName.Text = ""
            If Convert.ToString(dr("PointType")) <> "" Then
                'PointType.SelectedValue = dr("PointType").ToString '學分種類
                Common.SetListItem(PointType, dr("PointType").ToString)
                If Not PointType.SelectedItem Is Nothing Then
                    PointName.Text = PointType.SelectedItem.Text  '學分種類名稱
                End If
            End If

            PackageName.Text = ""
            If Convert.ToString(dr("PackageType")) <> "" Then '包班種類名稱
                Common.SetListItem(PackageType, dr("PackageType").ToString)
                If Not PackageType.SelectedItem Is Nothing Then
                    Select Case PackageType.SelectedValue
                        Case "1" '非包班
                        Case Else
                            PackageName.Text = "(" & PackageType.SelectedItem.Text & ")" '包班種類名稱
                    End Select
                End If
            End If

            'ClassName.Text = dr("ClassName").ToString
            ClassName.Text = Trim(dr("ClassName").ToString)
            '取得班級名稱去掉學士學分班,碩士學分班,博士學分班
            If PointName.Text <> "" Then
                ClassName.Text = Replace(ClassName.Text, Trim(PointName.Text), "") '學分班種類
            End If
            '非包班、企業包班、聯合企業包班
            If PackageName.Text <> "" Then
                ClassName.Text = Replace(ClassName.Text, Trim(PackageName.Text), "") '企業包班種類
            End If

            'Select Case Right(Trim(dr("ClassName").ToString), 5)
            '    Case "學士學分班", "碩士學分班", "博士學分班"
            '        '取得班級名稱去掉學士學分班,碩士學分班,博士學分班
            '        ClassName.Text = Left(Trim(dr("ClassName").ToString), Len(Trim(dr("ClassName").ToString)) - 5)
            '    Case Else
            '        ClassName.Text = dr("ClassName").ToString
            'End Select

            Class_Unit.Value = dr("Class_Unit").ToString

            TNum.Text = Convert.ToString(dr("TNum"))
            'Itemage.Value = dr("TNum").ToString
            THours.Text = Convert.ToString(dr("THours"))
            STDate.Text = TIMS.Cdate3(Convert.ToString(dr("STDate")))
            FDDate.Text = TIMS.Cdate3(Convert.ToString(dr("FDDate")))
            CyclType.Text = TIMS.FmtCyclType(dr("CyclType"))
            ClassCount.Text = If(Convert.ToString(dr("ClassCount")) <> "", Convert.ToString(dr("ClassCount")), "1")

            DefGovCost.Text = dr("DefGovCost").ToString
            DefUnitCost.Text = dr("DefUnitCost").ToString
            DefStdCost.Text = dr("DefStdCost").ToString
            'If dr("DefGovCost").ToString <> "" And TNum.Text <> "" Then
            '    Total1.Text = CInt(dr("DefGovCost").ToString) / CInt(TNum.Text)
            'End If
            'If dr("DefUnitCost").ToString <> "" And TNum.Text <> "" Then
            '    Total2.Text = CInt(dr("DefUnitCost").ToString) / CInt(TNum.Text)
            'End If
            'If dr("DefStdCost").ToString <> "" And TNum.Text <> "" Then
            '    Total3.Text = CInt(dr("DefStdCost").ToString) / CInt(TNum.Text)
            'End If

            '已存為正式資料，而且不是要複製計畫，草稿儲存功能不啟用
            If dr("IsApprPaper") = "Y" Then
                Button8.Visible = False
                TIMS.Tooltip(Button8, "已存為正式資料，草稿儲存功能不啟用!")
            End If

            '2007以前的資料只可查詢，不可儲存 by AMU 2008-01-14
            If CStr(sm.UserInfo.Years) <= "2007" Then
                If dr("IsApprPaper") = "Y" Then '已存為正式資料
                    Button8.Visible = False '草稿儲存
                    'btnAdd.Visible = False '正式儲存
                    btnAdd.ToolTip += "本班已正式儲存，再次儲存請小心謹慎"
                Else
                    Button8.Visible = True '草稿儲存
                    btnAdd.Visible = True '正式儲存
                End If
                Button24.Visible = True '回上一頁
            End If

            '審核狀況。
            Select Case Convert.ToString(dr("AppliedResult"))
                Case "Y"               '2005/6/20--Melody審核通過or審核後修正者,不可修改班級名稱,期別,開結訓日,課程時數
                    'Case "Y", "O" 改為 Case "Y" 2007-03-02 kevin同意開放
                    If Convert.ToString(dr("IsApprPaper")) = "Y" Then
                        ClassName.ReadOnly = True
                        CyclType.ReadOnly = True
                        CustomValidator4.Enabled = False
                        STDate.ReadOnly = True
                        FDDate.ReadOnly = True
                        date1.Visible = False
                        date2.Visible = False
                        SciHours.ReadOnly = True
                        GenSciHours.ReadOnly = True
                        ProSciHours.ReadOnly = True
                        ProTechHours.ReadOnly = True
                        OtherHours.ReadOnly = True
                        TotalHours.ReadOnly = True

                        'THours.ReadOnly = True
                        'If dr("TransFlag").ToString = "Y" Then
                        '    CTName.ReadOnly = True
                        '    Button27.Disabled = True
                        '    TAddress.ReadOnly = True
                        '    CTName.Attributes.Remove("onblur") '= "getzipname(this.value,'CTName','TaddressZip');"
                        '    TIMS.Tooltip(CTName, "班級轉入後，不可修改上課地址")
                        '    TIMS.Tooltip(Button27, "班級轉入後，不可修改上課地址")
                        '    TIMS.Tooltip(TAddress, "班級轉入後，不可修改上課地址")

                        'Else
                        '    CTName.ReadOnly = False
                        '    Button27.Disabled = False
                        '    TAddress.ReadOnly = False
                        '    Button27.Attributes("onclick") = "getZip('../../js/Openwin/zipcode.aspx', 'CTName', 'TaddressZip');"
                        '    CTName.Attributes("onblur") = "getzipname(this.value,'CTName','TaddressZip');"
                        '    TIMS.Tooltip(CTName, "班級尚未轉入，可修改上課地址")
                        '    TIMS.Tooltip(Button27, "班級尚未轉入，可修改上課地址")
                        '    TIMS.Tooltip(TAddress, "班級尚未轉入，可修改上課地址")

                        'End If
                    End If

                    If $"{dr("AppliedResult")}" = "Y" Then
                        If iPlanKind = 1 And sm.UserInfo.LID = 2 And $"{dr("TransFlag")}" = "N" Then
                            Disabled_Items("委訓單位限制") '委訓單位限制
                        End If
                        If iPlanKind = 2 Then '計畫種類為委外者
                            If sm.UserInfo.LID = 2 Then Disabled_Items("計畫種類為委辦")
                        End If

                        '---Gloria 2007/8/30同意分署(中心)可在審核後，再次修改---
                        If sm.UserInfo.LID = 1 Then
                            ClassName.ReadOnly = False '班別名稱
                            STDate.ReadOnly = False '訓練起日
                            FDDate.ReadOnly = False '訓練迄日
                            CyclType.ReadOnly = False '期別
                            ClassCount.ReadOnly = False '班數
                        End If

                    End If
                Case "N"
                Case "M", ""
                    If dr("TransFlag").ToString = "Y" Then
                        center.Enabled = False
                        Org.Disabled = True
                    End If
            End Select
        End If
        '複製狀態下,有些資料不複製--------------End

        If TIMS.ClearSQM(Request("todo")) = 1 Then '按鈕狀態控制
            Disabled_Items("僅顯示")
        End If

    End Sub

    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            Me.isBlack.Value = "Y"
            Me.Blackorgname.Value = sm.UserInfo.OrgName
        End If
        Return rst
    End Function

    '建立下拉選單物件
    Sub CreateItem()
        'Dim StrCharID As String
        AppStage = TIMS.Get_AppStage(AppStage)

        'With Sex
        '    .Items.Add(New ListItem("==請選擇==", ""))
        '    .Items.Add(New ListItem("不分", "0"))
        '    .Items.Add(New ListItem("男", "M"))
        '    .Items.Add(New ListItem("女", "F"))
        'End With
        'Sex.SelectedIndex = 1
        'With Solder
        '    .Items.Add(New ListItem("==請選擇==", ""))
        '    .Items.Add(New ListItem("不限", "00"))
        '    .Items.Add(New ListItem("在役", "04"))
        '    .Items.Add(New ListItem("役畢(含免役)", "0103"))
        '    .Items.Add(New ListItem("未役", "02"))
        'End With
        'Solder.SelectedIndex = 1

        'l_Age.Text = "年滿15歲以上"
        'Age_l.Text = "15"
        'Age_u.Text = "65" '"60"

        Call TIMS.Get_ClassCatelog(ClassCate, objconn)
        Weeks = TIMS.Get_ddlWeeks(Weeks)

        'With Weeks
        '    .Items.Add(New ListItem("==請選擇==", ""))
        '    .Items.Add(New ListItem("星期一", "星期一"))
        '    .Items.Add(New ListItem("星期二", "星期二"))
        '    .Items.Add(New ListItem("星期三", "星期三"))
        '    .Items.Add(New ListItem("星期四", "星期四"))
        '    .Items.Add(New ListItem("星期五", "星期五"))
        '    .Items.Add(New ListItem("星期六", "星期六"))
        '    .Items.Add(New ListItem("星期日", "星期日"))
        'End With

        '設定時間物件值（DropDownList）
        Call CreateTimesItem(ddlpnH1, ddlpnH2, ddlpnM1, ddlpnM2)

        'With depID
        '    .Items.Add(New ListItem("==請選擇==", ""))
        '    .Items.Add(New ListItem("學科教室1", "101"))
        '    .Items.Add(New ListItem("學科教室2", "102"))
        '    .Items.Add(New ListItem("術科教室1", "201"))
        '    .Items.Add(New ListItem("術科教室2", "202"))
        'End With

        Degree = TIMS.Get_Degree(Degree, 2, objconn)

        'Session("CHARID") = TIMS.Get_CHARID(sm.UserInfo.OrgID, sm.UserInfo.Years)
        '用Session 會導致系統誤判，未建訓練性質單位無法運作其他程式 Modify by Kevin 20061219
        'StrCharID = TIMS.Get_CHARID(sm.UserInfo.OrgID, sm.UserInfo.Years)

        '將ComidValue.Value 塞入有效值
        If ComidValue.Value = "" Then
            ComidValue.Value = TIMS.sUtl_GetRqValue(Me, "ComIDNO", ComidValue.Value)
            If ComidValue.Value = "" Then
                ComidValue.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
            End If
        End If

        SciPlaceID = TIMS.Get_SciPlaceID(SciPlaceID, ComidValue.Value, 1, "", objconn)
        TechPlaceID = TIMS.Get_TechPlaceID(TechPlaceID, ComidValue.Value, 1, "", objconn)
        SciPlaceID2 = TIMS.Get_SciPlaceID(SciPlaceID2, ComidValue.Value, 1, "", objconn)
        TechPlaceID2 = TIMS.Get_TechPlaceID(TechPlaceID2, ComidValue.Value, 1, "", objconn)
        TIMS.Tooltip(SciPlaceID, "學科場地以登入者的機構為準")
        TIMS.Tooltip(TechPlaceID, "術科場地以登入者的機構為準")
        TIMS.Tooltip(SciPlaceID2, "學科場地以登入者的機構為準")
        TIMS.Tooltip(TechPlaceID2, "術科場地以登入者的機構為準")

        'Common.SetListItem(Classification1, "1")
        'If ComidValue.Value <> "" Then
        '    PTID1 = TIMS.Get_SciPTID(PTID1, ComidValue.Value, 2)
        '    PTID2 = TIMS.Get_TechPTID(PTID2, ComidValue.Value, 2)
        'End If
        'If ComidValue.Value <> "" Then
        '    PTID1 = TIMS.Get_SciPTID(PTID1, ComidValue.Value)
        '    PTID2 = TIMS.Get_TechPTID(PTID2, ComidValue.Value)
        'Else
        '    If TIMS.ClearSQM(Request("ComIDNO") Is Nothing Then
        '        PTID1 = TIMS.Get_SciPTID(PTID1, TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID.ToString))
        '        PTID2 = TIMS.Get_TechPTID(PTID2, TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID.ToString))
        '    Else
        '        PTID1 = TIMS.Get_SciPTID(PTID1, TIMS.ClearSQM(Request("ComIDNO"))
        '        PTID2 = TIMS.Get_TechPTID(PTID2, TIMS.ClearSQM(Request("ComIDNO"))
        '    End If
        'End If
        'TIMS.Tooltip(PTID1, "上課地點以登入者的機構為準")
        'TIMS.Tooltip(PTID2, "上課地點以登入者的機構為準")

        Dim exErrmsg As String = ""
        Try
            exErrmsg = ""
            If RIDValue.Value <> "" Then
                exErrmsg &= "RIDValue.Value : " & RIDValue.Value & vbCrLf
                TIMS.CreateTeacherScript(Me, RIDValue.Value, objconn)
            Else
                If sm.IsLogin Then
                    exErrmsg &= "sm.(""RID"") : " & Convert.ToString(sm.UserInfo.RID) & vbCrLf
                    TIMS.CreateTeacherScript(Me, sm.UserInfo.RID, objconn)
                End If
            End If
        Catch ex As Exception
            Me.upt_PlanX.Value = ""
            exErrmsg &= ex.ToString
            Throw New Exception(exErrmsg)
        End Try

        'Me.Classification1.Attributes.Add("onBlur", "javascript:showPTID('Classification1','PTID1','PTID2');Layer_change(5);")
        Me.Classification1.Attributes.Add("onchange", "javascript:showPTID('Classification1','PTID1','PTID2');Layer_change(5);")
        Me.RadioButtonList1.Attributes.Add("onclick", "javascript:Layer_change(5);showCostType('" & RadioButtonList1.ClientID & "','" & DataGrid2Table.ClientID & "','" & Table6.ClientID & "');")
        Me.TotalCost3.Attributes.Add("onchange", "javascript:onchg_total3();Layer_change(6);")
        'Me.Radiobuttonlist1.Attributes.Add("onMouseup", "javascript:showCostType('" & Radiobuttonlist1.ClientID & "','" & Li_CostType.ClientID & "','" & DataGrid2Table.ClientID & "','" & Table6.ClientID & "');Layer_change(3);")

        '任課教師
        Me.OLessonTeah1.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addx','OLessonTeah1','OLessonTeah1Value');") 'SD/04/LessonTeah1.aspx
        Me.OLessonTeah1.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','OLessonTeah1');"
        OLessonTeah1.Style.Item("CURSOR") = "hand"
        '助教
        Me.OLessonTeah2.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addy','OLessonTeah2','OLessonTeah2Value');") 'SD/04/LessonTeah1.aspx
        Me.OLessonTeah2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2Value','OLessonTeah2');"
        OLessonTeah2.Style.Item("CURSOR") = "hand"

        ''<Asp@button id="Upload" OnClick="UploadFile" Text="上傳圖片" runat="server"/> 
        Dim selsql As String = ""
        selsql = "SELECT COSTID+','+ITEMCOSTNAME COSTID,COSTNAME FROM KEY_COSTITEM2 ORDER BY SORT"
        CostID2 = TIMS.Get_KeyControl2(CostID2, selsql, "COSTNAME", "COSTID", objconn)
        CostID2.Attributes("onchange") = "ShowItemCostName('CostID2','ItemCostName','Itemage')"
        TIMS.Tooltip(Itemage, "計價數量，若無請填寫1")

        btnAdd.Attributes("onclick") = "javascript:notshow_button(); init(); window.setTimeout('show_secs()',1);"
        Button8.Attributes("onclick") = "javascript:notshow_button(); init(); window.setTimeout('show_secs()',1);"

    End Sub

    '設定時間物件值（DropDownList）
    Sub CreateTimesItem(ByRef oddlTimesH1 As DropDownList,
                        ByRef oddlTimesH2 As DropDownList,
                        ByRef oddlTimesM1 As DropDownList,
                        ByRef oddlTimesM2 As DropDownList)

        oddlTimesH1.Items.Clear()
        oddlTimesH2.Items.Clear()
        oddlTimesM1.Items.Clear()
        oddlTimesM2.Items.Clear()
        For intTimeHM As Integer = 0 To 22
            If intTimeHM >= 0 AndAlso intTimeHM <= 5 Then
                If intTimeHM = 0 Then
                    oddlTimesM1.Items.Add(New ListItem("00", "00"))
                    oddlTimesM2.Items.Add(New ListItem("00", "00"))
                Else
                    oddlTimesM1.Items.Add(New ListItem(CStr(intTimeHM * 10), CStr(intTimeHM * 10)))
                    oddlTimesM2.Items.Add(New ListItem(CStr(intTimeHM * 10), CStr(intTimeHM * 10)))
                End If
            End If
            If intTimeHM >= 8 AndAlso intTimeHM <= 22 Then
                If CStr(intTimeHM).Length < 2 Then
                    oddlTimesH1.Items.Add(New ListItem("0" & CStr(intTimeHM), "0" & CStr(intTimeHM)))
                    oddlTimesH2.Items.Add(New ListItem("0" & CStr(intTimeHM), "0" & CStr(intTimeHM)))
                Else
                    oddlTimesH1.Items.Add(New ListItem("" & CStr(intTimeHM), "" & CStr(intTimeHM)))
                    oddlTimesH2.Items.Add(New ListItem("" & CStr(intTimeHM), "" & CStr(intTimeHM)))
                End If
            End If
        Next
    End Sub

    '建立計畫 企業包班事業單位
    Sub CreateBusPackage()
        Me.Datagrid4headTable.Visible = False '若選擇非包班，則不能再選企業包班喔
        'Me.Datagrid4Table.Visible = False

        If Not TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '非充電計畫者 不可做企業包班新增
            Session(cst_Plan_BusPackage) = Nothing
            Exit Sub
        End If

        If hTPlanID54.Value = "" Then '非  '充電起飛計畫  (hTPlanID54.Value = "1")
            If Not Session(cst_Plan_BusPackage) Is Nothing Then Session(cst_Plan_BusPackage) = Nothing
            Exit Sub
        End If
        '充電起飛計畫 '非 聯合企業包班
        Select Case PackageType.SelectedValue
            Case "3"  '充電起飛計畫' 聯合企業包班
                Me.Datagrid4headTable.Visible = True
                'Me.btnAddBusPackage.Visible = True
                Me.btnAddBusPackage.Style.Item("display") = ""
                'Me.Datagrid4Table.Visible = True
            Case "2"  '充電起飛計畫' 企業包班
                Me.Datagrid4headTable.Visible = True
                'Me.btnAddBusPackage.Visible = False
                Me.btnAddBusPackage.Style.Item("display") = "none"
                'Me.Datagrid4Table.Visible = False
            Case Else
                If Not Session(cst_Plan_BusPackage) Is Nothing Then Session(cst_Plan_BusPackage) = Nothing
                Exit Sub
        End Select

        Const Cst_PKName As String = "BPID"
        ' Session(cst_Plan_BusPackage)
        Dim sql As String = ""
        Dim dt As DataTable
        Dim dr As DataRow

        If Session(cst_Plan_BusPackage) Is Nothing Then
            Dim dt1 As DataTable = Nothing
            If Me.upt_PlanX.Value <> "" Then
                sql = "SELECT * FROM Plan_BusPackage WHERE 1<>1"
                dt1 = DbAccess.GetDataTable(sql, objconn)
            Else
                If Request(cst_ccopy) = "1" Then
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = "SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                End If
            End If
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = "SELECT * FROM Plan_BusPackage WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                If Request(cst_ccopy) = "1" Then
                    sql = "SELECT * FROM Plan_BusPackage WHERE 1<>1"
                Else
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = "SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    If g_flagNG Then sql = "SELECT * FROM PLAN_BUSPACKAGE WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
            dt.Columns(Cst_PKName).AutoIncrement = True
            dt.Columns(Cst_PKName).AutoIncrementSeed = -1
            dt.Columns(Cst_PKName).AutoIncrementStep = -1

            If Not dt1 Is Nothing Then
                For Each dr1 As DataRow In dt1.Rows
                    If Not dr1.RowState = DataRowState.Deleted Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        For i As Integer = 0 To dr1.ItemArray.Length - 1
                            If dr.Table.Columns(i).ColumnName <> Cst_PKName Then
                                dr(dr.Table.Columns(i).ColumnName) = dr1(dr.Table.Columns(i).ColumnName)
                            End If
                        Next
                    End If
                Next
            End If

        Else
            dt = Session(cst_Plan_BusPackage)
        End If
        Session(cst_Plan_BusPackage) = dt

        Datagrid4Table.Visible = False
        If dt.Rows.Count > 0 Then
            txtUname.Text = ""
            txtIntaxno.Text = ""
            txtUbno.Text = ""
            Select Case PackageType.SelectedValue
                Case "2"  '充電起飛計畫' 企業包班
                    dr = dt.Rows(0)
                    txtUname.Text = Convert.ToString(dr("UName"))
                    txtIntaxno.Text = Convert.ToString(dr("Intaxno"))
                    txtUbno.Text = Convert.ToString(dr("Ubno"))
            End Select
            Datagrid4Table.Visible = True

            Datagrid4.DataSource = dt
            Datagrid4.DataBind()
        End If


    End Sub

    '建立上課時間
    Sub CreateClassTime()
        Dim sql As String
        Dim dt As DataTable
        Dim dt1 As DataTable = Nothing

        If Session(cst_Plan_OnClass) Is Nothing Then
            If Me.upt_PlanX.Value <> "" Then
                sql = "SELECT * FROM PLAN_ONCLASS WHERE 1<>1"
                dt1 = DbAccess.GetDataTable(sql, objconn)
            Else
                If Request(cst_ccopy) = "1" Then
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = "SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                End If
            End If
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = "SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                If Request(cst_ccopy) = "1" Then
                    sql = "SELECT * FROM PLAN_ONCLASS WHERE 1<>1"
                Else
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = "SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    If g_flagNG Then sql = "SELECT * FROM PLAN_ONCLASS WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
            dt.Columns("POCID").AutoIncrement = True
            dt.Columns("POCID").AutoIncrementSeed = -1
            dt.Columns("POCID").AutoIncrementStep = -1
            If Not dt1 Is Nothing Then
                For Each dr1 As DataRow In dt1.Rows
                    If Not dr1.RowState = DataRowState.Deleted Then
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        For i As Integer = 0 To dr1.ItemArray.Length - 1
                            If dr.Table.Columns(i).ColumnName <> "POCID" Then
                                dr(dr.Table.Columns(i).ColumnName) = dr1(dr.Table.Columns(i).ColumnName)
                            End If
                        Next
                    End If
                Next
            End If
        Else
            dt = Session(cst_Plan_OnClass)
            dt.Columns("POCID").AutoIncrement = True
            dt.Columns("POCID").AutoIncrementSeed = -1
            dt.Columns("POCID").AutoIncrementStep = -1
        End If
        Session(cst_Plan_OnClass) = dt

        DataGrid1Table.Visible = False
        If dt.Rows.Count > 0 Then
            DataGrid1Table.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If

    End Sub

    '計畫訓練內容簡介
    Sub CreateTrainDesc()
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        If Session(cst_TrainDescTable) Is Nothing Then
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")

                sql = "SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                If Request(cst_ccopy) = "1" Then
                    sql = "SELECT * FROM PLAN_TRAINDESC WHERE 1<>1 " & vbCrLf
                Else
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = "SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    If g_flagNG Then sql = "SELECT * FROM PLAN_TRAINDESC WHERE 1<>1"
                    ''20090518 andy edit PTDID加入排序  
                    'sql &= " order by STrainDate,PName, PTDID "
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(cst_TrainDescTable)
        End If

        '20090401 andy edit 加入排序
        '-------------------
        dt.DefaultView.Sort = "STrainDate asc,PName asc,PTDID asc"
        dt = dt.DefaultView.Table
        '-------------------
        dt.Columns("PTDID").AutoIncrement = True
        dt.Columns("PTDID").AutoIncrementSeed = -1
        dt.Columns("PTDID").AutoIncrementStep = -1
        Session(cst_TrainDescTable) = dt

        Datagrid3Table.Visible = False
        Datagrid3Table.Style.Item("display") = "none"
        If dt.Rows.Count > 0 Then
            Datagrid3Table.Visible = True
            Datagrid3Table.Style.Item("display") = ""

            With Datagrid3
                .DataSource = dt
                .DataKeyField = "PTDID"
                .DataBind()
            End With
        End If
    End Sub

    '計畫經費項目檔
    Sub CreateCostItem()
        Dim dt As DataTable
        Dim Total As Double

        If Session(cst_CostItemTable) Is Nothing Then
            Dim sql As String = ""
            sql = " SELECT * FROM PLAN_COSTITEM WHERE 1=1" & vbCrLf
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql &= " AND CostMode='5'" & vbCrLf ' 為產學訓專用
                sql &= " AND PlanID='" & PlanID_value & "'" & vbCrLf
                sql &= " AND ComIDNO='" & ComIDNO_value & "'" & vbCrLf
                sql &= " AND SeqNO='" & SeqNO_value & "'" & vbCrLf
            Else
                If Request(cst_ccopy) = "1" Then
                    sql &= " AND 1<>1" & vbCrLf
                Else
                    sql &= " AND CostMode='5'" & vbCrLf ' 為產學訓專用
                    If Not g_flagNG Then
                        PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                        ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                        SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                        sql &= " and PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'" & vbCrLf
                    End If
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(cst_CostItemTable)
        End If
        dt.Columns("PCID").AutoIncrement = True
        dt.Columns("PCID").AutoIncrementSeed = -1
        dt.Columns("PCID").AutoIncrementStep = -1
        Session(cst_CostItemTable) = dt

        Total = 0
        dbld2TempTotal = 0
        TotalCost2.Text = 0 '非學分班(總價)
        'TotalCost3.Text = 0
        DefGovCost.Text = 0
        DefStdCost.Text = 0
        'Total = dbld2TempTotal '小計
        DataGrid2Table.Style.Item("display") = "none"

        If dt.Rows.Count > 0 Then
            '非學分班
            DataGrid2Table.Style.Item("display") = ""
            With DataGrid2
                .DataSource = dt
                .DataKeyField = "PCID"
                .DataBind()
            End With
            Total = dbld2TempTotal '小計
            'For Each item As DataGridItem In DataGrid2.Items
            '    Dim sItemCost As Label = item.FindControl("DataGrid2Label3") '小計
            '    'Total += CDbl(item.Cells(cst_小計).Text)
            '    Total += CDbl(sItemCost.Text)
            'Next
            TotalCost2.Text = TIMS.ROUND(Total) '非學分班(總價)
            'TotalCost3.Text = TIMS.Round(Total) '學分班(總價)
            DefGovCost.Text = TIMS.ROUND(Total * 0.8)
            DefStdCost.Text = TIMS.ROUND(Total * 0.2)
            If Me.hTPlanID54.Value = "1" Then
                DefGovCost.Text = TIMS.ROUND(Total)
                DefStdCost.Text = TIMS.ROUND(0)
            End If
        End If

        Call ChangNoteText(tmpNoteDt)
    End Sub

    '計畫材料品名項目檔
    Sub CreateMaterial()
        Dim dt As DataTable
        'Dim Total As Double
        Dim sql As String = ""
        If Session(Cst_MaterialTable) Is Nothing Then
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")

                sql = "SELECT * FROM Plan_Material WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                '材料品名項目
                If Request(cst_ccopy) = "1" Then
                    'COPY
                    sql = "SELECT * FROM Plan_Material where 1<>1"
                Else
                    '新增、修改
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = " SELECT * FROM PLAN_MATERIAL WHERE 1=1 and PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    If g_flagNG Then sql = "SELECT * FROM PLAN_MATERIAL WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(Cst_MaterialTable)
        End If
        dt.Columns(Cst_PMID).AutoIncrement = True
        dt.Columns(Cst_PMID).AutoIncrementSeed = -1
        dt.Columns(Cst_PMID).AutoIncrementStep = -1
        Session(Cst_MaterialTable) = dt

        If dt.Rows.Count = 0 Then
            DataGrid5.DataSource = dt
            DataGrid5.DataBind()
            DataGrid5.Style.Item("display") = "none"
        Else
            DataGrid5.Style.Item("display") = ""
            With DataGrid5
                .DataSource = dt
                .DataKeyField = Cst_PMID
                .DataBind()
            End With
        End If
    End Sub

    'Plan_PersonCost–一人份材料明細
    Function CreatePersonCost() As DataTable
        Dim dt As DataTable
        Dim DGobj As DataGrid = Me.DataGrid6
        Const cst_sSupFd As String = ",0 Total,0 subtotal" '補充欄位
        Dim sql As String = ""
        If Session(Cst_PersonCostTable) Is Nothing Then
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                'Copy機制
                sql = "SELECT Plan_PersonCost.*" & cst_sSupFd & " FROM Plan_PersonCost where PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                If Request(cst_ccopy) = "1" Then
                    'Copy機制
                    sql = "SELECT Plan_PersonCost.*" & cst_sSupFd & " FROM Plan_PersonCost where 1<>1"
                Else
                    '修改資料取得
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = ""
                    sql &= " SELECT Plan_PersonCost.*" & cst_sSupFd & " FROM Plan_PersonCost WHERE 1=1" & vbCrLf
                    sql &= " and PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'" & vbCrLf
                    sql &= " ORDER BY ItemNo" & vbCrLf
                    If g_flagNG Then sql = " SELECT Plan_PersonCost.*" & cst_sSupFd & " FROM Plan_PersonCost WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            '有資料
            dt = Session(Cst_PersonCostTable)
        End If
        dt.Columns(Cst_PersonCostpkName).AutoIncrement = True
        dt.Columns(Cst_PersonCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_PersonCostpkName).AutoIncrementStep = -1
        Session(Cst_PersonCostTable) = dt
        With DGobj
            .Style.Item("display") = "none"
            If dt.Rows.Count > 0 Then
                .Style.Item("display") = ""

                .DataSource = dt
                .DataKeyField = Cst_PersonCostpkName
                .DataBind()
            End If
        End With
        Dim subtotal As Integer = 0
        subtotal = 0
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If Not dr.RowState = DataRowState.Deleted Then
                    Dim iPerCount As Integer = Val(dr("PerCount"))
                    Dim iPrice As Integer = Val(dr("Price"))
                    Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                    dr("Total") = (iPerCount * iTNum) '顯示重算
                    dr("subtotal") = (iPrice * iPerCount * iTNum) '顯示重算 '小計
                    subtotal += Val(dr("subtotal"))
                End If
            Next
        End If
        Me.labTotal6.Text = subtotal
        'Me.labTotal67.Text = Val(Me.labTotal6.Text) + Val(Me.labTotal7.Text)
        trlabTotal6.Visible = False
        If subtotal > 0 Then trlabTotal6.Visible = True

        Call ChglabTotal67()
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function

    'Plan_CommonCost–共同材料明細
    Function CreateCommonCost() As DataTable
        Dim dt As DataTable
        Dim DGobj As DataGrid = Me.DataGrid7
        Const cst_sSupFd As String = ",0 subtotal, 0 eachCost" '補充欄位
        Dim sql As String = ""
        If Session(Cst_CommonCostTable) Is Nothing Then
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")

                sql = "SELECT Plan_CommonCost.*" & cst_sSupFd & " FROM Plan_CommonCost where PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                If Request(cst_ccopy) = "1" Then
                    'Copy機制
                    sql = "SELECT Plan_CommonCost.*" & cst_sSupFd & " FROM Plan_CommonCost where 1<>1"
                Else
                    '修改資料取得
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = ""
                    sql &= " SELECT Plan_CommonCost.*" & cst_sSupFd & " FROM Plan_CommonCost WHERE 1=1" & vbCrLf
                    sql &= " and PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'" & vbCrLf
                    sql &= " ORDER BY ItemNo" & vbCrLf
                    If g_flagNG Then sql = " SELECT Plan_CommonCost.*" & cst_sSupFd & " FROM Plan_CommonCost WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            '有資料
            dt = Session(Cst_CommonCostTable)
        End If
        dt.Columns(Cst_CommonCostpkName).AutoIncrement = True
        dt.Columns(Cst_CommonCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_CommonCostpkName).AutoIncrementStep = -1
        Session(Cst_CommonCostTable) = dt
        With DGobj
            .Style.Item("display") = "none"
            If dt.Rows.Count > 0 Then
                .Style.Item("display") = ""

                .DataSource = dt
                .DataKeyField = Cst_CommonCostpkName
                .DataBind()
            End If
        End With
        Dim subtotal As Integer = 0
        subtotal = 0
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If Not dr.RowState = DataRowState.Deleted Then
                    Dim iPrice As Integer = Val(dr("Price"))
                    Dim iAllCount As Integer = Val(dr("AllCount"))
                    Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                    dr("subtotal") = (iPrice * iAllCount)  '小計
                    dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
                    subtotal += Val(dr("subtotal"))
                End If
            Next
            'For Each dr As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
            '    Dim iPrice As Integer = Val(dr("Price"))
            '    Dim iAllCount As Integer = Val(dr("AllCount"))
            '    Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
            '    dr("subtotal") = (iPrice * iAllCount)  '小計
            '    dr("eachCost") = TIMS.Round(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
            '    subtotal += Val(dr("subtotal"))
            'Next
        End If
        Me.labTotal7.Text = subtotal
        'Me.labTotal67.Text = Val(Me.labTotal6.Text) + Val(Me.labTotal7.Text)
        trlabTotal7.Visible = False
        If subtotal > 0 Then trlabTotal7.Visible = True

        Call ChglabTotal67()
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function

    Function GetMaxSeqNum(ByVal Trans As SqlTransaction, ByVal conn As SqlConnection) As Integer
        Dim Rst As Integer = 1 '由1開始
        Dim exErrmsg As String = ""
        Try
            '取得SeqNO
            Dim sql As String = ""
            sql = " SELECT Max(SeqNO) MaxSeqNO From Plan_PlanInfo where ComIDNO='" & ComidValue.Value & "' and PlanID='" & sm.UserInfo.PlanID & "'"
            Dim dr As DataRow = DbAccess.GetOneRow(sql, Trans)
            If Not dr Is Nothing Then '有取得
                If Not IsDBNull(dr("MaxSeqNO")) Then '有值
                    Rst = CInt(dr("MaxSeqNO")) + 1 '最大值再加1
                End If
            End If

        Catch ex As Exception
            Me.upt_PlanX.Value = ""
            If Not Trans Is Nothing Then
                DbAccess.RollbackTrans(Trans)
            End If
            Call TIMS.CloseDbConn(conn)
            'If Not Trans Is Nothing Then DbAccess.RollbackTrans(Trans)
            exErrmsg &= ex.ToString & vbCrLf
            Throw New Exception(exErrmsg)
        End Try
        Return Rst
    End Function

    ''更新　Plan_VerReport(訓練計劃開班總表(產學訓))
    Sub Update_Plan_VerReport(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, ByRef TransConn As SqlConnection, ByRef Trans As SqlTransaction)
        If (PlanID <> "" And ComIDNO <> "" And SeqNo <> "") Then
            'TIMS.OpenDbConn(conn) 'Trans = DbAccess.BeginTrans(conn)
            Try
                Dim sql As String = "select * from Plan_VerReport where PlanID='" & PlanID & "' and ComIDNO='" & ComIDNO & "' and SeqNo='" & SeqNo & "'"
                'sql = "Update Plan_VerReport set Content = '" & Content.Text & "' where ComIDNO='" & ComidValue.Value & "' and PlanID='" & sm.UserInfo.PlanID & "' and SeqNo='" & TIMS.ClearSQM(Request("SeqNo") & "'"
                Dim da As SqlDataAdapter = Nothing
                Dim dt As DataTable = DbAccess.GetDataTable(sql, da, Trans)
                If dt.Rows.Count > 0 Then
                    Dim dr As DataRow = Nothing
                    dr = dt.Rows(0)
                    dr("Content") = Content.Text
                End If
                DbAccess.UpdateDataTable(dt, da, Trans)
                DbAccess.CommitTrans(Trans)
                'Call TIMS.CloseDbConn(conn)
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg6)
                'Common.MessageBox(Me, ex.ToString)
                Throw ex
            End Try
        End If
    End Sub

    ''更新　Class_ClassInfo
    Sub Update_Class_ClassInfo(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, ByRef TransConn As SqlConnection, ByRef Trans As SqlTransaction)
        If (PlanID <> "" AndAlso ComIDNO <> "" AndAlso SeqNo <> "") Then
            'Call TIMS.OpenDbConn(conn) 'Trans = DbAccess.BeginTrans(conn)
            Try
                Dim sql As String = " SELECT * FROM CLASS_CLASSINFO WHERE PlanID=" & PlanID & " AND ComIDNO='" & ComIDNO & "' AND SeqNo=" & SeqNo & ""
                Dim da As SqlDataAdapter = Nothing
                Dim dt As DataTable = DbAccess.GetDataTable(sql, da, Trans)
                If dt.Rows.Count > 0 Then
                    Dim dr As DataRow = Nothing
                    dr = dt.Rows(0)
                    If PointName.Text <> "" Then PointName.Text = TIMS.ClearSQM(PointName.Text)
                    If PackageName.Text <> "" Then PackageName.Text = TIMS.ClearSQM(PackageName.Text)
                    ClassName.Text = Replace(ClassName.Text, "&", "＆")
                    ClassName.Text = TIMS.ClearSQM(ClassName.Text)
                    ClassName.Text = Replace(ClassName.Text, PointName.Text, "") '學分班種類
                    ClassName.Text = Replace(ClassName.Text, PackageName.Text, "") '企業包班種類
                    Select Case RadioButtonList1.SelectedValue
                        Case cst_學分班 ' "Y"
                            dr("ClassCName") = ClassName.Text & PointName.Text & PackageName.Text
                        Case Else 'cst_非學分班
                            dr("ClassCName") = ClassName.Text & PackageName.Text
                    End Select
                    dr("TNum") = If(TNum.Text <> "", Val(TNum.Text), Convert.DBNull)
                    dr("THours") = If(THours.Text <> "", Val(THours.Text), Convert.DBNull)
                    dr("STDate") = TIMS.Cdate2(STDate.Text)
                    dr("FTDate") = TIMS.Cdate2(FDDate.Text)
                    CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
                    dr("CyclType") = If(CyclType.Text <> "", CyclType.Text, Convert.DBNull)
                    Dim vClassNum As String = ""
                    If ClassCount.Text = "" Then vClassNum = "01"
                    If Len(ClassCount.Text) = 1 Then vClassNum = "0" & ClassCount.Text
                    If Len(ClassCount.Text) > 1 Then vClassNum = ClassCount.Text
                    dr("ClassNum") = vClassNum
                    dr("LastState") = "M" 'M: 修改(最後異動狀態)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                DbAccess.UpdateDataTable(dt, da, Trans)
                DbAccess.CommitTrans(Trans)
                'Call TIMS.CloseDbConn(conn)
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg6)
                'Common.MessageBox(Me, ex.ToString)
                Throw ex
            End Try
        End If
    End Sub

    '(儲存) iNum:'1是正式 '2是草稿
    Private Sub Insert_Plan_Table(ByVal iNum As Integer)
        'iNum:'1是正式 '2是草稿
        'Dim SeqNO As Integer
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim ReportCount As String = ""
        Dim re_update_flag As Boolean = False '重新UPDATE
        Dim AppliedResult1 As String = ""
        Dim str_CommandArgument As String = "" 'TC_01_014.aspx用
        Dim sql As String = ""
        If Request(cst_ccopy) = "1" Then
            '外部copy而來
            ReportCount = 0
        Else
            PlanID_value = sm.UserInfo.PlanID 'TIMS.ClearSQM(Request("PlanID"))
            ComIDNO_value = ComidValue.Value 'TIMS.ClearSQM(Request("ComIDNO"))
            SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
            sql = ""
            sql &= " SELECT COUNT(1) ReportCount FROM PLAN_VERREPORT WHERE 1=1" & vbCrLf
            sql &= " and PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'" & vbCrLf
            If g_flagNG Then sql = " SELECT COUNT(1) ReportCount FROM PLAN_VERREPORT WHERE 1<>1"
            ReportCount = DbAccess.ExecuteScalar(sql, objconn)
        End If
        If ConNum.Text <> "" Then ConNum.Text = TIMS.ChangeIDNO(ConNum.Text)

        '計畫別：產業人才投資計畫呈現「A」、提升自主勞工計畫呈現「B」
        Dim vPLAN1 As String = TIMS.Get_PSNO28_PLAN1(ComidValue.Value, objconn)
        '取得產投流水號(前6碼)年度別(3)+上下年(1)+計畫別(1)+分署別(1) 課程申請流水號
        Dim vPSNO28_6 As String = TIMS.Get_PSNO28_6(sm.UserInfo.Years, STDate.Text, vPLAN1, sm.UserInfo.DistID)
        If iNum = 1 AndAlso Len(vPSNO28_6) <> 6 Then
            '(正式才檢核)'取得產投流水號(前6碼) '長度應該為6 
            Common.MessageBox(Me, cst_errmsg2)
            Exit Sub
        End If

        'If Age_l.Text <> "" Then Age_l.Text = TIMS.ChangeIDNO(Age_l.Text)
        'If Age_u.Text <> "" Then Age_u.Text = TIMS.ChangeIDNO(Age_u.Text)
        'If Me.Age_l.Text <> "" AndAlso Me.Age_u.Text <> "" Then
        '    If CInt(Me.Age_u.Text) < CInt(Me.Age_l.Text) Then
        '        Dim tmpAge As String = Me.Age_u.Text
        '        Me.Age_u.Text = Me.Age_l.Text
        '        Me.Age_l.Text = tmpAge
        '    End If
        'End If

        If ClassCount.Text <> "" Then
            ClassCount.Text = TIMS.ChangeIDNO(ClassCount.Text)
        End If
        FirstSort.Text = TIMS.ClearSQM(FirstSort.Text)
        If FirstSort.Text <> "" Then
            FirstSort.Text = TIMS.ChangeIDNO(FirstSort.Text)
            If IsNumeric(FirstSort.Text) Then
                FirstSort.Text = CInt(FirstSort.Text)
            Else
                FirstSort.Text = "1"
            End If
        End If

        Dim v_Taddress2 As String = TIMS.GetListValue(Taddress2)
        Dim v_Taddress3 As String = TIMS.GetListValue(Taddress3)
        Dim vsTaddressZip As String = ""
        Dim vsTAddress As String = ""
        Dim vsTaddressZIP6W As String = ""
        Dim tmpPTID As String = If(v_Taddress2 <> "", v_Taddress2, If(v_Taddress3 <> "", v_Taddress3, ""))
        TIMS.GetTaddressPTID(objconn, tmpPTID, vsTaddressZip, vsTAddress, vsTaddressZIP6W)

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                Dim dr As DataRow = Nothing
                Try
                    If Me.upt_PlanX.Value <> "" Then '有儲存資料過了
                        '有儲存資料過了 '準備儲存資料
                        tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了

                        PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                        ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                        SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                        sql = "SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        dr = dt.Rows(0)
                    Else
                        If (Convert.ToString(Request("PlanID")) = "" OrElse Convert.ToString(Request(cst_ccopy)) = "1") Then
                            '新增資料 、ccopy=1 、草稿新增 而來
                            PlanID_value = sm.UserInfo.PlanID
                            ComIDNO_value = ComidValue.Value
                            SeqNO_value = GetMaxSeqNum(Trans, TransConn) '+1 ComidValue.Value  sm.UserInfo.PlanID
                            '準備儲存資料
                            sql = "SELECT * FROM PLAN_PLANINFO WHERE 1<>1"
                            dt = DbAccess.GetDataTable(sql, da, Trans)
                            dr = dt.NewRow
                            dt.Rows.Add(dr)
                            dr("PlanID") = sm.UserInfo.PlanID
                            dr("ComIDNO") = ComidValue.Value
                            dr("SeqNO") = SeqNO_value

                            dr("RID") = RIDValue.Value '空的才存取
                            dr("PlanYear") = Label3.Text '空的才存取
                            dr("TPlanID") = TPlanID '空的才存取
                            '預防新增時選擇草稿儲存
                            '導致因為停留原畫面，再儲存時會第二次重複儲存。
                            ' ViewState("SeqNO") = SeqNO_value
                            Org.Disabled = True
                        Else
                            '修改
                            PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                            ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                            SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                            sql = "SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                            dt = DbAccess.GetDataTable(sql, da, Trans)
                            dr = dt.Rows(0)
                        End If
                        tmpPCS = ""
                        TIMS.SetMyValue(tmpPCS, "PlanID", PlanID_value)
                        TIMS.SetMyValue(tmpPCS, "ComIDNO", ComIDNO_value)
                        TIMS.SetMyValue(tmpPCS, "SeqNO", SeqNO_value)
                        Me.upt_PlanX.Value = tmpPCS
                    End If
                Catch ex As Exception
                    Me.upt_PlanX.Value = ""
                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Me, cst_errmsg6)
                    'Common.MessageBox(Me, ex.ToString)
                    Exit Sub
                End Try

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    dr("TMID") = If(Me.jobValue.Value <> "", jobValue.Value, Convert.DBNull)
                Else
                    dr("TMID") = If(Me.trainValue.Value <> "", trainValue.Value, Convert.DBNull)
                End If
                dr("CJOB_UNKEY") = If(Me.cjobValue.Value <> "", cjobValue.Value, Convert.DBNull)

                dr("PlanCause") = If(Me.PlanCause.Text <> "", Me.PlanCause.Text, Convert.DBNull)
                dr("PurScience") = If(Me.PurScience.Text <> "", Me.PurScience.Text, Convert.DBNull)
                dr("PurTech") = If(Me.PurTech.Text <> "", Me.PurTech.Text, Convert.DBNull)
                dr("PurMoral") = If(Me.PurMoral.Text <> "", Me.PurMoral.Text, Convert.DBNull)

                Dim v_Degree As String = TIMS.GetListValue(Degree)
                dr("CapDegree") = If(v_Degree <> "", v_Degree, "00") 'Convert.DBNull
                'Dim tmpValue As String = ""
                'If Degree.SelectedIndex <> 0 Then
                '    tmpValue = TIMS.ClearSQM(Me.Degree.SelectedValue)
                '    dr("CapDegree") = tmpValue 'Trim(Me.Degree.SelectedValue)
                'End If
                Dim v_AppStage As String = TIMS.GetListValue(AppStage)
                dr("AppStage") = If(v_AppStage <> "", v_AppStage, Convert.DBNull)
                'If AppStage.SelectedIndex <> 0 Then
                '    tmpValue = TIMS.ClearSQM(Me.AppStage.SelectedValue)
                '    dr("AppStage") = tmpValue 'Trim(Me.AppStage.SelectedValue)
                'End If

                If rdoAge1.Checked Then
                    dr("CapAge1") = "15" '15歲以上
                End If
                If rdoAge2.Checked AndAlso txtAge1.Text <> "" Then
                    dr("CapAge1") = Val(txtAge1.Text) '15歲以上
                End If
                dr("CapAge2") = Convert.DBNull
                '性別：不區分 Convert.DBNull 
                dr("CapSex") = Convert.DBNull 'IIf(Sex.SelectedIndex = 0, Convert.DBNull, Me.Sex.SelectedValue)
                '兵役：不限塞00
                dr("CapMilitary") = "00" 'IIf(Solder.SelectedIndex = 0, "00", Me.Solder.SelectedValue) '不限塞00
                dr("CapOther1") = Convert.DBNull
                If Other1.Text <> "" AndAlso Other1.Text <> Cst_msgother1 Then
                    dr("CapOther1") = Other1.Text
                End If
                dr("CapOther2") = Convert.DBNull
                If Other2.Text <> "" AndAlso Other2.Text <> Cst_msgother1 Then
                    dr("CapOther2") = Other2.Text
                End If
                dr("CapOther3") = Convert.DBNull
                If Other3.Text <> "" AndAlso Other3.Text <> Cst_msgother1 Then
                    dr("CapOther3") = Other3.Text
                End If
                dr("TMScience") = If(TMScience.Text <> "", Me.TMScience.Text, Convert.DBNull)

                dr("GenSciHours") = If(GenSciHours.Text <> "", Me.GenSciHours.Text, Convert.DBNull)
                dr("ProSciHours") = If(ProSciHours.Text <> "", Me.ProSciHours.Text, Convert.DBNull)
                dr("ProTechHours") = If(ProTechHours.Text <> "", Me.ProTechHours.Text, Convert.DBNull)
                If Me.OtherHours.Text <> "" Then
                    dr("OtherHours") = Me.OtherHours.Text
                Else
                    dr("OtherHours") = Convert.DBNull
                End If
                dr("TotalHours") = If(TotalHours.Text <> "", Me.TotalHours.Text, Convert.DBNull)
                'If DefGovCost.Text = "" Then
                '    dr("DefGovCost") = Convert.DBNull
                'Else
                '    dr("DefGovCost") = DefGovCost.Text
                'End If
                dr("DefUnitCost") = Convert.DBNull
                If DefUnitCost.Text <> "" Then
                    dr("DefUnitCost") = CInt(Val(DefUnitCost.Text))
                End If

                'If DefStdCost.Text = "" Then
                '    dr("DefStdCost") = Convert.DBNull
                'Else
                '    dr("DefStdCost") = DefStdCost.Text
                'End If

                If Me.hTPlanID54.Value = "1" Then
                    '充飛 100% 0%
                    dr("DefGovCost") = Convert.DBNull
                    dr("DefStdCost") = Convert.DBNull
                    If TotalCost3.Text <> "" Then
                        dr("DefGovCost") = CInt(Val(DefGovCost.Text)) '100%
                        dr("DefStdCost") = CInt(Val(DefStdCost.Text)) '0%
                    End If
                Else
                    '產投使用總價，並自動切80% 20%
                    Select Case RadioButtonList1.SelectedValue
                        Case cst_學分班
                            '學分班
                            dr("DefGovCost") = Convert.DBNull
                            dr("DefStdCost") = Convert.DBNull
                            If TotalCost3.Text <> "" Then
                                dr("DefGovCost") = CInt(Val(TotalCost3.Text) * 0.8) '0.8
                                dr("DefStdCost") = Val(TotalCost3.Text) - CInt(Val(TotalCost3.Text) * 0.8) '0.2
                            End If
                        Case Else
                            '非學分班(總價)
                            dr("DefGovCost") = Convert.DBNull
                            dr("DefStdCost") = Convert.DBNull
                            If TotalCost2.Text <> "" Then
                                dr("DefGovCost") = CInt(Val(TotalCost2.Text) * 0.8) '0.8
                                dr("DefStdCost") = Val(TotalCost2.Text) - CInt(Val(TotalCost2.Text) * 0.8) '0.2
                            End If
                    End Select
                End If

                'IF ProcID.
                'dr("ProcID") = "" 'ClassChar.SelectedValue
                'ProcID 2008 拿掉，因為完全沒有用到，寫了也是白寫  by amu 2008-01-14
                If FirstSort.Text <> "" Then
                    dr("FirstSort") = FirstSort.Text
                Else
                    dr("FirstSort") = Convert.DBNull
                End If
                Select Case RadioButtonList1.SelectedValue
                    Case cst_學分班 ' "Y"
                        dr("PointYN") = RadioButtonList1.SelectedValue
                    Case Else
                        If RadioButtonList1.SelectedValue <> "" Then
                            dr("PointYN") = RadioButtonList1.SelectedValue
                        Else
                            dr("PointYN") = Convert.DBNull
                        End If
                End Select

                dr("PointType") = Convert.DBNull
                If PointType.SelectedValue <> "" Then
                    dr("PointType") = PointType.SelectedValue '學分種類
                End If
                dr("PackageType") = Convert.DBNull
                If PackageType.SelectedValue <> "" Then
                    dr("PackageType") = PackageType.SelectedValue '包班種類
                End If
                dr("SciPlaceID") = SciPlaceID.SelectedValue
                dr("TechPlaceID") = TechPlaceID.SelectedValue
                dr("SciPlaceID2") = SciPlaceID2.SelectedValue
                dr("TechPlaceID2") = TechPlaceID2.SelectedValue

                If PointName.Text <> "" Then PointName.Text = TIMS.ClearSQM(PointName.Text)
                If PackageName.Text <> "" Then PackageName.Text = TIMS.ClearSQM(PackageName.Text)
                ClassName.Text = Replace(ClassName.Text, "&", "＆")
                ClassName.Text = TIMS.ClearSQM(ClassName.Text)
                ClassName.Text = Replace(ClassName.Text, PointName.Text, "") '學分班種類
                ClassName.Text = Replace(ClassName.Text, PackageName.Text, "") '企業包班種類
                'dr("ClassName") = If(ClassName.Text <> "", Me.ClassName.Text & PointName.Text, Convert.DBNull)

                Dim vsClassName As String = ""
                Select Case RadioButtonList1.SelectedValue
                    Case cst_學分班 ' "Y"
                        vsClassName = ClassName.Text & PointName.Text & PackageName.Text
                        'dr("ClassName") = Trim(ClassName.Text) & Trim(PointName.Text) & Trim(PackageName.Text)
                    Case Else 'cst_非學分班
                        vsClassName = ClassName.Text & PackageName.Text
                        'dr("ClassName") = Trim(ClassName.Text) & Trim(PackageName.Text)
                End Select
                dr("ClassName") = vsClassName

                dr("Class_Unit") = If(Class_Unit.Value <> "", Me.Class_Unit.Value, Convert.DBNull)
                dr("TNum") = If(TNum.Text <> "", TNum.Text, Convert.DBNull)
                dr("THours") = If(THours.Text <> "", THours.Text, Convert.DBNull)
                dr("STDate") = If(STDate.Text <> "", Me.STDate.Text, Convert.DBNull)
                dr("FDDate") = If(FDDate.Text <> "", Me.FDDate.Text, Convert.DBNull)

                'Dim v_Taddress2 As String = TIMS.GetListValue(Taddress2)
                dr("AddressSciPTID") = If(v_Taddress2 <> "", v_Taddress2, Convert.DBNull)
                'Dim v_Taddress3 As String = TIMS.GetListValue(Taddress3)
                dr("AddressTechPTID") = If(v_Taddress3 <> "", v_Taddress3, Convert.DBNull)

                dr("TaddressZip") = If(vsTaddressZip <> "", vsTaddressZip, Convert.DBNull)
                dr("TaddressZIP6W") = If(vsTaddressZIP6W <> "", vsTaddressZIP6W, Convert.DBNull)
                dr("TAddress") = If(vsTAddress <> "", vsTAddress, Convert.DBNull)

                CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
                dr("CyclType") = If(CyclType.Text <> "", CyclType.Text, Convert.DBNull)
                dr("ClassCount") = If(ClassCount.Text <> "", Me.ClassCount.Text, Convert.DBNull)
                dr("CredPoint") = If(CredPoint.Text <> "", Me.CredPoint.Text, Convert.DBNull)
                dr("RoomName") = If(RoomName.Text <> "", Me.RoomName.Text, Convert.DBNull)
                dr("FactMode") = If(FactMode.SelectedIndex = -1, Convert.DBNull, FactMode.SelectedValue)
                If FactMode.SelectedValue = "99" Then
                    dr("FactModeOther") = If(FactModeOther.Text <> "", Me.FactModeOther.Text, Convert.DBNull)
                Else
                    dr("FactModeOther") = Convert.DBNull
                End If
                dr("ConNum") = If(ConNum.Text <> "", ConNum.Text, Convert.DBNull)
                dr("ContactName") = If(ContactName.Text <> "", Me.ContactName.Text, Convert.DBNull)
                dr("ContactPhone") = If(ContactPhone.Text <> "", Me.ContactPhone.Text, Convert.DBNull)
                ContactEmail.Text = TIMS.ChangeEmail(ContactEmail.Text)
                dr("ContactEmail") = If(ContactEmail.Text <> "", Me.ContactEmail.Text, Convert.DBNull)
                dr("ContactFax") = If(ContactFax.Text <> "", Me.ContactFax.Text, Convert.DBNull)
                '訓練職能 
                dr("ClassCate") = If(ClassCate.SelectedIndex <> 0, Me.ClassCate.SelectedValue, Convert.DBNull)
                '早上'下午'晚上
                'Dim vTPERIOD28_C1 As String = "N" '早上
                'Dim vTPERIOD28_C2 As String = "N" '下午
                'Dim vTPERIOD28_C3 As String = "N" '晚上
                'If TPERIOD28_C1.Checked Then vTPERIOD28_C1 = "Y"
                'If TPERIOD28_C2.Checked Then vTPERIOD28_C2 = "Y"
                'If TPERIOD28_C3.Checked Then vTPERIOD28_C3 = "Y"
                'Dim vTPERIOD28 As String = vTPERIOD28_C1 & vTPERIOD28_C2 & vTPERIOD28_C3
                'dr("TPERIOD28") = vTPERIOD28 '(必定有值)
                '課程大綱 
                dr("Content") = If(Content.Text <> "", Me.Content.Text, Convert.DBNull)

                Select Case RadioButtonList1.SelectedValue
                    Case cst_學分班 ' "Y"
                        '學分班(總價)
                        dr("TotalCost") = If(TotalCost3.Text <> "", Me.TotalCost3.Text, Convert.DBNull)
                        dr("Note2") = If(tNote2b.Text <> "", tNote2b.Text, Convert.DBNull)
                    Case Else 'cst_非學分班
                        '非學分班(總價)
                        dr("TotalCost") = If(TotalCost2.Text <> "", Me.TotalCost2.Text, Convert.DBNull)
                        dr("Note2") = If(tNote2.Text <> "", tNote2.Text, Convert.DBNull)
                End Select
                'TotalCost3
                dr("Note") = Convert.DBNull
                If Me.Note.Text <> "" Then
                    dr("Note") = Me.Note.Text
                End If

                If iNum = 1 AndAlso Convert.ToString(dr("AppliedDate")) = "" Then
                    dr("AppliedDate") = Now.Date
                End If
                dr("AppliedOrigin") = 1

                'dr("AppliedResult") = Convert.DBNull
                If iNum = 1 Then '計畫為正式
                    If (Convert.ToString(Request("PlanID")) = "" OrElse Convert.ToString(Request(cst_ccopy)) = "1") Then
                        '新增的狀況
                        If iPlanKind = 1 Then
                            '自辦
                            dr("AppliedResult") = "Y" '分署內訓計畫為審核通過
                        Else '委辦
                            '分署(中心)不動
                            If sm.UserInfo.LID > 1 Then
                                '委訓清空
                                dr("AppliedResult") = Convert.DBNull
                            End If
                        End If
                    Else
                        '修改的狀況
                        If iPlanKind = 1 Then '自辦
                            dr("AppliedResult") = "Y" '分署(職訓中心)內訓計畫為審核通過
                        Else '委辦(取得 舊值 存入 AppliedResult1)
                            re_update_flag = True
                            AppliedResult1 = dr("AppliedResult").ToString
                            '分署(中心)不動
                            If sm.UserInfo.LID > 1 Then
                                '委訓清空
                                dr("AppliedResult") = Convert.DBNull
                            End If
                            'dr("AppliedResult") = Convert.DBNull
                        End If
                    End If
                End If

                dr("PlanEMail") = Convert.DBNull
                If EMail.Text <> "" Then
                    EMail.Text = TIMS.ChangeEmail(EMail.Text)
                    dr("PlanEMail") = Me.EMail.Text
                End If

                If iNum = 1 Then '正式
                    If Convert.ToString(dr("TransFlag")) = "" Then
                        dr("TransFlag") = "N"
                    End If
                    dr("IsApprPaper") = "Y"
                End If

                'dr("IsBusiness") = If(IsBusiness.Checked = True, "Y", "N")
                'dr("IsBusiness") = "N" '1:非包班
                dr("IsBusiness") = "Y"
                Select Case PackageType.SelectedValue
                    Case "1"        '非包班
                        dr("IsBusiness") = "N"
                    Case "2", "3"   '2:企業包班,3:聯合企業包班
                        dr("IsBusiness") = "Y"
                End Select

                dr("EnterpriseName") = EnterpriseName.Text

                'G:非勞工團體 W:勞工團體
                dr("EnterSupplyStyle") = Convert.DBNull
                If EnterSupplyStyle.SelectedValue <> "" Then
                    dr("EnterSupplyStyle") = EnterSupplyStyle.SelectedValue
                End If

                dr("GCID") = Convert.DBNull
                dr("GCID2") = Convert.DBNull
                Select Case strYears
                    Case "2014"
                        If GCIDValue.Value <> "" Then
                            dr("GCID") = GCIDValue.Value
                        End If
                    Case "2015"
                        If GCIDValue.Value <> "" Then
                            dr("GCID2") = GCIDValue.Value
                        End If
                End Select

                dr("AppStage") = Convert.DBNull
                If AppStage.SelectedValue <> "" Then
                    dr("AppStage") = AppStage.SelectedValue
                End If

                If (sm.UserInfo.LID = 2) Then
                    dr("ResultButton") = "Y" '被修改後尚未送出-待送出
                End If
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                'ProcessType @Insert/Update/View
                If ReportCount = "0" Then str_CommandArgument = "&ProcessType=Insert" Else str_CommandArgument = "&ProcessType=Update"

                Dim sCmdArg As String = ""
                sCmdArg = ""
                sCmdArg += "&PlanYear=" & Convert.ToString(dr("PlanYear"))
                sCmdArg += "&PlanID=" & Convert.ToString(dr("PlanID"))
                sCmdArg += "&TPlanID=" & Convert.ToString(dr("TPlanID"))
                sCmdArg += "&TMID=" & Convert.ToString(dr("TMID"))
                sCmdArg += "&RID=" & Convert.ToString(dr("RID"))
                sCmdArg += "&ComIDNO=" & Convert.ToString(dr("ComIDNO"))
                sCmdArg += "&SeqNO=" & Convert.ToString(dr("SeqNO"))
                str_CommandArgument &= sCmdArg

                'str_CommandArgument += "&PlanYear=" & dr("PlanYear")
                'str_CommandArgument += "&ClassName=" & Server.UrlEncode(dr("ClassName").ToString)
                'str_CommandArgument += "&ClassCate=" & dr("ClassCate")
                ''str_CommandArgument += "&ClassID=" & dr("ClassID")
                'str_CommandArgument += "&PlanID=" & dr("PlanID")
                'str_CommandArgument += "&TPlanID=" & dr("TPlanID")
                'str_CommandArgument += "&TMID=" & dr("TMID")
                'str_CommandArgument += "&RID=" & dr("RID")
                'str_CommandArgument += "&ComIDNO=" & dr("ComIDNO")
                'str_CommandArgument += "&SeqNO=" & dr("SeqNO")
                'str_CommandArgument += "&TNum=" & dr("TNum")
                'str_CommandArgument += "&THours=" & dr("THours")
                'str_CommandArgument += "&STDate=" & dr("STDate")
                'str_CommandArgument += "&FDDate=" & dr("FDDate")
                ''str_CommandArgument += "&ProcID=" & dr("ProcID")
                'str_CommandArgument += "&PointYN=" & dr("PointYN")
                'str_CommandArgument += "&TPeriod=" & ""
                'str_CommandArgument += "&Times=" & ""
                'str_CommandArgument += "&DefGovCost=" & dr("DefGovCost")
                'str_CommandArgument += "&DefStdCost=" & dr("DefStdCost")
                'str_CommandArgument += "&CapDegree=" & dr("CapDegree")
                'str_CommandArgument += "&AppStage=" & dr("AppStage")

                DbAccess.UpdateDataTable(dt, da, Trans)
                DbAccess.CommitTrans(Trans)
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg6)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try

        End Using

        '若為正式儲存且為分署(中心)，一並更改 Class_ClassInfo 資料
        If iNum = 1 And sm.UserInfo.LID = 1 Then
            Using TransConn As SqlConnection = DbAccess.GetConnection()
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Call Update_Class_ClassInfo(PlanID_value, ComIDNO_value, SeqNO_value, TransConn, Trans)
            End Using
        End If

        '若有 訓練計劃開班總表(產學訓) 則一並更改課程大綱(Plan_VerReport)
        If iNum = 1 And ReportCount > "0" Then
            Using TransConn As SqlConnection = DbAccess.GetConnection()
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Call Update_Plan_VerReport(PlanID_value, ComIDNO_value, SeqNO_value, TransConn, Trans)
            End Using
        End If

        'Try

        'Catch ex As Exception
        '    Me.upt_PlanX.Value = ""
        '    Throw ex
        'End Try

        If iNum = 1 Then
            '若為正式儲存'(更新PSNO28) 目前最大值(+1) 課程申請流水號
            Hid_PSNO28.Value = TIMS.UPDATE_PSNO28xPLANINFO(PlanID_value, ComIDNO_value, SeqNO_value, vPSNO28_6)
            If Hid_PSNO28.Value = "" Then
                Common.MessageBox(Me, cst_errmsg2)
                Exit Sub
            End If
        End If

        Dim dtTemp As DataTable
        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                'If iNum = 1 Then
                '若為正式儲存() (更新PSNO28) 目前最大值(+1)
                'Call TIMS.UPDATE_PCS_PSNO28(PlanID_value, ComIDNO_value, SeqNO_value, conn, Trans, vPSNO28_6)
                'End If

                'update (97產學訓課程大綱)(Plan_TrainDesc) 'Plan_Teacher
                If Not Session(cst_TrainDescTable) Is Nothing Then
                    dtTemp = Session(cst_TrainDescTable)
                    ''新增'或COPY
                    'sql = "SELECT * FROM PLAN_TRAINDESC where 1<>1"
                    'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                    '    '非:新增'或COPY '是:修改
                    '    sql = "SELECT * FROM PLAN_TRAINDESC where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                    'End If
                    da = New SqlDataAdapter
                    sql = "SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        If Not dr.RowState = DataRowState.Deleted Then
                            Dim iPTDID As Integer = 0
                            If dr("PTDID") <= 0 Then iPTDID = DbAccess.GetNewId(Trans, "PLAN_TRAINDESC_PTDID_SEQ,PLAN_TRAINDESC,PTDID")
                            If dr("PTDID") <= 0 Then dr("PTDID") = iPTDID
                            dr("PlanID") = PlanID_value
                            dr("ComIDNO") = ComIDNO_value
                            dr("SeqNO") = SeqNO_value
                        End If
                    Next
                    dt = dtTemp.Copy
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    '班級申請老師檔 'Plan_Teacher
                    Dim pms_S1 As New Hashtable
                    Dim sql_S1 As String = " SELECT 1 FROM PLAN_TEACHER WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNo='" & SeqNO_value & "' "
                    Dim dtS1 As DataTable = DbAccess.GetDataTable(sql_S1, Trans, pms_S1)
                    If TIMS.dtHaveDATA(dtS1) Then
                        Dim sql_D As String = " DELETE PLAN_TEACHER WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNo='" & SeqNO_value & "' "
                        DbAccess.ExecuteNonQuery(sql_D, Trans)
                    End If

                    'Dim CostIDArray As New ArrayList
                    da = New SqlDataAdapter
                    Dim aaTechID As String = ""
                    sql = " SELECT * FROM PLAN_TEACHER WHERE 1<>1"
                    dtTemp = DbAccess.GetDataTable(sql, da, Trans)
                    'Plan_TrainDesc
                    aaTechID = ""
                    For Each drw As DataRow In dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        If Not drw.RowState = DataRowState.Deleted _
                        AndAlso drw("TechID").ToString <> "" Then

                            '=========== 匯入時不要有相同 Start =============
                            Dim MYKEY As String = drw("TechID").ToString '本次師資
                            Dim Flag As Boolean = True '可新增
                            Flag = True '可新增
                            If aaTechID.IndexOf(MYKEY) > -1 Then
                                Flag = False  '不可新增
                            Else
                                aaTechID += "," & MYKEY '加入 可新增
                            End If
                            'For i As Integer = 0 To CostIDArray.Count - 1
                            '    If CostIDArray(i) = MYKEY Then Flag = False
                            'Next
                            If Flag Then
                                'CostIDArray.Add(MYKEY)
                                Dim dr As DataRow = dtTemp.NewRow
                                dtTemp.Rows.Add(dr)
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value

                                dr("TechID") = drw("TechID")

                                dr("ModifyAcct") = sm.UserInfo.UserID
                                dr("ModifyDate") = Now
                            End If
                            '=========== 匯入時不要有相同 -End- =============
                        End If
                    Next
                    dt = dtTemp.Copy
                    DbAccess.UpdateDataTable(dt, da, Trans)
                End If
                DbAccess.CommitTrans(Trans)
                Session(cst_TrainDescTable) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg7)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        'Call TIMS.OpenDbConn(conn)
        'Trans = DbAccess.BeginTrans(conn)
        'Try
        '    DbAccess.UpdateDataTable(dt, da, Trans)
        '    DbAccess.CommitTrans(Trans)
        'Catch ex As Exception
        '    DbAccess.RollbackTrans(Trans)
        '    Call TIMS.CloseDbConn(conn)
        '    Common.MessageBox(Me, "儲存資料有誤!!!")
        '    Common.MessageBox(Me, ex.ToString)
        '    Exit Sub
        'End Try

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                Select Case RadioButtonList1.SelectedValue
                    Case cst_學分班 ' "Y"
                        '學分班不儲存 CostItemTable (PLAN_COSTITEM)
                        Session(cst_CostItemTable) = Nothing
                        Dim pms_S1 As New Hashtable
                        Dim sql_S1 As String = " SELECT 1 FROM PLAN_COSTITEM WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "' and CostMode='5'"
                        Dim dtS1 As DataTable = DbAccess.GetDataTable(sql_S1, Trans, pms_S1)
                        If TIMS.dtHaveDATA(dtS1) Then
                            Dim sql_D1 As String = " DELETE PLAN_COSTITEM WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "' and CostMode='5'"
                            DbAccess.ExecuteNonQuery(sql_D1, Trans)
                        End If
                        'Case Else 'cst_非學分班
                End Select
                'update 計畫經費項目檔(Plan_CostItem)
                If Not Session(cst_CostItemTable) Is Nothing Then
                    dtTemp = Session(cst_CostItemTable) : da = Nothing
                    If dtTemp.Rows.Count > 0 Then
                        ''新增'或COPY
                        ''sql = "select * from Plan_CostItem 1<>1"
                        'sql = "select * from Plan_CostItem where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNO='" & SeqNO_value & "' and CostMode='5' "
                        'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                        '    '非:新增'或COPY '是:修改
                        '    sql = "select * from Plan_CostItem where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "' and CostMode='5' "
                        'End If
                        sql = "SELECT * FROM PLAN_COSTITEM WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "' and CostMode='5'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If dr("PCID") <= 0 Then
                                    dr("PCID") = DbAccess.GetNewId(Trans, "PLAN_COSTITEM_PCID_SEQ,PLAN_COSTITEM,PCID")
                                End If
                                'If Convert.ToString(dr("PlanID")) <> PlanID_value Then dr("PlanID") = PlanID_value
                                'If Convert.ToString(dr("ComIDNO")) <> ComIDNO_value Then dr("ComIDNO") = ComIDNO_value
                                'If Convert.ToString(dr("SeqNO")) <> SeqNO_value Then dr("SeqNO") = SeqNO_value
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(cst_CostItemTable) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg8)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                'update 計畫材料品名項目檔(Plan_Material)
                If Not Session(Cst_MaterialTable) Is Nothing Then
                    dtTemp = Session(Cst_MaterialTable)
                    If dtTemp.Rows.Count > 0 Then
                        ''新增'或COPY
                        ''sql = "select * from Plan_Material where 1<>1"
                        'sql = "select * from Plan_Material where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNO='" & SeqNO_value & "'"
                        'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                        '    '非:新增'或COPY '是:修改
                        '    sql = "select * from Plan_Material where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                        'End If
                        sql = "SELECT * FROM PLAN_MATERIAL WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                'PLAN_MATERIAL_PMID_SEQ
                                If dr("PMID") <= 0 Then
                                    dr("PMID") = DbAccess.GetNewId(Trans, "PLAN_MATERIAL_PMID_SEQ,PLAN_MATERIAL,PMID")
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(Cst_MaterialTable) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg9)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If Not Session(Cst_PersonCostTable) Is Nothing Then
                    dtTemp = Session(Cst_PersonCostTable)
                    If dtTemp.Rows.Count > 0 Then
                        ''新增'或COPY
                        ''sql = "select * from Plan_PersonCost where 1<>1"
                        'sql = "select * from Plan_PersonCost where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNO='" & SeqNO_value & "'"
                        'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                        '    '非:新增'或COPY '是:修改
                        '    sql = "select * from Plan_PersonCost where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                        'End If
                        sql = "SELECT * FROM PLAN_PERSONCOST WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If dr("PPCID") <= 0 Then
                                    dr("PPCID") = DbAccess.GetNewId(Trans, "PLAN_PERSONCOST_PPCID_SEQ,PLAN_PERSONCOST,PPCID")
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(Cst_PersonCostTable) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg10)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If Not Session(Cst_CommonCostTable) Is Nothing Then
                    dtTemp = Session(Cst_CommonCostTable)
                    If dtTemp.Rows.Count > 0 Then
                        ''新增'或COPY
                        ''sql = "select * from Plan_CommonCost where 1<>1"
                        'sql = "select * from Plan_CommonCost where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNO='" & SeqNO_value & "'"
                        'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                        '    '非:新增'或COPY '是:修改
                        '    sql = "select * from Plan_CommonCost where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                        'End If
                        sql = "SELECT * FROM PLAN_COMMONCOST WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If dr("PCMID") <= 0 Then
                                    dr("PCMID") = DbAccess.GetNewId(Trans, "PLAN_COMMONCOST_PCMID_SEQ,PLAN_COMMONCOST,PCMID")
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value

                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(Cst_CommonCostTable) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg11)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If Not Session(Cst_SheetCostTable) Is Nothing Then
                    dtTemp = Session(Cst_SheetCostTable)
                    If dtTemp.Rows.Count > 0 Then
                        '新增'或COPY
                        sql = "SELECT * FROM PLAN_SHEETCOST WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If dr(Cst_SheetCostpkName) <= 0 Then
                                    dr(Cst_SheetCostpkName) = DbAccess.GetNewId(Trans, "PLAN_SHEETCOST_PSHID_SEQ,PLAN_SHEETCOST,PSHID")
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(Cst_SheetCostTable) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg12)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If Not Session(Cst_OtherCostTable) Is Nothing Then
                    dtTemp = Session(Cst_OtherCostTable)
                    If dtTemp.Rows.Count > 0 Then
                        '新增'或COPY
                        sql = "SELECT * FROM PLAN_OTHERCOST WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If dr(Cst_OtherCostpkName) <= 0 Then
                                    dr(Cst_OtherCostpkName) = DbAccess.GetNewId(Trans, "PLAN_OTHERCOST_POTID_SEQ,PLAN_OTHERCOST,POTID")
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(Cst_OtherCostTable) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg13)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                '上課時間
                If Not Session(cst_Plan_OnClass) Is Nothing Then
                    dtTemp = Session(cst_Plan_OnClass)
                    If dtTemp.Rows.Count > 0 Then
                        ''新增'或COPY
                        ''sql = "SELECT * FROM PLAN_ONCLASS where 1<>1"
                        'sql = "SELECT * FROM PLAN_ONCLASS where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNO='" & SeqNO_value & "'"
                        'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                        '    '非:新增'或COPY '是:修改
                        '    sql = "SELECT * FROM PLAN_ONCLASS where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                        'End If
                        sql = "SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If dr("POCID") <= 0 Then
                                    dr("POCID") = DbAccess.GetNewId(Trans, "PLAN_ONCLASS_POCID_SEQ,PLAN_ONCLASS,POCID")
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(cst_Plan_OnClass) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg14)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                '計畫包班事業單位
                ' Session(cst_Plan_BusPackage)
                Select Case PackageType.SelectedValue
                    Case "3" '充電起飛計畫 '聯合企業包班
                        If Not Session(cst_Plan_BusPackage) Is Nothing Then
                            'Const Cst_PKName As String = "BPID"
                            dtTemp = Session(cst_Plan_BusPackage)
                            If dtTemp.Rows.Count > 0 Then
                                ''新增'或COPY
                                ''sql = "select * from Plan_BusPackage where 1<>1"
                                'sql = "select * from Plan_BusPackage where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNO='" & SeqNO_value & "'"
                                'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                                '    '非:新增'或COPY '是:修改
                                '    sql = "select * from Plan_BusPackage where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                                'End If
                                sql = "SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                                dt = DbAccess.GetDataTable(sql, da, Trans)
                                For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                                    If Not dr.RowState = DataRowState.Deleted Then
                                        If dr("BPID") <= 0 Then
                                            dr("BPID") = DbAccess.GetNewId(Trans, "PLAN_BUSPACKAGE_BPID_SEQ,PLAN_BUSPACKAGE,BPID")
                                        End If
                                        dr("PlanID") = PlanID_value
                                        dr("ComIDNO") = ComIDNO_value
                                        dr("SeqNO") = SeqNO_value

                                    End If
                                Next
                                dt = dtTemp.Copy
                                DbAccess.UpdateDataTable(dt, da, Trans)
                            End If
                        End If

                    Case "2" '充電起飛計畫' 企業包班(只有1筆)
                        ''新增'或COPY
                        ''sql = "select * from Plan_BusPackage where 1<>1"
                        'sql = "select * from Plan_BusPackage where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNO='" & SeqNO_value & "'"
                        'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                        '    '非:新增'或COPY '是:修改
                        '    '修改
                        '    sql = " delete Plan_BusPackage where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                        '    DbAccess.ExecuteNonQuery(sql, Trans)
                        '    sql = "select * from Plan_BusPackage where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                        'End If
                        Dim pms_S1 As New Hashtable
                        Dim sql_S1 As String = " SELECT 1 FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        Dim dtS1 As DataTable = DbAccess.GetDataTable(sql_S1, Trans, pms_S1)
                        If TIMS.dtHaveDATA(dtS1) Then
                            Dim sql_D1 As String = " DELETE PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                            DbAccess.ExecuteNonQuery(sql_D1, Trans)
                        End If
                        sql = "SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)

                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("BPID") = DbAccess.GetNewId(Trans, "PLAN_BUSPACKAGE_BPID_SEQ,PLAN_BUSPACKAGE,BPID")
                        If (Convert.ToString(Request("PlanID")) = "" OrElse Convert.ToString(Request(cst_ccopy)) = "1") Then
                            dr("PlanID") = sm.UserInfo.PlanID
                            dr("ComIDNO") = ComidValue.Value
                            dr("SeqNO") = SeqNO_value ' ViewState("SeqNO")
                        Else
                            dr("PlanID") = TIMS.ClearSQM(Request("PlanID"))
                            dr("ComIDNO") = TIMS.ClearSQM(Request("ComIDNO"))
                            dr("SeqNO") = TIMS.ClearSQM(Request("SeqNO"))
                        End If
                        txtUname.Text = TIMS.ClearSQM(txtUname.Text)
                        dr("Uname") = txtUname.Text 'Convert.ToString(txtUname.Text.Trim)
                        dr("Intaxno") = TIMS.ChangeIDNO(txtIntaxno.Text)
                        dr("Ubno") = TIMS.ChangeIDNO(txtUbno.Text)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        DbAccess.UpdateDataTable(dt, da, Trans)

                    Case Else '清除
                        'If TIMS.ClearSQM(Request("PlanID") <> "" AndAlso Request(cst_ccopy) <> "1" Then
                        '    '非:新增'或COPY '是:修改
                        '    sql = "delete Plan_BusPackage where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNO='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                        '    DbAccess.ExecuteNonQuery(sql, Trans)
                        'Else
                        '    '是:新增'或COPY
                        '    sql = "delete Plan_BusPackage where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" & SeqNO_value & "'"
                        '    DbAccess.ExecuteNonQuery(sql, Trans)
                        'End If
                        Dim pms_S1 As New Hashtable
                        Dim sql_S1 As String = " SELECT 1 FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        Dim dtS1 As DataTable = DbAccess.GetDataTable(sql_S1, Trans, pms_S1)
                        If TIMS.dtHaveDATA(dtS1) Then
                            Dim sql_D1 As String = " DELETE PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                            DbAccess.ExecuteNonQuery(sql_D1, Trans)
                        End If
                End Select
                DbAccess.CommitTrans(Trans)
                Session(cst_Plan_BusPackage) = Nothing
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Common.MessageBox(Me, cst_errmsg15)
                'Common.MessageBox(Me, ex.ToString)
                Exit Sub
            End Try
        End Using
        'Call TIMS.CloseDbConn(conn) '結束資料連線

        '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
        If re_update_flag Then
            Select Case AppliedResult1
                Case "Y", "O", "N"
                Case Else '"R"
                    ''dr("AppliedResult") = Convert.DBNull
                    'TIMS.Plan_VerRecord_Update(sm.UserInfo.PlanID, ComidValue.Value, TIMS.ClearSQM(Request("SeqNo"))
                    ''TIMS.Plan_VerReprot_Update(sm.UserInfo.PlanID, ComidValue.Value, TIMS.ClearSQM(Request("SeqNo"), "O", "")
                    'TIMS.Plan_VerReprot_Update(sm.UserInfo.PlanID, ComidValue.Value, TIMS.ClearSQM(Request("SeqNo"), "", "O")
                    Call TIMS.Plan_VerRecord_Update(PlanID_value, ComIDNO_value, SeqNO_value, objconn)
                    Call TIMS.PLAN_VERREPROT_UPDATE(Me, PlanID_value, ComIDNO_value, SeqNO_value, "O", objconn)
            End Select
        End If

        '☆2011-12-11 for 28 ,54,56計畫，在正式儲存後要先提示訊息
        If iNum = 1 Then
            Common.RespWrite(Me, "<script language='javascript'>alert('請記得於班級查詢後，針對本班按下【送出】鍵！');</script>")
        End If

        If TIMS.ClearSQM(Request("PlanID")) = "" Then
            If iNum = 1 Then
                '正式送出
                'Common.RespWrite(Me, "<script>alert('計畫申請成功!!');location.href='TC_03_003.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</script>")
                Session("saveok") = True
                'PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                'ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                'SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If ReportCount > "0" Then
                    Common.RespWrite(Me, "<SCRIPT>if(confirm('是否要繼續編輯「開班計劃表資料維護」'))location.href='../01/TC_01_014_add.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & str_CommandArgument & "';else location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</SCRIPT>")
                Else
                    Common.RespWrite(Me, "<SCRIPT>if(confirm('是否要繼續新增「開班計劃表資料維護」'))location.href='../01/TC_01_014_add.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & str_CommandArgument & "';else location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</SCRIPT>")
                End If
                'Page.RegisterStartupScript()
                'Page.RegisterClientScriptBlock("存檔成功", "<SCRIPT>if(confirm('是否要繼續新增「開班計劃表資料維護」'))location='index.aspx' ;else location='WriteMessage.aspx'</SCRIPT>")
                'Page.RegisterStartupScript("計畫申請成功!!", "<SCRIPT>if(confirm('是否要繼續新增「開班計劃表資料維護」'))location='index.aspx' ;else location='WriteMessage.aspx'</SCRIPT>")
                'Session(cst_CostItemTable) = Nothing
                'Session(Cst_MaterialTable) = Nothing
                'Session(Cst_PersonCostTable) = Nothing
                'Session(Cst_CommonCostTable) = Nothing
                'Session(cst_TrainDescTable) = Nothing
            Else
                '這一格應該是沒有，因為正式儲存後，應該不會有草稿儲存。
                '(新增的)草稿儲存
                Common.RespWrite(Me, "<script>alert('草稿儲存成功!!');</script>")

                Call CreateClassTime()
                Call CreateCostItem()
                Call CreateMaterial()
                Call CreateTrainDesc()
                Call CreateBusPackage()
                Call CreatePersonCost()
                Call CreateCommonCost()

                'If TIMS.ClearSQM(Request("PlanID") = "" OrElse Request(cst_ccopy) = "1" Then
                '    Common.RespWrite(Me, "<script>location.href='./TC_03_003.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "&PlanID=" & sm.UserInfo.PlanID & "&ComIDNO=" & ComidValue.Value & "&SeqNO=" & SeqNO_value & "'</script>")
                'Else
                '    Common.RespWrite(Me, "<script>location.href='./TC_03_003.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "&PlanID=" & TIMS.ClearSQM(Request("PlanID") & "&ComIDNO=" & TIMS.ClearSQM(Request("ComIDNO") & "&SeqNO=" & TIMS.ClearSQM(Request("SeqNO") & "'</script>")
                'End If
                'If TIMS.ClearSQM(Request("PlanID") = "" OrElse Request(cst_ccopy) = "1" Then
                '    '新增或copy
                '    'sql = " select * from Plan_CostItem where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" &  ViewState("SeqNO") & "' and CostMode='5' "
                '    sql = " select * from Plan_CostItem where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" & SeqNO_value & "' and CostMode='5' "
                '    Session(cst_CostItemTable) = DbAccess.GetDataTable(sql, objconn)
                '    'sql = " select * from Plan_Material where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" &  ViewState("SeqNO") & "'"
                '    sql = " select * from Plan_Material where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" & SeqNO_value & "'"
                '    Session(Cst_MaterialTable) = DbAccess.GetDataTable(sql, objconn)
                '    sql = " select *,0 Total,0 subtotal from Plan_PersonCost where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" & SeqNO_value & "'"
                '    sql &= " ORDER BY ItemNo"
                '    Session(Cst_PersonCostTable) = DbAccess.GetDataTable(sql, objconn)
                '    sql = " select *,0 subtotal, 0 eachCost from Plan_CommonCost where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" & SeqNO_value & "'"
                '    sql &= " ORDER BY ItemNo"
                '    Session(Cst_CommonCostTable) = DbAccess.GetDataTable(sql, objconn)
                '    'sql = "SELECT * FROM PLAN_TRAINDESC where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" &  ViewState("SeqNO") & "'"
                '    sql = "SELECT * FROM PLAN_TRAINDESC where PlanID='" & sm.UserInfo.PlanID & "' and ComIDNO='" & ComidValue.Value & "' and SeqNo='" & SeqNO_value & "'"
                '    'sql &= " order by  STrainDate,PName  " '20090401 andy edit 加入排序
                '    sql &= " order by  STrainDate,PName, PTDID  " '20090518 andy edit   PTDID加入排序  
                '    Session(cst_TrainDescTable) = DbAccess.GetDataTable(sql, objconn)
                'Else
                '    '修改
                '    sql = " select * from Plan_CostItem where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNo='" & TIMS.ClearSQM(Request("SeqNO") & "' and CostMode='5' "
                '    Session(cst_CostItemTable) = DbAccess.GetDataTable(sql, objconn)
                '    sql = " select * from Plan_Material where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNo='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                '    Session(Cst_MaterialTable) = DbAccess.GetDataTable(sql, objconn)
                '    sql = " select *,0 Total,0 subtotal from Plan_PersonCost where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNo='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                '    sql &= " ORDER BY ItemNo"
                '    Session(Cst_PersonCostTable) = DbAccess.GetDataTable(sql, objconn)
                '    sql = " select *,0 subtotal, 0 eachCost from Plan_CommonCost where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNo='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                '    sql &= " ORDER BY ItemNo"
                '    Session(Cst_CommonCostTable) = DbAccess.GetDataTable(sql, objconn)
                '    sql = "SELECT * FROM PLAN_TRAINDESC where PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNo='" & TIMS.ClearSQM(Request("SeqNO") & "'"
                '    'sql &= " order by  STrainDate,PName  " '20090401 andy edit 加入排序
                '    sql &= " order by  STrainDate,PName, PTDID  " '20090518 andy edit   PTDID加入排序  
                '    Session(cst_TrainDescTable) = DbAccess.GetDataTable(sql, objconn)
                'End If
            End If
        Else
            If ViewState("search") <> "" Then
                Session("search") = ViewState("search")
            End If
            If iNum = 1 Then
                '正式送出
                Session("saveok") = True
                If ReportCount > "0" Then

                    If Request(cst_ccopy) = "1" Then
                        '班級複製作業
                        Common.RespWrite(Me, "<SCRIPT>if(confirm('是否要繼續編輯「開班計劃表資料維護」'))location.href='../01/TC_01_014_add.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & str_CommandArgument & "' ;else location.href='../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</SCRIPT>")
                        'Common.RespWrite(Me, "<script>location.href='../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</script>")
                    Else
                        '班級查詢作業
                        Common.RespWrite(Me, "<SCRIPT>if(confirm('是否要繼續編輯「開班計劃表資料維護」'))location.href='../01/TC_01_014_add.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & str_CommandArgument & "' ;else location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</SCRIPT>")
                    End If
                Else
                    Common.RespWrite(Me, "<SCRIPT>if(confirm('是否要繼續新增「開班計劃表資料維護」'))location.href='../01/TC_01_014_add.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & str_CommandArgument & "' ;else location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</SCRIPT>")
                End If

            Else
                '(修改的)草稿儲存
                'Common.MessageBox(Me, "草稿儲存成功!!")
                'Exit Sub
                Common.RespWrite(Me, "<script>alert('草稿儲存成功!!');</script>")
                CreateClassTime()
                CreateCostItem()
                CreateMaterial()
                CreateTrainDesc()
                CreateBusPackage()
                Call CreatePersonCost()
                Call CreateCommonCost()
                Call CreateSheetCost()
                Call CreateOtherCost()

                'If TIMS.ClearSQM(Request("PlanID") = "" OrElse Request(cst_ccopy) = "1" Then
                '    Common.RespWrite(Me, "<script>location.href='./TC_03_003.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "&PlanID=" & sm.UserInfo.PlanID & "&ComIDNO=" & ComidValue.Value & "&SeqNO=" & SeqNO_value & "'</script>")
                'Else
                '    Common.RespWrite(Me, "<script>location.href='./TC_03_003.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "&PlanID=" & TIMS.ClearSQM(Request("PlanID") & "&ComIDNO=" & TIMS.ClearSQM(Request("ComIDNO") & "&SeqNO=" & TIMS.ClearSQM(Request("SeqNO") & "'</script>")
                'End If
                'If Request(cst_ccopy) = "1" Then
                '    '班級複製作業
                '    Common.RespWrite(Me, "<script>location.href='../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</script>")
                'Else
                '    '班級查詢作業
                '    Common.RespWrite(Me, "<script>location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</script>")
                'End If
            End If
            'If Request(cst_ccopy) = "1" Then
            '    Common.RespWrite(Me, "<script>location.href='../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</script>")
            'Else
            '    Common.RespWrite(Me, "<script>location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "'</script>")
            'End If
            'Session(cst_CostItemTable) = Nothing
            'Session(Cst_MaterialTable) = Nothing
            'Session(Cst_PersonCostTable) = Nothing
            'Session(Cst_CommonCostTable) = Nothing
            'Session(cst_TrainDescTable) = Nothing
        End If
        'Catch ex As Exception
        '    DbAccess.RollbackTrans(Trans)
        '    If num = 1 Then
        '        If TIMS.ClearSQM(Request("PlanID") = "" Then
        '            Common.MessageBox(Page, "計畫申請失敗!!")
        '        Else
        '            Common.MessageBox(Page, "計畫儲存失敗!!")
        '        End If
        '    Else
        '        Common.MessageBox(Page, "草稿儲存失敗!!")
        '    End If
        '    Throw ex
        'End Try
        ' ViewState("dtTaddress") = Nothing
    End Sub

    '草稿儲存
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim dtTemp As DataTable
        Dim Errmsg As String = ""
        'Errmsg = ""
        If Not Session(cst_Plan_OnClass) Is Nothing Then
            dtTemp = Session(cst_Plan_OnClass)
            For Each drv As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                If Not drv.RowState = DataRowState.Deleted Then
                    If Convert.ToString(drv("Times")) <> "" Then
                        If Convert.ToString(drv("Times")).ToString.Length > 50 Then
                            Errmsg &= "上課時間／時間內容，長度超過限制範圍50文字長度" & vbCrLf
                        End If
                    Else
                        Errmsg &= "上課時間／時間內容，不可為空字串" & vbCrLf
                    End If
                End If
            Next
        End If
        If tNote2.Text <> "" Then
            If Me.tNote2.Text.Length > 1000 Then
                Errmsg &= "其他說明(欄位字數為1000)，超過欄位字數" & vbCrLf
            End If
        End If
        If tNote2b.Text <> "" Then
            If Me.tNote2b.Text.Length > 1000 Then
                Errmsg &= "其他說明(欄位字數為1000)，超過欄位字數" & vbCrLf
            End If
        End If

        '檢核是否有業務權限
        If Not TIMS.Chk_RIDPLAN(RIDValue.Value, sm.UserInfo.PlanID, objconn) Then
            Errmsg &= $"{cst_errmsg22},{RIDValue.Value},{sm.UserInfo.PlanID},{sm.UserInfo.Years}{vbCrLf}" '"登入者無正確的業務權限，不提供儲存服務!!" & vbCrLf
        End If
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>document.getElementById('Button8').style.display="""";Layer_change(5);</script>") 'window.scroll(0,document.body.scrollHeight);
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        If Session("GUID1") = ViewState("GUID1") Then
            'by Milor 20080815--草稿儲存的時候，不應該把session清空，這樣往後按下草稿儲存都沒作用
            'Session("GUID1") = ""
            Call Insert_Plan_Table(2)
            Page.RegisterStartupScript("_onload", "<script language=""javascript"">document.getElementById('Button8').style.display="""";Layer_change(1);</script>")
        End If
    End Sub



    '正式儲存資料確認!!
    Sub CheckAddData(ByRef ErrMsg As String)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        FirstSort.Text = TIMS.ClearSQM(FirstSort.Text)
        'Dim ErrMsg As String = ""
        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    ErrMsg = "於處分日期起的期間，班級申請資料建檔不可正式儲存。"
            End Select
        End If
        If ErrMsg <> "" Then
            '有錯誤訊息
            'Common.MessageBox(Me, ErrMsg)
            Exit Sub 'Return False '不可儲存
        End If

        Dim A_PHour As Integer = 0 '學科總時數(課程大綱)
        Dim T_PHour As Integer = 0 '術科總時數(課程大綱)

        Dim iThours2 As Integer = 0 '訓練小時數
        Dim iALL_PHour As Integer = 0 '總時數(課程大綱)
        Dim iALL_PHour2 As Integer = 0 '總時數(課程大綱)
        Dim rowi As Integer = 0

        Dim cdr4 As DataRow = Nothing '材料費
        Dim cost04 As Integer = 0 '材料費
        'Dim Itemage10 As Integer = 0 '[工作人員費]的計價數量 

        Dim gvid19 As String = TIMS.GetGlobalVar(Me, "19", "1", objconn)
        'sql = "select ItemVar1 from Sys_GlobalVar where TPlanID = '" & sm.UserInfo.TPlanID & "' and DistID = '" & sm.UserInfo.DistID & "' and GVID = 19 "
        'Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        'If dr Is Nothing Then
        'End If

        ErrMsg = ""
        If gvid19 = "" Then
            '(未設定)訓練人數上限
            ErrMsg &= "請至首頁>>系統管理>>系統參數管理>>參數設定-裡設定訓練人數(上限)"
            'Common.MessageBox(Me, ErrMsg)
            Exit Sub
        End If

        TNum.Text = TIMS.ClearSQM(TNum.Text)
        STDate.Text = TIMS.ClearSQM(STDate.Text)
        FDDate.Text = TIMS.ClearSQM(FDDate.Text)
        STDate.Text = TIMS.Cdate3(STDate.Text)
        FDDate.Text = TIMS.Cdate3(FDDate.Text)
        If STDate.Text = "" Then ErrMsg &= "請輸入訓練起日" & vbCrLf
        If FDDate.Text = "" Then ErrMsg &= "請輸入訓練迄日" & vbCrLf
        If TNum.Text = "" Then ErrMsg &= "請輸入訓練人數" & vbCrLf

        If iPYNum >= 3 Then
            'https://jira.turbotech.com.tw/browse/TIMSC-235
            '修改說明：班級申請的正式儲存時，請依據使用者登入的年度，
            '判斷該班級的訓練起日的年度，相同年度，才可儲存，若不是相同年度，則不可儲存。
            If STDate.Text <> "" Then
                Dim STDateYearS As String = CDate(STDate.Text).ToString("yyyy")
                If STDateYearS <> Convert.ToString(sm.UserInfo.Years) Then
                    ErrMsg &= "訓練起日的年度 與 使用者登入的年度，不是相同年度，不可儲存!" & vbCrLf
                End If
            End If
        End If

        If TNum.Text <> "" Then
            '訓練人數上限
            If CInt(TNum.Text) > Val(gvid19) Then
                ErrMsg &= "訓練人數不得大於" & Val(gvid19) & "人" & vbCrLf
            End If
        End If

        Dim v_AppStage As String = TIMS.GetListValue(AppStage)
        If AppStage.SelectedIndex = -1 Then
            ErrMsg &= "請選擇申請階段,申請階段為必填" & vbCrLf
        Else
            If v_AppStage = "" Then ErrMsg &= "請選擇申請階段,申請階段為必填" & vbCrLf
            If v_AppStage = "0" Then ErrMsg &= "請選擇申請階段,申請階段為必填大於0" & vbCrLf
        End If

        '不管什麼都是「年滿15歲以上」。
        'Const cst_ageoDef As Integer = 16 'other Years Start
        txtAge1.Text = TIMS.ClearSQM(txtAge1.Text)
        If Not rdoAge1.Checked AndAlso Not rdoAge2.Checked Then
            ErrMsg &= "請選擇 受訓資格 年齡選項為必選" & vbCrLf
        End If
        If ErrMsg = "" Then
            If rdoAge2.Checked Then
                If txtAge1.Text = "" Then
                    ErrMsg &= "受訓資格 年齡選項2 未輸入有效年齡" & vbCrLf
                Else
                    If Not TIMS.IsNumeric2(txtAge1.Text) Then
                        ErrMsg &= "請檢查 受訓資格 年齡選項2 未輸入有效年齡:" & txtAge1.Text & vbCrLf
                    End If
                    If ErrMsg = "" Then
                        If Val(txtAge1.Text) < cst_AgeOtherDef Then
                            ErrMsg &= "請檢查 受訓資格 年齡選項2 有效年齡(須大於15歲(不含)以上)" & vbCrLf
                        End If
                        If Val(txtAge1.Text) > 99 Then
                            ErrMsg &= "請檢查 受訓資格 年齡選項2 有效年齡(須小於99歲(含)以下)" & vbCrLf
                        End If
                    End If
                End If
            End If
        End If

        If SciPlaceID.SelectedValue <> "" Then
            If Not TIMS.Check_SciPlaceID(ComidValue.Value, SciPlaceID.SelectedValue, objconn) Then
                ErrMsg &= "學科場地已被刪除，請重新選擇" & vbCrLf
            End If
        End If

        If TechPlaceID.SelectedValue <> "" Then
            If Not TIMS.Check_TechPlaceID(ComidValue.Value, TechPlaceID.SelectedValue, objconn) Then
                ErrMsg &= "術科場地已被刪除，請重新選擇" & vbCrLf
            End If
        End If

        If SciPlaceID2.SelectedValue <> "" Then
            If Not TIMS.Check_SciPlaceID(ComidValue.Value, SciPlaceID2.SelectedValue, objconn) Then
                ErrMsg &= "學科場地2已被刪除，請重新選擇" & vbCrLf
            End If
        End If

        If TechPlaceID2.SelectedValue <> "" Then
            If Not TIMS.Check_TechPlaceID(ComidValue.Value, TechPlaceID2.SelectedValue, objconn) Then
                ErrMsg &= "術科場地2已被刪除，請重新選擇" & vbCrLf
            End If
        End If

        If SciPlaceID.SelectedIndex = 0 AndAlso SciPlaceID2.SelectedIndex = 0 And TechPlaceID.SelectedIndex = 0 And TechPlaceID2.SelectedIndex = 0 Then
            ErrMsg &= "【學科場地】、【學科場地2】、【術科場地】、【術科場地2】至少要設定其中一項" & vbCrLf
        End If

        'If TechPlaceID.SelectedIndex = 0 And TechPlaceID2.SelectedIndex = 0 Then
        '    Errmsg &= "【術科場地】與【術科場地2】至少要設定其中一項" & vbCrLf
        'End If

        If Taddress2.SelectedValue = "" AndAlso Taddress3.SelectedValue = "" Then
            ErrMsg &= "【學科上課地址】與【術科上課地址】至少要設定其中一項" & vbCrLf
        End If
        If jobValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
        'cjobValue
        If cjobValue.Value = "" Then ErrMsg &= "請選擇通俗職類，通俗職類為必須選擇" & vbCrLf

        THours.Text = TIMS.ClearSQM(THours.Text)
        If THours.Text <> "" Then
            Dim flagOk As Boolean = True
            Try
                THours.Text = CInt(Val(THours.Text))
            Catch ex As Exception
                ErrMsg &= "訓練時數必須為數字" & vbCrLf
                flagOk = False
            End Try
            If flagOk Then
                iThours2 = Val(THours.Text) '訓練小時數
                If Not IsNumeric(THours.Text) Then
                    ErrMsg &= "訓練時數必須為數字" & vbCrLf
                End If
                If CInt(THours.Text) < 16 Then
                    ErrMsg &= "訓練時數必須大於等於16" & vbCrLf
                End If
            End If
        Else
            iThours2 = 0 '訓練小時數
            ErrMsg &= "訓練時數必須填寫" & vbCrLf
        End If

        'Dim iTPERIOD28 As Integer = 0
        'If TPERIOD28_C1.Checked Then iTPERIOD28 += 1
        'If TPERIOD28_C2.Checked Then iTPERIOD28 += 1
        'If TPERIOD28_C3.Checked Then iTPERIOD28 += 1
        'If iTPERIOD28 = 0 Then
        '    Errmsg &= "授課時段:早上、下午、晚上 至少要設定其中一項" & vbCrLf
        'End If

        If DefGovCost.Text = "0" Then ErrMsg &= "政府補助金額必須大於 0 " & vbCrLf
        If Me.CredPoint.Text <> "" Then
            If Not IsNumeric(Me.CredPoint.Text) Then
                ErrMsg &= "學分數必須為數字" & vbCrLf
            End If
        Else
            Select Case RadioButtonList1.SelectedValue
                Case cst_學分班 ' "Y"
                    'Me.CredPoint.Text =""
                    ErrMsg &= "選擇學分班，學分數為必填數字" & vbCrLf
                Case Else 'cst_非學分班
            End Select
        End If

        Select Case PackageType.SelectedValue
            Case "1" '非包班
            Case "2", "3" '2:企業包班,3:聯合企業包班
            Case Else
                ErrMsg &= "選擇包班種類，包班種類為必填!!" & vbCrLf
                'If PackageType.SelectedIndex = -1 Then '包班種類
                '    ErrMsg &= "選擇包班種類，包班種類為必填!!" & vbCrLf
                'End If
        End Select

        If ErrMsg <> "" Then Exit Sub

        If hTPlanID54.Value = "1" Then   '充電起飛計畫 (hTPlanID54.Value = "1")
            txtUname.Text = TIMS.ClearSQM(txtUname.Text)
            txtIntaxno.Text = TIMS.ClearSQM(txtIntaxno.Text)
            Select Case PackageType.SelectedValue
                Case "2" '充電起飛計畫 '企業包班
                    If Not Session(cst_Plan_BusPackage) Is Nothing Then Session(cst_Plan_BusPackage) = Nothing
                    If txtUname.Text = "" Then
                        txtUname.Text = ""
                        ErrMsg &= "包班事業單位 企業名稱，不可為空" & vbCrLf
                    Else
                        '錯誤檢查
                        'txtUname.Text = Trim(txtUname.Text)
                        If txtUname.Text.ToString.Length > 50 Then
                            ErrMsg &= "包班事業單位 企業名稱，長度超過限制範圍50文字長度" & vbCrLf
                        End If
                    End If
                    If txtIntaxno.Text <> "" Then
                        'txtIntaxno.Text = Trim(txtIntaxno.Text)
                        If Not TIMS.CheckIsECFA(TIMS.ChangeIDNO(txtIntaxno.Text), objconn) Then
                            '未填寫 ECFA包班事業單位資料
                            ErrMsg &= "「" & Convert.ToString(txtUname.Text) & "」該包班事業單位 企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf
                        End If
                    Else
                        txtIntaxno.Text = ""
                        ErrMsg &= "包班事業單位 服務單位統一編號，不可為空" & vbCrLf
                    End If

                Case "3" '充電起飛計畫 '聯合企業包班
                    If Not Session(cst_Plan_BusPackage) Is Nothing Then
                        Dim j As Integer = 0
                        Dim dt As DataTable = Session(cst_Plan_BusPackage)
                        If dt.Rows.Count > 0 Then
                            For i As Integer = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then
                                    Dim dr1 As DataRow = dt.Rows(i)
                                    If Not TIMS.CheckIsECFA(Convert.ToString(dr1("Intaxno")), objconn) Then
                                        '未填寫 包班事業單位資料
                                        ErrMsg &= "「" & Convert.ToString(dr1("Uname")) & "」該包班事業單位 企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf
                                    End If
                                    j += 1
                                End If
                            Next
                            If j = 0 Then
                                '未填寫 包班事業單位資料
                                ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf
                            End If
                        Else
                            '未填寫 包班事業單位資料
                            ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf
                        End If
                    Else
                        '未填寫 包班事業單位資料
                        ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf
                    End If
                Case Else
                    If Not Session(cst_Plan_BusPackage) Is Nothing Then Session(cst_Plan_BusPackage) = Nothing
            End Select

        End If

        Dim vsErrMsg2 As String = ""
        If Session(cst_TrainDescTable) Is Nothing Then
            ErrMsg &= cst_errmsg24 & vbCrLf
        End If
        If Not Session(cst_TrainDescTable) Is Nothing Then
            Dim dt As DataTable = Session(cst_TrainDescTable)
            If dt.Rows.Count > 0 Then
                Dim bolSDateFlag As Boolean = False
                Dim bolEDateFlag As Boolean = False
                rowi = 1
                For i As Int16 = 0 To dt.Rows.Count - 1
                    If Not dt.Rows(i).RowState = DataRowState.Deleted Then
                        iALL_PHour += CInt(dt.Rows(i)("PHour"))
                        If dt.Rows(i)("Classification1") = "1" Then '學科總時數
                            A_PHour += CInt(dt.Rows(i)("PHour"))
                        ElseIf dt.Rows(i)("Classification1") = "2" Then '術科總時數
                            T_PHour += CInt(dt.Rows(i)("PHour"))
                        End If

                        Select Case True
                            Case DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(dt.Rows(i)("STrainDate"))) < 0
                                vsErrMsg2 += "第" & rowi & "筆：課程大網日期不能超過訓練起日" & STDate.Text & vbCrLf
                            Case DateDiff(DateInterval.Day, CDate(dt.Rows(i)("STrainDate")), CDate(FDDate.Text)) < 0
                                vsErrMsg2 += "第" & rowi & "筆：課程大網日期不能超過訓練迄日" & FDDate.Text & vbCrLf
                        End Select

                        '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
                        'Select Case CInt(dt.Rows(i)("Classification1"))
                        '    Case 1, 2
                        '        If IsDBNull(dt.Rows(i)("PTID")) Then
                        '            Me.vsErrMsg2 += "第" & rowi & "筆：請選擇課程大網的上課地點(必填)" & vbCrLf
                        '        End If

                        '    Case Else
                        '        Me.vsErrMsg2 += "第" & rowi & "筆：請選擇課程大網的上課地點(必填)" & vbCrLf
                        'End Select

                        If IsDBNull(dt.Rows(i)("PTID")) Then
                            vsErrMsg2 += "第" & rowi & "筆：請選擇課程大網的上課地點(必填)" & vbCrLf
                        End If

                        If IsDBNull(dt.Rows(i)("TechID")) Then
                            vsErrMsg2 += "第" & rowi & "筆：請選擇課程大網的任課教師(必填)" & vbCrLf
                        End If
                        '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901

                        '(產業人才投資方案) 增加設定課程大綱時,新增內容需有固定二筆資料日期為訓練起日及訓練迄日
                        If CDate(STDate.Text) = CDate(dt.Rows(i)("STrainDate")) Then bolSDateFlag = True
                        If CDate(FDDate.Text) = CDate(dt.Rows(i)("STrainDate")) Then bolEDateFlag = True
                        rowi += 1
                    End If
                    If vsErrMsg2 <> "" Then Exit For
                Next

                If (Not bolSDateFlag Or Not bolEDateFlag) Then
                    vsErrMsg2 += "課程大綱內容資料，需固定二筆資料日期為訓練起日及訓練迄日" & vbCrLf
                End If
            End If

            If iALL_PHour > 0 AndAlso ErrMsg = "" Then
                iALL_PHour2 = 0
                rowi = 1
                Try
                    For Each eItem As DataGridItem In Datagrid3.Items
                        Dim PHourLabel As Label = eItem.FindControl("PHourLabel")
                        If PHourLabel.Text <> "" Then
                            iALL_PHour2 += Val(PHourLabel.Text)
                        End If
                        rowi += 1
                    Next
                Catch ex As Exception
                    ErrMsg &= cst_errmsg25 & vbCrLf
                    ' 課程大綱 有誤
                    'Common.MessageBox(Me, ex.ToString)
                    Exit Sub
                End Try
            End If
        End If

        If Not TIMS.sUtl_ChkTest() Then
            If ErrMsg <> "" Then Exit Sub
            If vsErrMsg2 <> "" Then ErrMsg += vsErrMsg2
            If ErrMsg <> "" Then Exit Sub
        End If

        If rowi = 0 Then
            '97產學訓課程大綱，為必填資料
            'ErrMsg &= "產學訓課程大綱，為必填資料" & vbCrLf
            ErrMsg &= cst_errmsg24 & vbCrLf
        Else
            If iALL_PHour <= 0 Then
                ErrMsg &= "課程大綱的時數，為必填資料，請修正!!" & vbCrLf
            End If
            If ErrMsg = "" AndAlso iALL_PHour2 <= 0 Then
                ErrMsg &= "課程大綱的時數，為必填資料，請修正!!" & vbCrLf
            End If
            If Val(THours.Text) <= 0 Then
                ErrMsg &= "訓練時數，為必填資料，請修正!!" & vbCrLf
            End If
            If ErrMsg = "" Then
                '沒有錯誤整理
                THours.Text = CInt(Val(THours.Text))
                If iALL_PHour <> CInt(Val(THours.Text)) Then
                    '98產業人才投資方案課程大綱時數的加總不得大於訓練時數
                    'If ALL_PHour > CInt(THours.Text) Then
                    '    Errmsg &= "產業人才投資方案課程大綱時數的加總不得大於訓練時數" & vbCrLf
                    'End If
                    '98產業人才投資方案課程大綱時數加總需等於訓練時數
                    ErrMsg &= "產業人才投資方案課程大綱時數加總需等於訓練時數" & vbCrLf
                End If
                If ErrMsg = "" AndAlso iALL_PHour2 <> CInt(Val(THours.Text)) Then
                    ErrMsg &= "產業人才投資方案課程大綱時數加總需等於訓練時數" & vbCrLf
                End If
            End If

        End If

        If Me.STDate.Text <> "" And Me.FDDate.Text <> "" Then
            If IsDate(Me.STDate.Text) And IsDate(Me.FDDate.Text) Then
                'If CInt(DateDiff(DateInterval.DayOfYear, CDate(Me.STDate.Text), CDate(Me.FDDate.Text))) < 77 Then
                '    Errmsg &= "若為學分班，則訓練日期起迄，得大於等於12週" & vbCrLf
                'End If
                If DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FDDate.Text)) < 0 Then
                    ErrMsg &= "訓練日期起迄，訓練迄日需大於訓練起日" & vbCrLf
                End If
                If DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FDDate.Text)) = 0 Then
                    ErrMsg &= "訓練日期起迄，訓練起日不能和訓練迄日同一天" & vbCrLf
                End If
            Else
                ErrMsg &= "訓練日期起迄，格式有誤" & vbCrLf
            End If
        End If
        If Convert.ToString(GCIDValue.Value) = "156" Then
            ErrMsg &= "經費分類代碼有誤(其他 停用),請重新選擇!!" & vbCrLf
        End If
        If Convert.ToString(GCIDValue.Value) = "157" Then
            ErrMsg &= "經費分類代碼有誤(學分班依教育部規定辦理 停用),請重新選擇!!" & vbCrLf
        End If
        If Convert.ToString(GCIDValue.Value) = "158" Then
            ErrMsg &= "經費分類代碼有誤(3C共通核心職能課程 停用),請重新選擇!!" & vbCrLf
        End If

        'If FirstSort.Text <> "" AndAlso STDate.Text <> "" Then
        If ErrMsg = "" AndAlso AppStage.SelectedValue <> "" AndAlso FirstSort.Text <> "" Then
            'https://jira.turbotech.com.tw/browse/TIMSC-138
            '修改說明：班別資料之「優先排序」欄位，
            '如有重複植入之序號者，即無法儲存並跳出出提醒文字。此外，107年度啟用，區隔上、下年度，以6/30做為區隔，開訓日為1/1~6/30為上半年度，7/1~12/31為下半年度，上、下半年的數字要區分，同一個半年度內，不可重複填寫相同數字。
            If ComidValue.Value = "" Then ComidValue.Value = Hid_ComIDNO.Value
            Dim ss As String = ""
            TIMS.SetMyValue(ss, "PlanID", Convert.ToString(sm.UserInfo.PlanID))
            TIMS.SetMyValue(ss, "ComIDNO", Convert.ToString(ComidValue.Value))
            TIMS.SetMyValue(ss, "FirstSort", FirstSort.Text)
            TIMS.SetMyValue(ss, "AppStage", AppStage.SelectedValue)
            'TIMS.SetMyValue(ss, "STDate", STDate.Text)
            Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
            Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
            Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
            Dim PCSVALUE As String = If(rqPlanID <> "" AndAlso rqComIDNO <> "" AndAlso rqSeqNO <> "", String.Concat(rqPlanID, "x", rqComIDNO, "x", rqSeqNO), "")

            TIMS.SetMyValue(ss, "PCSVALUE", PCSVALUE)
            Dim o_OthClassName1 As String = ""
            Dim flagFSort As Boolean = TIMS.Chk_FirstSort1(ss, objconn, o_OthClassName1)
            If flagFSort AndAlso o_OthClassName1 <> "" Then ErrMsg &= String.Concat(cst_errmsg35, "重複：", o_OthClassName1) & vbCrLf

        End If

        Select Case RadioButtonList1.SelectedValue
            Case cst_學分班 ' "Y"
                If PointType.SelectedIndex = -1 Then
                    ErrMsg &= "學分班種類為必填欄位" & vbCrLf
                End If
            Case Else 'cst_非學分班
                '非學分班
                Dim gvid20x1 As String = TIMS.GetGlobalVar(Me, "20", "1", objconn)
                Dim gvid20x2 As String = TIMS.GetGlobalVar(Me, "20", "2", objconn)
                gvid20x1 = TIMS.ClearSQM(gvid20x1)
                gvid20x2 = TIMS.ClearSQM(gvid20x2)

                If gvid20x1 = "" OrElse gvid20x2 = "" Then
                    '(未設定) '訓練時數設定
                    ErrMsg = "請至首頁>>系統管理>>系統參數管理>>參數設定裡設定訓練時數"
                    Exit Sub
                End If
                '訓練時數設定
                If THours.Text <> "" And gvid20x1 <> "" Then
                    If CInt(Val(THours.Text)) > CInt(gvid20x1) Then
                        ErrMsg &= "若為【非學分班】，訓練時數不得大於" & CInt(gvid20x1) & vbCrLf
                    End If
                End If
                If THours.Text <> "" And gvid20x2 <> "" Then
                    If CInt(Val(THours.Text)) < CInt(gvid20x2) Then
                        ErrMsg &= "若為【非學分班】，訓練時數必須大於等於" & CInt(gvid20x2) & vbCrLf
                    End If
                End If

                '---------------------------------------------------------------------
                Dim sqls As String = ""
                Dim drs As DataRow = Nothing
                If jobValue.Value <> "" Then
                    sqls = "select GCID,GCID2 from Key_TrainType where TMID =" & jobValue.Value
                    drs = DbAccess.GetOneRow(sqls, objconn)
                End If
                If drs Is Nothing Then
                    ErrMsg &= "[訓練業別]資料異常,請更正" & vbCrLf
                    Exit Sub
                End If
                Select Case strYears
                    Case "2014"
                        If Convert.ToString(drs("GCID")) <> "" Then
                            Dim sqls2 As String = "select GCode1 from ID_GovClassCast where GCID=" & drs("GCID")
                            Dim drs2 As DataRow = DbAccess.GetOneRow(sqls2, objconn)
                            If Convert.ToString(GCID1Value.Value) <> Convert.ToString(drs2("GCode1")) Then
                                'Errmsg &= "若為【非學分班】，[訓練業別]必須等於[經費分類代碼]的類別" & vbCrLf
                                Dim msgG1 As String = "(G1:" & GCID1Value.Value & "/G2:" & drs2("GCode1") & "/GC:" & drs("GCID") & "/J:" & jobValue.Value & ")"
                                ErrMsg &= "訓練費用編列說明的[經費分類代碼]與[訓練業別]不符,請更正" & msgG1 & vbCrLf
                            End If
                        Else
                            ErrMsg &= "[訓練業別]查無[經費分類代碼]資料異常,請更正" & vbCrLf
                        End If
                    Case "2015"
                        If Convert.ToString(drs("GCID2")) <> "" Then
                            Dim sqls2 As String = "SELECT GCODE1 FROM V_GOVCLASSCAST2 WHERE GCID2=" & drs("GCID2")
                            Dim drs2 As DataRow = DbAccess.GetOneRow(sqls2, objconn)
                            If Convert.ToString(GCID1Value.Value) <> Convert.ToString(drs2("GCODE1")) Then
                                'Errmsg &= "若為【非學分班】，[訓練業別]必須等於[經費分類代碼]的類別" & vbCrLf
                                Dim msgG1 As String = "(G1:" & GCID1Value.Value & "/G2:" & drs2("GCODE1") & "/GC:" & drs("GCID2") & "/J:" & jobValue.Value & ")"
                                ErrMsg &= "訓練費用編列說明的[經費分類代碼2]與[訓練業別]不符,請更正" & msgG1 & vbCrLf
                            End If
                        Else
                            ErrMsg &= "[訓練業別]查無[經費分類代碼2]資料異常,請更正" & vbCrLf
                        End If
                End Select
                '----------------------------------------------------------------------------------------------------

                'ItemVar1(, ItemVar2)
                If Me.STDate.Text <> "" And Me.FDDate.Text <> "" Then
                    If IsDate(Me.STDate.Text) And IsDate(Me.FDDate.Text) Then
                        Dim tempDate As Date = DateAdd(DateInterval.Month, 4, CDate(STDate.Text))
                        If DateDiff(DateInterval.Day, tempDate, CDate(FDDate.Text)) > 0 Then
                            ErrMsg &= "若為【非學分班】，訓練起迄日期區間，不得超過4個月" & vbCrLf
                        End If
                        'If CDate(Me.FDDate.Text) >= tempDate Then
                        '    ErrMsg &= "若為【非學分班】，訓練起迄日期區間，不得超過4個月" & vbCrLf
                        'End If
                    End If
                End If

                Dim HaveCostID04 As Boolean = False
                Dim HaveMaterial As Boolean = False
                HaveCostID04 = False '查詢是否有材料費項目
                HaveMaterial = False '是否有材料品名項目表
                If Not Session(cst_CostItemTable) Is Nothing Then
                    Dim dt2 As DataTable
                    dt2 = Session(cst_CostItemTable)
                    For i As Int16 = 0 To dt2.Rows.Count - 1
                        If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                            AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                            cdr4 = dt2.Select("CostID='04'")(0)
                            cost04 = cdr4("Itemage") * cdr4("OPrice")
                            HaveCostID04 = True
                            Exit For
                        End If
                    Next
                    '費用檢核。
                    Call CheckCostItemTable(iThours2, T_PHour, A_PHour, dt2, ErrMsg)
                End If
                If Not Session(Cst_MaterialTable) Is Nothing Then
                    Dim dt2 As DataTable
                    dt2 = Session(Cst_MaterialTable)
                    For i As Int16 = 0 To dt2.Rows.Count - 1
                        If Not dt2.Rows(i).RowState = DataRowState.Deleted Then  '已刪除者不可做更動
                            HaveMaterial = True
                            Exit For
                        End If
                    Next
                End If
                If Not HaveCostID04 AndAlso HaveMaterial Then
                    ErrMsg &= "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
                End If
        End Select
        If ErrMsg <> "" Then Exit Sub

        Select Case RadioButtonList1.SelectedValue
            Case cst_學分班 ' "Y"
                'Dim Flag4 As Boolean = TIMS.Chk_STFDateMn4(CDate(STDate.Text), CDate(FDDate.Text), 4)
                'If Not Flag4 Then
                '    ErrMsg &= "學分班:訓練起日及訓練迄日間隔不得超過4個月" & vbCrLf
                'End If
                Dim Flag82 As Boolean = TIMS.Chk_STFDateMn82(CDate(STDate.Text), CDate(FDDate.Text))
                If Not Flag82 Then
                    ErrMsg &= "學分班:訓練迄日不得設定超過8月份(上半年)及2月份(下半年)" & vbCrLf
                End If
            Case Else 'cst_非學分班
                Dim Flag82 As Boolean = TIMS.Chk_STFDateMn82(CDate(STDate.Text), CDate(FDDate.Text))
                If Not Flag82 Then
                    ErrMsg &= "非學分班:訓練迄日不得設定超過8月份(上半年)及2月份(下半年)" & vbCrLf
                End If
        End Select
        If ErrMsg <> "" Then Exit Sub

        '---------------檢查所選的學術科場地是否有被選中--------------------------
        Dim dtTaddress As DataTable = CType(ViewState("dtTaddress"), DataTable)
        Dim TrainDescTable As DataTable = CType(Me.Session(cst_TrainDescTable), DataTable)
        If TrainDescTable Is Nothing Then
            ErrMsg &= cst_errmsg24 & vbCrLf
            'Page.RegisterStartupScript("_onload", "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";Layer_change(5);</script>")
            'Common.MessageBox(Me, ErrMsg)
            Exit Sub
        End If

        'If Not TrainDescTable Is Nothing Then
        'End If

        Try
            For z As Integer = 0 To TrainDescTable.Rows.Count - 1
                If Not TrainDescTable.Rows(z).RowState = DataRowState.Deleted Then
                    Dim mach As Integer = 0
                    Dim tdbdr As DataRow = TrainDescTable.Rows(z)
                    For x As Integer = 0 To dtTaddress.Rows.Count - 1
                        If Not dtTaddress.Rows(x).RowState = DataRowState.Deleted Then
                            Dim dtsdr As DataRow = dtTaddress.Rows(x)
                            If Convert.ToString(tdbdr("PTID")) = Convert.ToString(dtsdr("PTID")) Then
                                mach = 1
                            End If
                        End If
                    Next
                    If mach = 0 Then
                        ErrMsg &= "[課程大綱]的[上課地點]不在所選的[學術科場地]範圍內,請修改!!" & vbCrLf
                    End If
                    Exit For
                End If
            Next
        Catch ex As Exception
            ErrMsg &= "[課程大綱]的[上課地點]不在所選的[學術科場地]範圍內,請修改!!" & vbCrLf
        End Try
        '---------------end 檢查所選的學術科場地是否有被選中--------------------------

        '材料費用總計 檢核 與材料費用不相符
        If CInt(Val(labTotal67.Text)) > 0 OrElse cost04 > 0 Then
            If cost04 <> CInt(Val(labTotal67.Text)) Then
                ErrMsg &= "材料費用總計(" & labTotal67.Text & ") 與 材料費項目金額(" & cost04 & ") 等 檢核不相等,請修改!!" & vbCrLf
            End If
        End If
        If Me.tNote2.Text <> "" Then
            If Me.tNote2.Text.Length > 1000 Then
                ErrMsg &= "其他說明(欄位字數為1000)，超過欄位字數" & vbCrLf
            End If
        End If
        If tNote2b.Text <> "" Then
            If Me.tNote2b.Text.Length > 1000 Then
                ErrMsg &= "其他說明(欄位字數為1000)，超過欄位字數" & vbCrLf
            End If
        End If

        '(檢核)結束-----
        If ErrMsg <> "" Then
            'Page.RegisterStartupScript("_onload", "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";Layer_change(5);</script>")
            'Common.MessageBox(Me, ErrMsg)
            Exit Sub
        End If
        '(檢核)結束-----

        ''儲存點
        'If Session("GUID1") =  ViewState("GUID1") Then
        '    'Session("GUID1") = ""
        '    Call Insert_Plan_Table(1) '若儲存成功，則下列不執行，直接跳頁 ../01/TC_01_014_add.aspx
        '     ViewState("dtTaddress") = Nothing
        '    Page.RegisterStartupScript("_onload", "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";</script>")
        'End If

    End Sub

    '正式儲存 (檢核)
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim ErrMsg As String = ""
        Call CheckAddData(ErrMsg)
        If ErrMsg <> "" Then
            '有錯誤訊息
            Page.RegisterStartupScript("_onload", "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";Layer_change(5);</script>")
            Common.MessageBox(Me, ErrMsg)
            Exit Sub 'Return False '不可儲存
        End If

        '檢核是否有業務權限
        If Not TIMS.Chk_RIDPLAN(RIDValue.Value, sm.UserInfo.PlanID, objconn) Then
            ErrMsg &= $"{cst_errmsg22},{RIDValue.Value},{sm.UserInfo.PlanID},{sm.UserInfo.Years}{vbCrLf}" '"登入者無正確的業務權限，不提供儲存服務!!" & vbCrLf
        End If
        If ErrMsg <> "" Then
            '有錯誤訊息
            Page.RegisterStartupScript("_onload", "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";Layer_change(5);</script>")
            Common.MessageBox(Me, ErrMsg)
            Exit Sub 'Return False '不可儲存
        End If

        '儲存點
        If Session("GUID1") = ViewState("GUID1") Then
            'Session("GUID1") = ""
            Call Insert_Plan_Table(1) '若儲存成功，則下列不執行，直接跳頁 ../01/TC_01_014_add.aspx
            ViewState("dtTaddress") = Nothing
            Page.RegisterStartupScript("_onload", "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";</script>")
        End If

    End Sub

    '費用檢核。
    Sub CheckCostItemTable(ByVal Thours2 As Integer, ByVal T_PHour As Integer, ByVal A_PHour As Integer,
        ByRef dt2 As DataTable, ByRef ErrMsg As String)

        'Dim Thours2 As Integer = 0
        'Dim dt2 As DataTable
        Dim cdr1 As DataRow = Nothing
        Dim cdr2 As DataRow = Nothing
        Dim cdr12 As DataRow = Nothing
        Dim cdr13 As DataRow = Nothing

        Dim cdr3 As DataRow = Nothing
        Dim cdr4 As DataRow = Nothing
        Dim cdr5 As DataRow = Nothing
        Dim cdr6 As DataRow = Nothing
        Dim cdr7 As DataRow = Nothing
        Dim cdr8 As DataRow = Nothing
        Dim cdr9 As DataRow = Nothing
        Dim cdr10 As DataRow = Nothing
        Dim cdr11 As DataRow = Nothing
        Dim cdr14 As DataRow = Nothing '雜支
        Dim cdr15 As DataRow = Nothing '(補充保險費)'二代健保補充保險費

        Dim cost01 As Integer = 0 '學科_鐘點費(外聘)
        Dim cost02 As Integer = 0 '術科_鐘點費(外聘)
        Dim cost12 As Integer = 0 '學科_鐘點費(內聘)
        Dim cost13 As Integer = 0 '術科_鐘點費(內聘)
        Dim cost14 As Integer = 0 '雜支
        Dim cost15 As Integer = 0 '(補充保險費)'二代健保補充保險費

        Dim cost03 As Integer = 0 '教材費
        Dim cost04 As Integer = 0 '材料費
        Dim cost05 As Integer = 0 '場地費
        Dim cost06 As Integer = 0 '宣導費

        Dim cost07 As Integer = 0 '行政管理費
        Dim cost08 As Integer = 0 '保險費
        Dim cost09 As Integer = 0 '術科助教費用
        Dim cost10 As Integer = 0 '工作人員費
        Dim cost11 As Integer = 0 '其他費用

        Dim ItemageT4 As Integer = 0 '學科-鐘點費(內、外聘)與術科-鐘點費(內、外聘)的計價數量總合
        Dim ItemageA As Integer = 0 '學科-鐘點費(內、外聘)的計價數量總合
        Dim ItemageT As Integer = 0 '術科-鐘點費(內、外聘)的計價數量總合

        'Dim A_PHour As Integer = 0 '學科總時數(課程大綱)
        'Dim T_PHour As Integer = 0 '術科總時數(課程大綱)

        '取得Basic 訓練費用項目 
        If dt2.Select("CostID = '01'").Length <> 0 Then '學科_鐘點費(外聘)
            cdr1 = dt2.Select("CostID = '01'")(0)
            cost01 = cdr1("Itemage") * cdr1("OPrice")
        End If
        If dt2.Select("CostID = '02'").Length <> 0 Then '術科_鐘點費(外聘)
            cdr2 = dt2.Select("CostID = '02'")(0)
            cost02 = cdr2("Itemage") * cdr2("OPrice")
        End If
        If dt2.Select("CostID = '12'").Length <> 0 Then '學科_鐘點費(內聘)
            cdr12 = dt2.Select("CostID = '12'")(0)
            cost12 = cdr12("Itemage") * cdr12("OPrice")
        End If
        If dt2.Select("CostID = '13'").Length <> 0 Then '術科_鐘點費(內聘)
            cdr13 = dt2.Select("CostID = '13'")(0)
            cost13 = cdr13("Itemage") * cdr13("OPrice")
        End If
        If dt2.Select("CostID = '03'").Length <> 0 Then '教材費
            cdr3 = dt2.Select("CostID = '03'")(0)
            cost03 = cdr3("Itemage") * cdr3("OPrice")
        End If
        If dt2.Select("CostID = '04'").Length <> 0 Then '材料費
            cdr4 = dt2.Select("CostID = '04'")(0)
            cost04 = cdr4("Itemage") * cdr4("OPrice")
        End If
        If dt2.Select("CostID = '05'").Length <> 0 Then '場地費
            cdr5 = dt2.Select("CostID = '05'")(0)
            cost05 = cdr5("Itemage") * cdr5("OPrice")
        End If
        If dt2.Select("CostID = '06'").Length <> 0 Then '宣導費
            cdr6 = dt2.Select("CostID = '06'")(0)
            cost06 = cdr6("Itemage") * cdr6("OPrice")
        End If
        If dt2.Select("CostID = '07'").Length <> 0 Then '行政管理費
            cdr7 = dt2.Select("CostID = '07'")(0)
            cost07 = cdr7("Itemage") * cdr7("OPrice")
        End If
        If dt2.Select("CostID = '08'").Length <> 0 Then '保險費
            cdr8 = dt2.Select("CostID = '08'")(0)
            cost08 = cdr8("Itemage") * cdr8("OPrice")
        End If
        If dt2.Select("CostID = '09'").Length <> 0 Then '術科助教費用
            cdr9 = dt2.Select("CostID = '09'")(0)
            cost09 = cdr9("Itemage") * cdr9("OPrice")
        End If
        If dt2.Select("CostID = '10'").Length <> 0 Then '工作人員費
            cdr10 = dt2.Select("CostID = '10'")(0)
            cost10 = cdr10("Itemage") * cdr10("OPrice")
        End If
        If dt2.Select("CostID = '11'").Length <> 0 Then '其他費用
            cdr11 = dt2.Select("CostID = '11'")(0)
            cost11 = cdr11("Itemage") * cdr11("OPrice")
        End If
        If dt2.Select("CostID = '14'").Length <> 0 Then '雜支
            cdr14 = dt2.Select("CostID = '14'")(0)
            cost14 = cdr14("Itemage") * cdr14("OPrice")
        End If
        If dt2.Select("CostID = '15'").Length <> 0 Then '(補充保險費)'二代健保補充保險費
            cdr15 = dt2.Select("CostID = '15'")(0)
            cost15 = cdr15("Itemage") * cdr15("OPrice")
        End If

        '判斷條件如下
        'If Not cdr9 Is Nothing And Not cdr2 Is Nothing Then '如果有輸入術科助教費用和術科-鐘點費
        '    If cdr9("Itemage") > cdr2("Itemage") Then  '訓練費用的"術科助教費用"的計價數量不得大於"術科-鐘點費"的計價數量.
        '        Errmsg &= "[術科助教費用]的計價數量不得大於[術科-鐘點費]的計價數量 " & vbCrLf
        '    End If
        'End If
        If Not cdr9 Is Nothing Then   '2010改
            If cdr9("Itemage") > T_PHour Then
                ErrMsg &= "[術科助教費用]的計價數量不得大於[課程大綱]的[術科時數]加總!!" & vbCrLf
            End If
        End If

        If Not cdr3 Is Nothing Then '教材費
            If cdr3("Itemage") > TNum.Text Then
                ErrMsg &= "[教材費]的計價數量不得大於[班別資料]的[訓練人數] " & vbCrLf
            End If
            'If cdr3("Itemage") <> TNum.Text Then
            '    Errmsg &= "[教材費]的計價數量須等於[班別資料]的[訓練人數] " & vbCrLf
            'End If
        End If
        If Not cdr4 Is Nothing Then '材料費
            If cdr4("Itemage") > TNum.Text Then
                ErrMsg &= "[材料費]的計價數量不得大於[班別資料]的[訓練人數] " & vbCrLf
            End If
            'If cdr4("Itemage") <> TNum.Text Then
            '    Errmsg &= "[材料費]的計價數量須等於[班別資料]的[訓練人數] " & vbCrLf
            'End If
        End If
        If Not cdr6 Is Nothing Then '宣導費
            If cost06 > 10000 Then
                ErrMsg &= "[宣導費]的小計(單價 * 計價數量)不得大於10,000元 " & vbCrLf
            End If
        End If
        If Not cdr10 Is Nothing Then '工作人員費
            If cdr10("Itemage") > Thours2 Then
                ErrMsg &= "[工作人員費]的計價數量不得大於[班別資料]的[訓練時數] " & vbCrLf
            End If
        End If
        If Not cdr7 Is Nothing Then ' 行政管理費
            'If cost07 > ((cost01 + cost02 + cost12 + cost13 + cost03 + cost04 + cost05 + cost06 + cost08 + cost09 + cost10 + cost11 + cost15) * 9) / 100 Then    '原來為15% 2010/05/27 改為 9%
            If cost07 > ((cost01 + cost02 + cost12 + cost13 + cost03 + cost06 + cost04 + cost05 + cost09 + cost10 + cost08 + cost15 + cost11) * 9) / 100 Then    '原來為15% 2010/05/27 改為 9%
                ErrMsg &= "[行政管理費]的小計不得大於(學科-鐘點費(內、外聘)+術科-鐘點費(內、外聘)+教材費+宣導費+材料費+場地費+術科助教費用+工作人員費+保險費+補充保險費+其他費用)的小計總額的9% " & vbCrLf
            End If
        End If
        If Not cdr14 Is Nothing Then ' 雜支
            '(十一)	雜支：以實際辦訓費用總額（不包括出席費、鐘點費、稿費、差旅費、工作人員服務費及、管理費及'(補充保險費)二代健保補充保險費）之百分之五以內編列。
            If cost14 > ((cost03 + cost04 + cost05 + cost06 + cost08 + cost11) * 5) / 100 Then
                ErrMsg &= "[雜支]的小計不得大於(教材費+宣導費+材料費+場地費+保險費用+其他費用)的小計總額的5% " & vbCrLf
            End If
        End If
        If Not cdr15 Is Nothing Then '(補充保險費)二代健保補充保險費
            If cost15 > ((cost01 + cost02 + cost12 + cost13 + cost09 + cost10) * 2) / 100 Then
                ErrMsg &= "[補充保險費]的小計不得大於(學科-鐘點費(內、外聘)+術科-鐘點費(內、外聘)+術科助教費用+工作人員費)的小計總額的2% "
                ErrMsg &= "，補充保險費:" & cost15 & "、小計總額的2%:" & (((cost01 + cost02 + cost12 + cost13 + cost09 + cost10) * 2) / 100) & vbCrLf
            End If
        End If

        If Not cdr1 Is Nothing Then
            '學科_鐘點費(外聘)
            ItemageT4 += cdr1("Itemage")
            ItemageA += cdr1("Itemage")
        End If
        If Not cdr2 Is Nothing Then
            '術科_鐘點費(外聘)
            ItemageT4 += cdr2("Itemage")
            ItemageT += cdr2("Itemage")
        End If

        If Not cdr12 Is Nothing Then
            '學科_鐘點費(內聘)
            ItemageT4 += cdr12("Itemage")
            ItemageA += cdr12("Itemage")   'ItemageA學科-鐘點費(內、外聘)的計價數量總合
        End If

        If Not cdr13 Is Nothing Then
            '術科_鐘點費(內聘)
            ItemageT4 += cdr13("Itemage") 'ItemageT4學科-鐘點費(內、外聘)與術科-鐘點費(內、外聘)的計價數量總合
            ItemageT += cdr13("Itemage")  'ItemageT術科-鐘點費(內、外聘)的計價數量總合
        End If

        '學科-鐘點費(內、外聘)與術科-鐘點費(內、外聘)的計價數量總合
        If ItemageT4 > Thours2 Then
            ErrMsg &= "[學科-鐘點費(內、外聘)]與[術科-鐘點費(內、外聘)]的計價數量總合,不得大於[班別資料]的[訓練時數] " & vbCrLf
        End If
        If ItemageA > A_PHour Then
            ErrMsg &= "[學科-鐘點費(內、外聘)]的計價數量總合,不得大於[課程大綱]的[學科時數]加總!! " & vbCrLf
        End If
        If ItemageT > T_PHour Then
            ErrMsg &= "[術科-鐘點費(內、外聘)]的計價數量總合,不得大於[課程大綱]的[術科時數]加總!! " & vbCrLf
        End If

        If Not cdr9 Is Nothing Then '如果有輸入術科助教費用和術科-鐘點費
            If cdr9("Itemage") > ItemageT Then  '訓練費用的"術科助教費用"的計價數量不得大於"術科-鐘點費"的計價數量.
                ErrMsg &= "[術科助教費用]的計價數量不得大於[術科(內、外聘)-鐘點費]的計價數量 " & vbCrLf
            End If
        End If
    End Sub

    '鎖定輸入項。
    Sub Disabled_Items(ByVal sTitle As String)
        Me.EMail.ReadOnly = True
        Me.trainValue.Disabled = True
        Me.cjobValue.Disabled = True
        Me.jobValue.Disabled = True
        Me.PlanCause.ReadOnly = True
        Me.PurScience.ReadOnly = True
        Me.PurTech.ReadOnly = True
        Me.PurMoral.ReadOnly = True
        Me.Degree.Enabled = False
        'Me.Age_l.ReadOnly = True
        'Me.Age_u.ReadOnly = True
        'Me.Sex.Enabled = False
        'Me.Solder.Enabled = False
        Me.Other1.ReadOnly = True
        Me.Other2.ReadOnly = True
        Me.Other3.ReadOnly = True
        'Me.Other1.Enabled = False
        'Me.Other2.Enabled = False
        'Me.Other3.Enabled = False
        TIMS.Tooltip(Other1, Cst_msgother1)
        TIMS.Tooltip(Other2, Cst_msgother1)
        TIMS.Tooltip(Other3, Cst_msgother1)

        Me.TMScience.ReadOnly = True
        Me.SciHours.ReadOnly = True
        Me.GenSciHours.ReadOnly = True
        Me.ProSciHours.ReadOnly = True
        Me.ProTechHours.ReadOnly = True
        Me.TotalHours.ReadOnly = True
        Me.ClassName.ReadOnly = True
        Me.TNum.ReadOnly = True
        'Me.THours.ReadOnly = True
        Me.STDate.ReadOnly = True
        Me.FDDate.ReadOnly = True
        Me.CyclType.ReadOnly = True
        Me.CustomValidator4.Enabled = False
        Me.ClassCount.ReadOnly = True

        Me.DefGovCost.ReadOnly = True
        Me.DefUnitCost.ReadOnly = True
        Me.DefStdCost.ReadOnly = True
        TIMS.Tooltip(DefGovCost, sTitle)
        TIMS.Tooltip(DefUnitCost, sTitle)
        TIMS.Tooltip(DefStdCost, sTitle)

        Me.Note.ReadOnly = True
        Me.CredPoint.ReadOnly = True
        Me.RoomName.ReadOnly = True
        Me.FactMode.Enabled = False
        Me.FactModeOther.ReadOnly = True
        Me.ConNum.ReadOnly = True
        Me.ContactName.ReadOnly = True
        Me.ContactPhone.ReadOnly = True
        Me.ContactEmail.ReadOnly = True
        Me.ContactFax.ReadOnly = True
        ClassCate.Enabled = False
        Me.Content.ReadOnly = True
        '授課時段:'早上'下午'晚上
        'TPERIOD28_C1.Enabled = False
        'TPERIOD28_C2.Enabled = False
        'TPERIOD28_C3.Enabled = False

        Button29.Enabled = False
        btnAddBusPackage.Enabled = False
        center.Enabled = False
        Org.Disabled = True
        Button8.Visible = False
        btnAdd.Visible = False
        Button9.Enabled = False
        btnAddMaterial.Enabled = False
        btu_sel.Disabled = True
        'Button27.Disabled = True
    End Sub

    '新增 '計價數量 項目金額 含檢核
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim dt As DataTable
        Dim dr As DataRow
        Dim Errmsg As String = ""
        Dim CostID2Value As String = CostID2.SelectedValue.Split(",")(0)

        If Session(cst_CostItemTable) Is Nothing Then
            Call CreateCostItem()
        End If
        dt = Session(cst_CostItemTable)
        dt.Columns("PCID").AutoIncrement = True
        dt.Columns("PCID").AutoIncrementSeed = -1
        dt.Columns("PCID").AutoIncrementStep = -1

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If dt.Rows(i).Item("CostID").ToString = CostID2Value Then
                        Errmsg &= "此項目已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If

        If CostID2.SelectedValue = "" Then
            Errmsg &= "請選擇項目" & vbCrLf
        End If

        OPrice2.Text = TIMS.ClearSQM(OPrice2.Text)
        Dim dOPrice2 As Double = TIMS.VAL1(OPrice2.Text)
        If OPrice2.Text = "" OrElse dOPrice2 <= 0 Then
            Errmsg &= "單價請填數字須大於0,不可為0" & vbCrLf
        Else
            'If Not IsNumeric(OPrice2.Text) Then
            '    Errmsg &= "單價請填數字格式" & vbCrLf
            'End If
            If Not TIMS.IsNumeric2(OPrice2.Text) Then
                Errmsg &= "單價請填數字格式(正整數)須大於0,不可為0" & vbCrLf
            End If
        End If

        Itemage.Value = TIMS.ClearSQM(Itemage.Value)
        Dim iItemage As Double = TIMS.VAL1(Itemage.Value)
        If Itemage.Value = "" OrElse iItemage <= 0 Then
            Errmsg &= "計價數量 請填數字須大於0,不可為0" & vbCrLf
        Else
            If Not TIMS.IsNumeric2(Itemage.Value) Then
                Errmsg &= "計價數量 請填數字格式(正整數)須大於0,不可為0" & vbCrLf
            End If
        End If

        If Errmsg = "" AndAlso RadioButtonList1.SelectedValue = "N" Then  '非學分班
            If CostID2Value = "01" And OPrice2.Text > 1600 Then
                Errmsg &= "學科-鐘點費(外聘)的單價不得大於1,600元" & vbCrLf
            End If
            If CostID2Value = "12" And OPrice2.Text > 800 Then
                Errmsg &= "學科-鐘點費(內聘)的單價不得大於800元" & vbCrLf
            End If
            If CostID2Value = "02" And OPrice2.Text > 1600 Then
                Errmsg &= "術科-鐘點費(外聘)的單價不得大於1,600元" & vbCrLf
            End If
            If CostID2Value = "13" And OPrice2.Text > 800 Then
                Errmsg &= "術科-鐘點費(內聘)的單價不得大於800元" & vbCrLf
            End If
            If CostID2Value = "03" And OPrice2.Text > 800 Then
                Errmsg &= "教材費的單價不得大於800元" & vbCrLf
            End If
            If CostID2Value = "05" And OPrice2.Text > 2500 Then
                Errmsg &= "場地費的單價不得大於2,500元" & vbCrLf
            End If
            If CostID2Value = "06" And OPrice2.Text > 10000 Then
                Errmsg &= "宣導費的單價不得大於10,000元" & vbCrLf
            End If
            If CostID2Value = "06" And Itemage.Value <> 1 Then
                Errmsg &= "宣導費的計價數量須等於1" & vbCrLf
            End If
            If CostID2Value = "09" And OPrice2.Text > 400 Then
                Errmsg &= "術科助教費用的單價不得大於400元" & vbCrLf
            End If
            If CostID2Value = "10" And OPrice2.Text > 160 Then
                Errmsg &= "工作人員費的單價不得大於160元" & vbCrLf
            End If
            'If CostID2Value = "14" Then
            '    Common.MessageBox(Me, "【雜支】：" & vbCrLf & vbCrLf & Space(20) & "以實際辦訓費用總額" & vbCrLf & vbCrLf & Space(20) & "  (不包括出席費、鐘點費、稿費、差旅費、工作人員服務費及管理費)" & vbCrLf & vbCrLf & Space(20) & "  之百分之五以內編列。")
            'End If
        End If

        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("CostMode") = 5 '產業人才投資方案專用

        'dr("CostID") = CostID2.SelectedValue '項目
        dr("CostID") = CostID2Value '項目
        dr("OPrice") = TIMS.ChangeIDNO(OPrice2.Text) '單價
        dr("Itemage") = TIMS.ChangeIDNO(Itemage.Value) '計價數量

        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(cst_CostItemTable) = dt
        Call CreateCostItem()

        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    '回上一頁
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        If ViewState("search") <> "" Then
            Session("search") = ViewState("search")
        End If

        Dim url1 As String = ""
        If TIMS.ClearSQM(Request("todo")) = 1 Then
            url1 = "../04/TC_04_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        ElseIf Request(cst_ccopy) = "1" Then
            url1 = "../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        Else
            url1 = "../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        End If
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '機構資訊(隱藏)
    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Dim sql As String = ""
        Dim dr As DataRow
        sql = ""
        sql &= " SELECT b.ComIDNO,c.ContactEmail,c.ZipCode,c.Address,c.ContactName"
        sql &= " ,c.Phone,c.ContactEmail,c.ContactFax"
        sql &= " FROM Auth_Relship a"
        sql &= " JOIN Org_OrgInfo b On a.OrgID=b.OrgID"
        sql &= " JOIN Org_OrgPlanInfo c ON a.RSID=c.RSID"
        sql &= " WHERE a.RID='" & RIDValue.Value & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            'Common.MessageBox(Me, "程式出現例外狀況，請聯絡東柏人員!")
            Common.MessageBox(Me, cst_errmsg1)
            Page.RegisterStartupScript("Londing", "<script>Layer_change('');</script>")
            Exit Sub
        End If

        ComidValue.Value = Convert.ToString(dr("ComIDNO"))
        If Table1_Email.Visible = True Then
            EMail.Text = Convert.ToString(dr("ContactEmail"))
        End If

        Me.ContactName.Text = dr("ContactName").ToString
        Me.ContactPhone.Text = dr("Phone").ToString
        Me.ContactEmail.Text = dr("ContactEmail").ToString
        Me.ContactFax.Text = dr("ContactFax").ToString

        SciPlaceID = TIMS.Get_SciPlaceID(SciPlaceID, ComidValue.Value, 1, "", objconn)
        TechPlaceID = TIMS.Get_TechPlaceID(TechPlaceID, ComidValue.Value, 1, "", objconn)
        SciPlaceID2 = TIMS.Get_SciPlaceID(SciPlaceID2, ComidValue.Value, 1, "", objconn)
        TechPlaceID2 = TIMS.Get_TechPlaceID(TechPlaceID2, ComidValue.Value, 1, "", objconn)
        Page.RegisterStartupScript("Londing", "<script>Layer_change('');</script>")
    End Sub

    '新增 上課時間／內容
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim Errmsg As String = ""

        Times.Text = TIMS.ClearSQM(Times.Text)
        If Times.Text <> "" Then
            'Times.Text = Trim(Times.Text)
            If Times.Text.ToString.Length > 50 Then
                Errmsg &= "上課時間／時間內容，長度超過限制範圍50文字長度" & vbCrLf
            End If
        Else
            Times.Text = ""
            Errmsg &= "上課時間／時間內容，不可為空字串" & vbCrLf
        End If

        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        'Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow

        If Session(cst_Plan_OnClass) Is Nothing Then
            Call CreateClassTime()
        End If
        dt = Session(cst_Plan_OnClass)
        dt.Columns("POCID").AutoIncrement = True
        dt.Columns("POCID").AutoIncrementSeed = -1
        dt.Columns("POCID").AutoIncrementStep = -1

        'If Session(cst_Plan_OnClass) Is Nothing Then
        '    sql = "SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & TIMS.ClearSQM(Request("PlanID") & "' and ComIDNO='" & TIMS.ClearSQM(Request("ComIDNO") & "' and SeqNo='" & TIMS.ClearSQM(Request("SeqNo") & "'"
        '    dt = DbAccess.GetDataTable(sql)
        'Else
        '    dt = Session(cst_Plan_OnClass)
        'End If

        dr = dt.NewRow
        dt.Rows.Add(dr)
        If TIMS.ClearSQM(Request("PlanID")) <> "" Then
            dr("PlanID") = TIMS.ClearSQM(Request("PlanID"))
            dr("ComIDNO") = TIMS.ClearSQM(Request("ComIDNO"))
            dr("SeqNo") = TIMS.ClearSQM(Request("SeqNo"))
        End If
        dr("Weeks") = Weeks.SelectedValue
        dr("Times") = Times.Text
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        'DataGrid1Table.Visible = True
        'DataGrid1.DataSource = dt
        'DataGrid1.DataBind()
        Session(cst_Plan_OnClass) = dt
        Call CreateClassTime()

        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    '再次檢核 (server check)
    Function Chk_TrainDescInput(ByRef Errmsg As String, ByRef htTrainDesc As Hashtable, ByRef dt As DataTable) As Boolean
        Dim rst As Boolean = True
        Dim sTPERIOD28 As String = htTrainDesc("sTPERIOD28")
        Dim sSTDate As String = htTrainDesc("sSTDate")
        Dim sFDDate As String = htTrainDesc("sFDDate")
        Dim sPName As String = htTrainDesc("sPName")
        Dim sddlpnH1 As String = htTrainDesc("sddlpnH1")
        Dim sddlpnM1 As String = htTrainDesc("sddlpnM1")
        Dim sddlpnH2 As String = htTrainDesc("sddlpnH2")
        Dim sddlpnM2 As String = htTrainDesc("sddlpnM2")
        Dim sPHour As String = htTrainDesc("sPHour")
        Dim sPCont As String = htTrainDesc("sPCont")
        Dim sSTrainDate As String = htTrainDesc("sSTrainDate")
        Dim sClassification1 As String = htTrainDesc("sClassification1")
        Dim sComidValue As String = htTrainDesc("sComidValue")
        Dim sPTID1 As String = htTrainDesc("sPTID1")
        Dim sPTID2 As String = htTrainDesc("sPTID2")
        Dim sOLessonTeah1Value As String = htTrainDesc("sOLessonTeah1Value")
        Dim sOLessonTeah2Value As String = htTrainDesc("sOLessonTeah2Value")
        Dim sPTDID As String = htTrainDesc("sPTDID")

        'Dim Errmsg As String = ""
        Errmsg = ""
        Const cst_NNN As String = "NNN"
        If sTPERIOD28 = cst_NNN Then
            Errmsg &= "授課時段:早上、下午、晚上 至少要設定其中一項" & vbCrLf
        End If

        If sSTDate = "" Then
            Errmsg &= "請先輸入訓練起日" & vbCrLf
        Else
            If Not TIMS.IsDate1(sSTDate) Then
                Errmsg &= "訓練起日請填日期格式" & vbCrLf
            End If
            sSTDate = TIMS.Cdate3(sSTDate)
        End If

        If sFDDate = "" Then
            Errmsg &= "請先輸入訓練迄日" & vbCrLf
        Else
            If Not TIMS.IsDate1(sFDDate) Then
                Errmsg &= "訓練迄日請填日期格式" & vbCrLf
            End If
            sFDDate = TIMS.Cdate3(sFDDate)
        End If
        If Errmsg = "" Then
            Select Case True
                Case DateDiff(DateInterval.Day, CDate(sSTDate), CDate(sFDDate)) < 0
                    Errmsg &= "訓練起日不能超過訓練迄日" & vbCrLf
                Case DateDiff(DateInterval.Day, CDate(sSTDate), CDate(sFDDate)) = 0
                    Errmsg &= "訓練起日不能和訓練迄日同一天" & vbCrLf
            End Select
        End If

        'sPName = "" '清除PName改為組合時間
        sPName = sddlpnH1 & ":" & sddlpnM1 & "~" & sddlpnH2 & ":" & sddlpnM2
        'If Trim(PName.Text) <> "" Then
        '    PName.Text = Trim(PName.Text)
        '    If PName.Text.ToString.Length > 50 Then
        '        Errmsg &= "授課時間，長度超過限制範圍50文字長度" & vbCrLf
        '    End If
        'Else
        '    PName.Text = ""
        '    PName.Text = ddlpnH1.SelectedValue & ":" & ddlpnM1.SelectedValue & "~" & ddlpnH2.SelectedValue & ":" & ddlpnM2.SelectedValue
        'End If

        '20090318--原程式判斷是否整數是mark起來的，將其mark取消。
        sPHour = TIMS.ClearSQM(sPHour)
        If sPHour = "" Then '上課時數
            Errmsg &= "時數未填寫，請填數字" & vbCrLf
        Else
            If Not IsNumeric(sPHour) Then
                Errmsg &= "時數請填數字格式" & vbCrLf
            Else
                If Not TIMS.IsNumeric2(sPHour) Then
                    Errmsg &= "課程大綱的時數欄位 內容需為整數，請修正!!" & vbCrLf
                Else
                    sPHour = CInt(Val(sPHour))
                    If Val(sPHour) <= 0 Then
                        Errmsg &= "時數必須大於0" & vbCrLf
                    ElseIf Val(sPHour) > 4 Then
                        Errmsg &= "時數必須小於等於4" & vbCrLf
                    End If
                End If
            End If
        End If

        sPCont = TIMS.ClearSQM(sPCont)
        'sPCont = TIMS.HtmlDecode1(sPCont)
        '課程進度／內容
        If sPCont = "" Then
            Errmsg &= "課程進度／內容未填寫" & vbCrLf
        Else
            'PCont.Text = Trim(PCont.Text)
            If sPCont.Length > 250 Then
                Errmsg &= "課程進度／內容，長度超過限制範圍250文字長度" & vbCrLf
            End If
        End If

        If sSTrainDate = "" Then
            Errmsg &= "請輸入上課日期" & vbCrLf
        Else
            If Not TIMS.IsDate1(sSTrainDate) Then
                Errmsg &= "上課日期請填日期格式" & vbCrLf
            Else
                If Errmsg = "" Then
                    sSTrainDate = CDate(sSTrainDate).ToString("yyyy/MM/dd")
                End If
                If Not dt Is Nothing Then
                    If sPTDID <> "" Then
                        '修改
                        If dt.Rows.Count > 0 AndAlso IsNumeric(sPTDID) Then
                            Dim xHour As Integer = 0 '計算當日時數不可大於8小時
                            xHour = 0
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted _
                                    AndAlso dt.Select("PTDID='" & sPTDID & "'").Length = 0 Then '已刪除者不可做更動
                                    If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 _
                                        AndAlso dt.Rows(i).Item("PName").ToString = sPName Then
                                        Errmsg &= "此日期+授課時間已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                    If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 Then
                                        Try
                                            xHour += CInt(dt.Rows(i).Item("PHour"))
                                        Catch ex As Exception
                                            Errmsg &= "上課時數異常，請重新載入計算" & vbCrLf
                                            Exit For
                                        End Try
                                    End If
                                End If
                            Next
                            If Errmsg = "" Then
                                xHour += CInt(sPHour)
                                If xHour > 8 Then
                                    Errmsg &= "該日上課時數超過8小時，請重新填寫" & vbCrLf
                                End If
                            End If
                        End If
                    Else
                        '新增
                        If dt.Rows.Count > 0 Then
                            Dim xHour As Integer = 0 '計算當日時數不可大於8小時
                            xHour = 0
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 _
                                        AndAlso dt.Rows(i).Item("PName").ToString = sPName Then
                                        Errmsg &= "此日期+授課時間已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                    If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 Then
                                        Try
                                            xHour += CInt(dt.Rows(i).Item("PHour"))
                                        Catch ex As Exception
                                            Errmsg &= "上課時數異常，請重新載入計算" & vbCrLf
                                            Exit For
                                        End Try
                                    End If
                                End If
                            Next
                            If Errmsg = "" Then
                                xHour += CInt(sPHour)
                                If xHour > 8 Then
                                    Errmsg &= "該日上課時數超過8小時，請重新填寫" & vbCrLf
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If Errmsg = "" Then
            Select Case True
                Case DateDiff(DateInterval.Day, CDate(sSTDate), CDate(sSTrainDate)) < 0
                    Errmsg &= "課程大網日期不能超過訓練起日" & vbCrLf
                Case DateDiff(DateInterval.Day, CDate(sSTrainDate), CDate(sFDDate)) < 0
                    Errmsg &= "課程大網日期不能超過訓練迄日" & vbCrLf
            End Select
        End If
        '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
        Select Case CInt(sClassification1)
            Case 0
                Errmsg &= "請選擇課程大網的學/術科(必填)" & vbCrLf
            Case 1
                If sPTID1 = "" Then
                    Errmsg &= "請選擇課程大網的上課地點(必填)" & vbCrLf
                End If
                If sPTID1 <> "" Then
                    If Not TIMS.Check_SciPTID(sComidValue, sPTID1, objconn) Then
                        Errmsg &= "課程大網的上課地點學科場地已被刪除，請重新選擇" & vbCrLf
                    End If
                End If
            Case 2
                If sPTID2 = "" Then
                    Errmsg &= "請選擇課程大網的上課地點(必填)" & vbCrLf
                End If
                If sPTID2 <> "" Then
                    If Not TIMS.Check_TechPTID(sComidValue, sPTID2, objconn) Then
                        Errmsg &= "課程大網的上課地點術科場地已被刪除，請重新選擇" & vbCrLf
                    End If
                End If
        End Select

        If sOLessonTeah1Value = "" Then
            Errmsg &= "請選擇課程大網的任課教師(必填)" & vbCrLf
        End If
        If sOLessonTeah2Value <> "" Then
            If sOLessonTeah2Value = sOLessonTeah1Value Then
                Errmsg &= "任課教師與助教為同一人錯誤" & vbCrLf
            End If
        End If

        '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
        'If CostID2.SelectedValue = "" Then
        '    Errmsg &= "請選擇項目" & vbCrLf
        'End If
        'If Errmsg <> "" Then
        '    Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
        '    Common.MessageBox(Me, Errmsg)
        '    Exit Function
        'End If

        'Dim htTrainDesc As New Hashtable()
        htTrainDesc("sTPERIOD28") = sTPERIOD28
        htTrainDesc("sSTDate") = sSTDate
        htTrainDesc("sFDDate") = sFDDate
        htTrainDesc("sPName") = sPName
        htTrainDesc("sddlpnH1") = sddlpnH1
        htTrainDesc("sddlpnM1") = sddlpnM1
        htTrainDesc("sddlpnH2") = sddlpnH2
        htTrainDesc("sddlpnM2") = sddlpnM2
        htTrainDesc("sPHour") = sPHour
        htTrainDesc("sPCont") = sPCont
        htTrainDesc("sSTrainDate") = sSTrainDate
        htTrainDesc("sClassification1") = sClassification1
        htTrainDesc("sComidValue") = sComidValue
        htTrainDesc("sPTID1") = sPTID1
        htTrainDesc("sPTID2") = sPTID2
        htTrainDesc("sOLessonTeah1Value") = sOLessonTeah1Value
        htTrainDesc("sOLessonTeah2Value") = sOLessonTeah2Value

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

#Region "NO USE"
    'If STDate.Text = "" Then
    '    Errmsg &= "請先輸入訓練起日" & vbCrLf
    'Else
    '    If Not IsDate(STDate.Text) Then
    '        Errmsg &= "訓練起日請填日期格式" & vbCrLf
    '    End If
    'End If
    'Try
    '    If Errmsg = "" Then STDate.Text = CDate(STDate.Text).ToString("yyyy/MM/dd")
    'Catch ex As Exception
    '    Errmsg &= "訓練起日請填日期格式" & vbCrLf
    'End Try

    'If FDDate.Text = "" Then
    '    Errmsg &= "請先輸入訓練迄日" & vbCrLf
    'Else
    '    If Not IsDate(FDDate.Text) Then
    '        Errmsg &= "訓練迄日請填日期格式" & vbCrLf
    '    End If
    'End If
    'Try
    '    If Errmsg = "" Then FDDate.Text = CDate(FDDate.Text).ToString("yyyy/MM/dd")
    'Catch ex As Exception
    '    Errmsg &= "訓練迄日請填日期格式" & vbCrLf
    'End Try

    'If Errmsg = "" Then
    '    Select Case True
    '        Case DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FDDate.Text)) < 0
    '            Errmsg &= "訓練起日不能超過訓練迄日" & vbCrLf
    '        Case DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FDDate.Text)) = 0
    '            Errmsg &= "訓練起日不能和訓練迄日同一天" & vbCrLf
    '    End Select
    'End If

    'PNameTxt.Text = "" '清除PName改為組合時間
    'PNameTxt.Text = Eddlh1.SelectedValue & ":" & Eddlm1.SelectedValue & "~" & Eddlh2.SelectedValue & ":" & Eddlm2.SelectedValue
    ''If PNameTxt.Text.ToString.Length > 50 Then
    ''    Errmsg &= "授課時間，長度超過限制範圍50文字長度" & vbCrLf
    ''End If

    ''20090318--原程式判斷是否整數是mark起來的，將其mark取消。
    'If PHourTxt.Text.Trim = "" Then '上課時數
    '    Errmsg &= "時數未填寫，請填數字" & vbCrLf
    'Else
    '    PHourTxt.Text = PHourTxt.Text.Trim
    '    If Not IsNumeric(PHourTxt.Text) Then
    '        Errmsg &= "時數請填數字格式" & vbCrLf
    '    Else
    '        Try
    '            PHourTxt.Text = CInt(PHourTxt.Text)
    '            If PHourTxt.Text <= 0 Then
    '                Errmsg &= "時數必須大於0" & vbCrLf
    '            ElseIf PHourTxt.Text > 4 Then
    '                Errmsg &= "時數必須小於等於4" & vbCrLf
    '            End If
    '        Catch ex As Exception
    '            Errmsg &= "時數請填數字格式" & vbCrLf
    '        End Try
    '    End If
    'End If

    ''課程進度／內容
    'If Trim(PContEdit.Text) = "" Then
    '    Errmsg &= "課程進度／內容未填寫" & vbCrLf
    'Else
    '    PContEdit.Text = Trim(PContEdit.Text)
    '    If PContEdit.Text.ToString.Length > 250 Then
    '        Errmsg &= "課程進度／內容，長度超過限制範圍250文字長度" & vbCrLf
    '    End If
    'End If
    ''If PContEdit.Text.ToString.Length > 250 Then
    ''    Errmsg &= "課程進度／內容，長度超過限制範圍250文字長度" & vbCrLf
    ''End If
    'If STrainDateTxt.Text = "" Then
    '    Errmsg &= "請輸入上課日期" & vbCrLf
    'Else

    '    If Not TIMS.IsDate1(STrainDateTxt.Text) Then
    '        Errmsg &= "上課日期請填日期格式" & vbCrLf
    '    Else
    '        Try
    '            STrainDateTxt.Text = CDate(STrainDateTxt.Text).ToString("yyyy/MM/dd")
    '        Catch ex As Exception
    '            Errmsg &= "上課日期請填日期格式" & vbCrLf
    '        End Try
    '        If Errmsg = "" Then
    '            If Not dt Is Nothing Then
    '                If dt.Rows.Count > 0 And IsNumeric(e.CommandArgument) Then
    '                    Dim xHour As Integer = 0 '計算當日時數不可大於8小時
    '                    xHour = 0
    '                    For i As Int16 = 0 To dt.Rows.Count - 1
    '                        If Not dt.Rows(i).RowState = DataRowState.Deleted _
    '                            AndAlso dt.Select("PTDID='" & e.CommandArgument & "'").Length = 0 Then '已刪除者不可做更動
    '                            If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(STrainDateTxt.Text)) = 0 _
    '                                AndAlso dt.Rows(i).Item("PName").ToString = PNameTxt.Text Then
    '                                Errmsg &= "此日期+授課時間已在表格中" & vbCrLf
    '                                Exit For
    '                            End If
    '                            If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(STrainDateTxt.Text)) = 0 Then
    '                                Try
    '                                    xHour += CInt(dt.Rows(i).Item("PHour"))
    '                                Catch ex As Exception
    '                                    Errmsg &= "上課時數異常，請重新載入計算" & vbCrLf
    '                                    Exit For
    '                                End Try
    '                            End If
    '                        End If
    '                    Next
    '                    If Errmsg = "" Then
    '                        xHour += CInt(PHourTxt.Text)
    '                        If xHour > 8 Then
    '                            Errmsg &= "該日上課時數超過8小時，請重新填寫" & vbCrLf
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If

    '    End If
    'End If
    ''If PHourTxt.Text = "" Then
    ''    Errmsg &= "時數請填數字" & vbCrLf
    ''Else
    ''    If Not IsNumeric(PHourTxt.Text) Then
    ''        Errmsg &= "時數請填數字格式" & vbCrLf
    ''        'Else
    ''        '    If CInt(PHourTxt.Text).ToString <> PHourTxt.Text.ToString Then
    ''        '        Errmsg &= "時數請填整數格式" & vbCrLf
    ''        '    End If
    ''    End If
    ''End If
    'If Errmsg = "" Then
    '    Select Case True
    '        Case DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(STrainDateTxt.Text)) < 0
    '            Errmsg &= "課程大網日期不能超過訓練起日" & vbCrLf
    '        Case DateDiff(DateInterval.Day, CDate(STrainDateTxt.Text), CDate(FDDate.Text)) < 0
    '            Errmsg &= "課程大網日期不能超過訓練迄日" & vbCrLf
    '    End Select
    'End If
    ''(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
    'Select Case drpClassEdit.SelectedValue
    '    Case "1"
    '        If drpPTIDEdit1.SelectedValue = "" Then
    '            Errmsg &= "請選擇課程大網的上課地點(必填)" & vbCrLf
    '        End If
    '        If drpPTIDEdit1.SelectedValue <> "" Then
    '            If Not TIMS.Check_SciPTID(ComidValue.Value, drpPTIDEdit1.SelectedValue) Then
    '                Errmsg &= "課程大網的上課地點學科場地已被刪除，請重新選擇" & vbCrLf
    '            End If
    '        End If
    '    Case "2"
    '        If drpPTIDEdit2.SelectedValue = "" Then
    '            Errmsg &= "請選擇課程大網的上課地點(必填)" & vbCrLf
    '        End If
    '        If drpPTIDEdit2.SelectedValue <> "" Then
    '            If Not TIMS.Check_TechPTID(ComidValue.Value, drpPTIDEdit2.SelectedValue) Then
    '                Errmsg &= "課程大網的上課地點術科場地已被刪除，請重新選擇" & vbCrLf
    '            End If
    '        End If
    'End Select
    'If Tech1ValueEdit.Value = "" Then
    '    Errmsg &= "請選擇課程大網的任課教師(必填)" & vbCrLf
    'End If
    'If Tech2ValueEdit.Value <> "" Then
    '    If Tech2ValueEdit.Value = Tech1ValueEdit.Value Then
    '        Errmsg &= "任課教師與助教為同一人錯誤" & vbCrLf
    '    End If
    'End If
#End Region

    '新增 課程大網
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Session(cst_TrainDescTable) Is Nothing Then
            Call CreateTrainDesc()
        End If
        Dim dt As DataTable = Session(cst_TrainDescTable) 'PLAN_TRAINDESC 
        dt.Columns("PTDID").AutoIncrement = True
        dt.Columns("PTDID").AutoIncrementSeed = -1
        dt.Columns("PTDID").AutoIncrementStep = -1

        Dim dr As DataRow = Nothing
        Dim Errmsg As String = ""

        Dim sTPERIOD28 As String = "" '授課時段'早上'下午'晚上
        sTPERIOD28 = ""
        If TPERIOD28_1.Checked Then sTPERIOD28 &= "Y" Else sTPERIOD28 &= "N"
        If TPERIOD28_2.Checked Then sTPERIOD28 &= "Y" Else sTPERIOD28 &= "N"
        If TPERIOD28_3.Checked Then sTPERIOD28 &= "Y" Else sTPERIOD28 &= "N"

        'PCont.Text = TIMS.HtmlDecode1(PCont.Text)
        PCont.Text = TIMS.ClearSQM(PCont.Text)
        Dim htTrainDesc As New Hashtable()
        htTrainDesc.Add("sTPERIOD28", sTPERIOD28) '授課時段
        htTrainDesc.Add("sSTDate", STDate.Text)
        htTrainDesc.Add("sFDDate", FDDate.Text)
        'PName.Text = TIMS.ClearSQM(PName.Text)
        htTrainDesc.Add("sPName", PName.Text)
        htTrainDesc.Add("sddlpnH1", ddlpnH1.SelectedValue)
        htTrainDesc.Add("sddlpnM1", ddlpnM1.SelectedValue)
        htTrainDesc.Add("sddlpnH2", ddlpnH2.SelectedValue)
        htTrainDesc.Add("sddlpnM2", ddlpnM2.SelectedValue)
        htTrainDesc.Add("sPHour", PHour.Text)
        htTrainDesc.Add("sPCont", PCont.Text)
        htTrainDesc.Add("sSTrainDate", STrainDate.Text)
        htTrainDesc.Add("sClassification1", Classification1.SelectedValue)
        htTrainDesc.Add("sComidValue", ComidValue.Value)
        htTrainDesc.Add("sPTID1", PTID1.SelectedValue)
        htTrainDesc.Add("sPTID2", PTID2.SelectedValue)
        htTrainDesc.Add("sOLessonTeah1Value", OLessonTeah1Value.Value)
        htTrainDesc.Add("sOLessonTeah2Value", OLessonTeah2Value.Value)
        htTrainDesc.Add("sPTDID", "")
        Call Chk_TrainDescInput(Errmsg, htTrainDesc, dt)
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        'STDate.Text = htTrainDesc("sSTDate")
        'FDDate.Text = htTrainDesc("sFDDate")
        PName.Text = htTrainDesc("sPName")
        PName.Text = TIMS.ClearSQM(PName.Text)
        PHour.Text = htTrainDesc("sPHour")
        PCont.Text = TIMS.ClearSQM(htTrainDesc("sPCont"))
        STrainDate.Text = htTrainDesc("sSTrainDate")

        dr = dt.NewRow
        dt.Rows.Add(dr) '產業人才投資方案--
        dr("TPERIOD28") = sTPERIOD28 '授課時段
        dr("STrainDate") = CDate(STrainDate.Text) '97年產學訓課程大綱-日期
        dr("ETrainDate") = CDate(STrainDate.Text) '97年產學訓課程大綱-日期
        dr("PName") = Trim(PName.Text) '97年產學訓授課時間
        dr("PHour") = Trim(PHour.Text) '時數
        'dr("PCont") = Trim(PCont.Text) '97年產學訓課程內容
        dr("PCont") = TIMS.ClearSQM(PCont.Text)
        dr("Classification1") = CInt(Classification1.SelectedValue) '學科術科
        Select Case CInt(Classification1.SelectedValue)
            Case 1
                If PTID1.SelectedValue <> "" Then
                    dr("PTID") = PTID1.SelectedValue '上課地點
                Else
                    dr("PTID") = Convert.DBNull
                End If
            Case 2
                If PTID2.SelectedValue <> "" Then
                    dr("PTID") = PTID2.SelectedValue '上課地點
                Else
                    dr("PTID") = Convert.DBNull
                End If
        End Select
        dr("TechID") = Convert.DBNull
        If OLessonTeah1Value.Value <> "" Then
            dr("TechID") = OLessonTeah1Value.Value '任課教師
        End If
        dr("TechID2") = Convert.DBNull
        If OLessonTeah2Value.Value <> "" Then
            dr("TechID2") = OLessonTeah2Value.Value '助教
        End If
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        'dt.AcceptChanges()
        Session(cst_TrainDescTable) = dt
        Call CreateTrainDesc()

        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                DataGrid1.EditItemIndex = e.Item.ItemIndex
            Case "del"
                Dim dt As DataTable = Session(cst_Plan_OnClass)
                Dim DGobj As DataGrid = DataGrid1
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If
                'dt = Session(cst_Plan_OnClass)
                If dt.Select("POCID='" & e.CommandArgument & "'").Length <> 0 Then
                    dt.Select("POCID='" & e.CommandArgument & "'")(0).Delete()
                End If
                Session(cst_Plan_OnClass) = dt
                If dt.Rows.Count = 0 Then
                    DataGrid1Table.Visible = False
                Else
                    DataGrid1Table.Visible = True
                    DataGrid1.DataSource = dt
                End If

                DataGrid1.EditItemIndex = -1
            Case "save"
                Dim dt As DataTable
                Dim dr As DataRow
                Dim Weeks As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times As TextBox = e.Item.FindControl("Times2")

                If Not Session(cst_Plan_OnClass) Is Nothing Then
                    dt = Session(cst_Plan_OnClass)
                    If dt.Select("POCID='" & e.CommandArgument & "'").Length <> 0 Then
                        dr = dt.Select("POCID='" & e.CommandArgument & "'")(0)
                        dr("Weeks") = Weeks.SelectedValue
                        dr("Times") = Times.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    End If
                    Session(cst_Plan_OnClass) = dt

                    DataGrid1.EditItemIndex = -1
                End If
            Case "cancel"
                DataGrid1.EditItemIndex = -1
        End Select

        CreateClassTime()
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim strSechObjID As String = "" '查詢按鈕ID
        Dim strAddsObjID As String = "" '維護按鈕ID
        Dim strPrntObjID As String = "" '列印按鈕ID

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Weeks1 As Label = e.Item.FindControl("Weeks1")
                Dim Times1 As Label = e.Item.FindControl("Times1")
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button2")
                Dim btn2 As Button = e.Item.FindControl("Button3")

                btn1.Enabled = Button29.Enabled
                btn2.Enabled = Button29.Enabled
                Weeks1.Text = drv("Weeks").ToString
                Times1.Text = drv("Times").ToString
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn2.CommandArgument = drv("POCID")

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    btn1.Visible = False
                '    btn2.Visible = False
                'End If
                'strAddsObjID = btn1.ClientID & "," & btn2.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End
            Case ListItemType.EditItem
                Dim Weeks2 As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times2 As TextBox = e.Item.FindControl("Times2")
                Dim btn1 As Button = e.Item.FindControl("Button4")
                Dim btn2 As Button = e.Item.FindControl("Button5")
                Dim drv As DataRowView = e.Item.DataItem
                Weeks2 = TIMS.Get_ddlWeeks(Weeks2)

                'With Weeks
                '    .Items.Add(New ListItem("==請選擇==", ""))
                '    .Items.Add(New ListItem("星期一", "星期一"))
                '    .Items.Add(New ListItem("星期二", "星期二"))
                '    .Items.Add(New ListItem("星期三", "星期三"))
                '    .Items.Add(New ListItem("星期四", "星期四"))
                '    .Items.Add(New ListItem("星期五", "星期五"))
                '    .Items.Add(New ListItem("星期六", "星期六"))
                '    .Items.Add(New ListItem("星期日", "星期日"))
                'End With
                Common.SetListItem(Weeks2, drv("Weeks").ToString)
                Times2.Text = drv("Times").ToString
                btn1.CommandArgument = drv("POCID")

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    btn1.Visible = False
                '    btn2.Visible = False
                'End If
                'strAddsObjID = btn1.ClientID & "," & btn2.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim DGobj As DataGrid = DataGrid2
        Dim dt As DataTable = Session(cst_CostItemTable)
        'Dim dr As DataRow
        Select Case e.CommandName
            Case "edit1" '修改
                DataGrid2.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case "del1" '刪除
                If DGobj Is Nothing _
                        OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                Dim sfilter As String = "PCID='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                   AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr As DataRow In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then
                            dr.Delete() '刪除
                        End If
                    Next
                End If

            Case "update1" '更新
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                Dim tdrpCostID As DropDownList = e.Item.FindControl("drpCostID2") '項目
                Dim tOPrice As TextBox = e.Item.FindControl("DataGrid2TextBox1") '單價
                Dim tItemage As TextBox = e.Item.FindControl("DataGrid2TextBox2") '計價數量
                'Dim tItemCost As Label = e.Item.FindControl("DataGrid2Label3b") '小計

                Dim ErrorMsg As String = ""
                ErrorMsg = ""
                'If Convert.ToString(tOPrice.Text).Trim = "" Then
                '    tOPrice.Text = ""
                '    ErrorMsg &= "單價內容 不可為空" & vbCrLf
                'Else
                '    tOPrice.Text = Trim(tOPrice.Text)
                '    If Not IsNumeric(tOPrice.Text) Then
                '        ErrorMsg &= "單價內容 應輸入數字格式" & vbCrLf
                '    End If
                'End If

                tOPrice.Text = TIMS.ClearSQM(tOPrice.Text)
                Dim dOPrice As Double = TIMS.VAL1(tOPrice.Text)
                If tOPrice.Text = "" OrElse dOPrice <= 0 Then
                    ErrorMsg &= "單價內容 不可為空須大於0,不可為0" & vbCrLf
                Else
                    If Not TIMS.IsNumeric2(tOPrice.Text) Then
                        ErrorMsg &= "單價請填數字格式(正整數)須大於0,不可為0" & vbCrLf
                    End If
                End If

                tItemage.Text = TIMS.ClearSQM(tItemage.Text)
                Dim itItemage As Double = TIMS.VAL1(tItemage.Text)
                If tItemage.Text = "" OrElse itItemage <= 0 Then
                    ErrorMsg &= "計價數量 請填數字須大於0,不可為0" & vbCrLf
                Else
                    If Not TIMS.IsNumeric2(tItemage.Text) Then
                        ErrorMsg &= "計價數量 請填數字格式(正整數)須大於0,不可為0" & vbCrLf
                    End If
                End If

                Dim CostID2Value As String = tdrpCostID.SelectedValue
                '(沒錯誤繼續驗證)非學分班
                If ErrorMsg = "" AndAlso RadioButtonList1.SelectedValue = "N" Then  '(沒錯誤繼續驗證)非學分班
                    If CostID2Value = "01" And tOPrice.Text > 1600 Then
                        ErrorMsg &= "學科-鐘點費(外聘)的單價不得大於1,600元" & vbCrLf
                    End If
                    If CostID2Value = "12" And tOPrice.Text > 800 Then
                        ErrorMsg &= "學科-鐘點費(內聘)的單價不得大於800元" & vbCrLf
                    End If
                    If CostID2Value = "02" And tOPrice.Text > 1600 Then
                        ErrorMsg &= "術科-鐘點費(外聘)的單價不得大於1,600元" & vbCrLf
                    End If
                    If CostID2Value = "13" And tOPrice.Text > 800 Then
                        ErrorMsg &= "術科-鐘點費(內聘)的單價不得大於800元" & vbCrLf
                    End If
                    If CostID2Value = "03" And tOPrice.Text > 800 Then
                        ErrorMsg &= "教材費的單價不得大於800元" & vbCrLf
                    End If
                    If CostID2Value = "05" And tOPrice.Text > 2500 Then
                        ErrorMsg &= "場地費的單價不得大於2,500元" & vbCrLf
                    End If
                    If CostID2Value = "06" And tOPrice.Text > 10000 Then
                        ErrorMsg &= "宣導費的單價不得大於10,000元" & vbCrLf
                    End If
                    If CostID2Value = "06" And tItemage.Text <> 1 Then
                        ErrorMsg &= "宣導費的計價數量須等於1" & vbCrLf
                    End If
                    If CostID2Value = "09" And tOPrice.Text > 400 Then
                        ErrorMsg &= "術科助教費用的單價不得大於400元" & vbCrLf
                    End If
                    If CostID2Value = "10" And tOPrice.Text > 160 Then
                        ErrorMsg &= "工作人員費的單價不得大於160元" & vbCrLf
                    End If
                    'If CostID2Value = "14" Then
                    '    Common.MessageBox(Me, "【雜支】：" & vbCrLf & vbCrLf & Space(20) & "以實際辦訓費用總額" & vbCrLf & vbCrLf & Space(20) & "  (不包括出席費、鐘點費、稿費、差旅費、工作人員服務費及管理費)" & vbCrLf & vbCrLf & Space(20) & "  之百分之五以內編列。")
                    'End If
                End If

                If ErrorMsg <> "" Then
                    'DataGrid2.EditItemIndex = -1 '還原修改列數
                    Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
                    Common.MessageBox(Me, ErrorMsg)
                    Exit Sub
                Else
                    '無錯誤存取
                    If Convert.ToString(DataGrid2.DataKeys(e.Item.ItemIndex)) <> "" _
                        AndAlso dt.Select("PCID='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                        Dim dr As DataRow = dt.Select("PCID='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'")(0)
                        dr("OPrice") = TIMS.ChangeIDNO(tOPrice.Text)
                        dr("Itemage") = TIMS.ChangeIDNO(tItemage.Text)
                        'dr("ItemCost") = ItemCost.Text
                    End If
                    DataGrid2.EditItemIndex = -1 '還原修改列數
                End If

            Case "cancel1" '取消
                DataGrid2.EditItemIndex = -1 '還原修改列數
        End Select

        Session(cst_CostItemTable) = dt '要新 Session(cst_CostItemTable)
        CreateCostItem() '建立 Session(cst_CostItemTable)

        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")

    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Dim strSechObjID As String = "" '查詢按鈕ID
        Dim strAddsObjID As String = "" '維護按鈕ID
        Dim strPrntObjID As String = "" '列印按鈕ID

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim sdrpCostID As DropDownList = e.Item.FindControl("drpCostID1") '項目
                Dim sOPrice As Label = e.Item.FindControl("DataGrid2Label1") '單價
                Dim sItemage As Label = e.Item.FindControl("DataGrid2Label2") '計價數量
                Dim sItemCost As Label = e.Item.FindControl("DataGrid2Label3") '小計
                'Dim ItemCostLab As Label = e.Item.FindControl("ItemCostLab")
                Dim btn1 As Button = e.Item.FindControl("Button12") '修改(del1)
                Dim btn2 As Button = e.Item.FindControl("Button13") '刪除(edit1)

                sdrpCostID = TIMS.Get_KeyControl(sdrpCostID, "KEY_COSTITEM2", "CostName", "CostID", objconn)
                If drv("CostID").ToString <> "" Then
                    Common.SetListItem(sdrpCostID, drv("CostID").ToString)
                End If
                sOPrice.Text = drv("OPrice")
                sItemage.Text = 1
                If Convert.ToString(drv("Itemage")) <> "" AndAlso IsNumeric(drv("Itemage")) Then
                    If CInt(drv("Itemage")) > 1 Then
                        sItemage.Text = Convert.ToString(drv("Itemage"))
                    End If
                End If
                'e.Item.Cells(3).Text = CDbl(OPrice.Text) * CDbl(Itemage.Text) * CDbl(ItemCost.Text)
                'e.Item.Cells(cst_小計).Text = CInt(CDbl(sOPrice.Text) * CDbl(sItemage.Text))  '**by Milor 20080611--User要求小計作四捨五入取整數
                Try
                    sItemCost.Text = CInt(CDbl(sOPrice.Text) * CDbl(sItemage.Text)) '--User要求小計作四捨五入取整數
                Catch ex As Exception
                    sItemCost.Text = 0
                End Try
                dbld2TempTotal += CDbl(sItemCost.Text) '暫存小計，供外部使用

                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                btn1.Enabled = Button9.Enabled
                btn2.Enabled = Button9.Enabled

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    btn1.Visible = False
                '    btn2.Visible = False
                'End If
                'strAddsObjID = btn1.ClientID & "," & btn2.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End

            Case ListItemType.EditItem
                Dim tdrpCostID As DropDownList = e.Item.FindControl("drpCostID2") '項目
                Dim tOPrice As TextBox = e.Item.FindControl("DataGrid2TextBox1") '單價
                Dim tItemage As TextBox = e.Item.FindControl("DataGrid2TextBox2") '計價數量
                Dim tItemCost As Label = e.Item.FindControl("DataGrid2Label3b") '小計

                Dim Button14 As Button = e.Item.FindControl("Button14") '更新(update1)
                Dim Button15 As Button = e.Item.FindControl("Button15") '取消(cancel1)

                tdrpCostID = TIMS.Get_KeyControl(tdrpCostID, "KEY_COSTITEM2", "CostName", "CostID", objconn)
                If drv("CostID").ToString <> "" Then
                    Common.SetListItem(tdrpCostID, drv("CostID").ToString)
                End If
                tOPrice.Text = drv("OPrice")

                tItemage.Text = 1
                If Convert.ToString(drv("Itemage")) <> "" AndAlso IsNumeric(drv("Itemage")) Then
                    If CInt(drv("Itemage")) > 1 Then
                        tItemage.Text = Convert.ToString(drv("Itemage"))
                    End If
                End If

                Try
                    tItemCost.Text = CInt(CDbl(tOPrice.Text) * CDbl(tItemage.Text)) '--User要求小計作四捨五入取整數
                Catch ex As Exception
                    tItemCost.Text = 0
                End Try
                Button14.Enabled = Button9.Enabled
                Button15.Enabled = True

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    Button14.Visible = False
                '    Button15.Visible = False
                'End If
                'strAddsObjID = Button14.ClientID & "," & Button15.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End

        End Select
    End Sub

    '修改 課程大網
    Private Sub Datagrid3_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid3.ItemCommand
        Dim dt As DataTable = Session(cst_TrainDescTable) '取得SESSION到 dt

        Select Case e.CommandName
            Case "edit"
                Datagrid3.EditItemIndex = e.Item.ItemIndex
                Session(cst_TrainDescTable) = dt
                '編輯列 開啟
            Case "del" '刪除
                Dim DGobj As DataGrid = Datagrid3
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                If Convert.ToString(Datagrid3.DataKeys(e.Item.ItemIndex)) <> "" Then
                    If dt.Select("PTDID='" & Datagrid3.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                        Dim dr As DataRow = dt.Select("PTDID='" & Datagrid3.DataKeys(e.Item.ItemIndex) & "'")(0)
                        dr.Delete()
                    End If
                End If
                'dt.AcceptChanges()
                Session(cst_TrainDescTable) = dt
                Datagrid3Table.Visible = False
                If dt.Rows.Count > 0 Then
                    Datagrid3Table.Visible = True
                    Datagrid3.DataSource = dt
                End If
                Datagrid3.EditItemIndex = -1 '關閉編輯列

            Case "save" '存檔
                Dim DGobj As DataGrid = Datagrid3
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                Dim Errmsg As String = ""
                Dim TPERIOD28_1e As CheckBox = e.Item.FindControl("TPERIOD28_1e")
                Dim TPERIOD28_2e As CheckBox = e.Item.FindControl("TPERIOD28_2e")
                Dim TPERIOD28_3e As CheckBox = e.Item.FindControl("TPERIOD28_3e")
                Dim STrainDateTxt As TextBox = e.Item.FindControl("STrainDateTxt")
                Dim Img1 As HtmlImage = e.Item.FindControl("Img2")
                Dim Eddlh1 As DropDownList = e.Item.FindControl("Eddlh1")
                Dim Eddlm1 As DropDownList = e.Item.FindControl("Eddlm1")
                Dim Eddlh2 As DropDownList = e.Item.FindControl("Eddlh2")
                Dim Eddlm2 As DropDownList = e.Item.FindControl("Eddlm2")
                Dim PNameTxt As TextBox = e.Item.FindControl("PNameTxt")
                Dim PHourTxt As TextBox = e.Item.FindControl("PHourTxt")
                Dim PContEdit As TextBox = e.Item.FindControl("PContEdit")
                Dim drpClassEdit As DropDownList = e.Item.FindControl("drpClassEdit")
                Dim drpPTIDEdit1 As DropDownList = e.Item.FindControl("drpPTIDEdit1")
                Dim drpPTIDEdit2 As DropDownList = e.Item.FindControl("drpPTIDEdit2")
                Dim Tech1ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech1ValueEdit")
                Dim Tech1Edit As TextBox = e.Item.FindControl("Tech1Edit")
                Dim Tech2ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech2ValueEdit")
                Dim Tech2Edit As TextBox = e.Item.FindControl("Tech2Edit")

                Dim sTPERIOD28 As String = ""
                sTPERIOD28 = ""
                If TPERIOD28_1e.Checked Then sTPERIOD28 &= "Y" Else sTPERIOD28 &= "N"
                If TPERIOD28_2e.Checked Then sTPERIOD28 &= "Y" Else sTPERIOD28 &= "N"
                If TPERIOD28_3e.Checked Then sTPERIOD28 &= "Y" Else sTPERIOD28 &= "N"

                'PContEdit.Text = TIMS.HtmlDecode1(PContEdit.Text)
                PContEdit.Text = TIMS.ClearSQM(PContEdit.Text)
                Dim htTrainDesc As New Hashtable()
                htTrainDesc.Add("sTPERIOD28", sTPERIOD28)
                htTrainDesc.Add("sSTDate", STDate.Text)
                htTrainDesc.Add("sFDDate", FDDate.Text)
                htTrainDesc.Add("sPName", PNameTxt.Text)
                htTrainDesc.Add("sddlpnH1", Eddlh1.SelectedValue)
                htTrainDesc.Add("sddlpnM1", Eddlm1.SelectedValue)
                htTrainDesc.Add("sddlpnH2", Eddlh2.SelectedValue)
                htTrainDesc.Add("sddlpnM2", Eddlm2.SelectedValue)
                htTrainDesc.Add("sPHour", PHourTxt.Text)
                htTrainDesc.Add("sPCont", PContEdit.Text)
                htTrainDesc.Add("sSTrainDate", STrainDateTxt.Text)
                htTrainDesc.Add("sClassification1", drpClassEdit.SelectedValue)
                htTrainDesc.Add("sComidValue", ComidValue.Value)
                htTrainDesc.Add("sPTID1", drpPTIDEdit1.SelectedValue)
                htTrainDesc.Add("sPTID2", drpPTIDEdit2.SelectedValue)
                htTrainDesc.Add("sOLessonTeah1Value", Tech1ValueEdit.Value)
                htTrainDesc.Add("sOLessonTeah2Value", Tech2ValueEdit.Value)
                htTrainDesc.Add("sPTDID", e.CommandArgument)
                Call Chk_TrainDescInput(Errmsg, htTrainDesc, dt)
                '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
                If Errmsg <> "" Then
                    Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
                    Page.RegisterStartupScript("Londing3", "<script>showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');</script>")
                    Common.MessageBox(Me, Errmsg)
                    Exit Sub
                End If
                'STDate.Text = htTrainDesc("sSTDate")
                'FDDate.Text = htTrainDesc("sFDDate")
                PNameTxt.Text = htTrainDesc("sPName")
                PHourTxt.Text = htTrainDesc("sPHour")
                PContEdit.Text = TIMS.ClearSQM(htTrainDesc("sPCont"))
                STrainDateTxt.Text = htTrainDesc("sSTrainDate")

                If IsNumeric(e.CommandArgument) Then
                    If dt.Select("PTDID='" & e.CommandArgument & "'").Length <> 0 Then
                        Dim dr As DataRow = dt.Select("PTDID='" & e.CommandArgument & "'")(0)
                        dr("TPERIOD28") = sTPERIOD28
                        dr("STrainDate") = STrainDateTxt.Text
                        dr("ETrainDate") = STrainDateTxt.Text
                        dr("PName") = Trim(PNameTxt.Text)
                        dr("PHour") = Trim(PHourTxt.Text)
                        'dr("PCont") = Trim(PContEdit.Text)
                        dr("PCont") = TIMS.ClearSQM(PContEdit.Text)
                        dr("Classification1") = CInt(drpClassEdit.SelectedValue) '學科術科
                        Select Case drpClassEdit.SelectedValue
                            Case "1"
                                If drpPTIDEdit1.SelectedValue <> "" Then
                                    dr("PTID") = drpPTIDEdit1.SelectedValue '上課地點
                                Else
                                    dr("PTID") = Convert.DBNull
                                End If
                            Case "2"
                                If drpPTIDEdit2.SelectedValue <> "" Then
                                    dr("PTID") = drpPTIDEdit2.SelectedValue '上課地點
                                Else
                                    dr("PTID") = Convert.DBNull
                                End If
                        End Select
                        If Tech1ValueEdit.Value <> "" Then
                            dr("TechID") = Tech1ValueEdit.Value '任課教師
                        Else
                            dr("TechID") = Convert.DBNull
                        End If
                        If Tech2ValueEdit.Value <> "" Then
                            dr("TechID2") = Tech2ValueEdit.Value '助教
                        Else
                            dr("TechID2") = Convert.DBNull
                        End If
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    End If
                End If
                'dt.AcceptChanges()
                Session(cst_TrainDescTable) = dt
                Datagrid3.EditItemIndex = -1
            Case "cancel"
                Datagrid3.EditItemIndex = -1
        End Select
        Call CreateTrainDesc()
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")

        'Select Case e.CommandName
        '    Case "edit"
        '        Page.RegisterStartupScript("Londing2", "<script>Layer_change(5);showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');</script>")
        '    Case Else
        '        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
        'End Select
    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Dim strSechObjID As String = "" '查詢按鈕ID
        Dim strAddsObjID As String = "" '維護按鈕ID
        Dim strPrntObjID As String = "" '列印按鈕ID

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'SHOW
                Dim drv As DataRowView = e.Item.DataItem
                'Dim TPERIOD28_label As Label = e.Item.FindControl("TPERIOD28_label")
                Dim TPERIOD28_1t As CheckBox = e.Item.FindControl("TPERIOD28_1t")
                Dim TPERIOD28_2t As CheckBox = e.Item.FindControl("TPERIOD28_2t")
                Dim TPERIOD28_3t As CheckBox = e.Item.FindControl("TPERIOD28_3t")
                Dim STrainDateLabel As Label = e.Item.FindControl("STrainDateLabel")
                Dim PNameLabel As Label = e.Item.FindControl("PNameLabel")
                Dim PHourLabel As Label = e.Item.FindControl("PHourLabel")
                Dim PContText As TextBox = e.Item.FindControl("PContText")
                Dim drpClassification1 As DropDownList = e.Item.FindControl("drpClassification1")
                Dim drpPTID As DropDownList = e.Item.FindControl("drpPTID")
                Dim Tech1Value As HtmlInputHidden = e.Item.FindControl("Tech1Value")
                Dim Tech1Text As TextBox = e.Item.FindControl("Tech1Text")
                Dim Tech2Value As HtmlInputHidden = e.Item.FindControl("Tech2Value")
                Dim Tech2Text As TextBox = e.Item.FindControl("Tech2Text")

                Dim btn1 As Button = e.Item.FindControl("Button6") 'edit
                Dim btn2 As Button = e.Item.FindControl("Button7") 'del

                'TPERIOD28_label.Text = TIMS.Chg_TPERIOD28_VAL(Convert.ToString(drv("TPERIOD28")))
                TPERIOD28_1t.Checked = False
                TPERIOD28_2t.Checked = False
                TPERIOD28_3t.Checked = False
                If Convert.ToString(drv("TPERIOD28")) <> "" _
                    AndAlso Convert.ToString(drv("TPERIOD28")).Length >= 3 Then
                    If Convert.ToString(drv("TPERIOD28")).Substring(0, 1) = "Y" Then TPERIOD28_1t.Checked = True
                    If Convert.ToString(drv("TPERIOD28")).Substring(1, 1) = "Y" Then TPERIOD28_2t.Checked = True
                    If Convert.ToString(drv("TPERIOD28")).Substring(2, 1) = "Y" Then TPERIOD28_3t.Checked = True
                End If

                If drv("STrainDate").ToString <> "" Then
                    'STrainDateLabel.Text = Common.FormatDate(drv("STrainDate").ToString)
                    STrainDateLabel.Text = TIMS.Cdate3(drv("STrainDate"))
                End If
                PNameLabel.Text = drv("PName").ToString '時間
                PHourLabel.Text = drv("PHour").ToString '時數
                PContText.Text = drv("PCont").ToString '內容
                PContText.Text = TIMS.HtmlDecode1(PContText.Text)

                If drv("Classification1").ToString <> "" Then
                    Common.SetListItem(drpClassification1, drv("Classification1").ToString)

                    Select Case drpClassification1.SelectedValue
                        Case "1"  '學科
                            '將Hid_ComIDNO.Value 塞入有效值
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "ComIDNO", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then
                                Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                            End If
                            drpPTID = TIMS.Get_SciPTID(drpPTID, Hid_ComIDNO.Value, 1, objconn)

                        Case "2"  '術科
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "ComIDNO", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then
                                Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                            End If
                            drpPTID = TIMS.Get_TechPTID(drpPTID, Hid_ComIDNO.Value, 1, objconn)

                    End Select

                    If drv("PTID").ToString <> "" Then
                        Common.SetListItem(drpPTID, drv("PTID").ToString)
                    End If
                End If

                If Convert.ToString(drv("TechID")) <> "" Then
                    Tech1Value.Value = drv("TechID").ToString
                    Tech1Text.Text = TIMS.Get_TeachCName(Tech1Value.Value, objconn) '
                End If
                If Convert.ToString(drv("TechID2")) <> "" Then
                    Tech2Value.Value = drv("TechID2").ToString
                    Tech2Text.Text = TIMS.Get_TeachCName(Tech2Value.Value, objconn) '
                End If

                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn2.CommandArgument = drv("PTDID").ToString

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    btn1.Visible = False
                '    btn2.Visible = False
                'End If
                'strAddsObjID = btn1.ClientID & "," & btn2.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End
            Case ListItemType.EditItem
                'EDIT
                Dim drv As DataRowView = e.Item.DataItem
                Dim TPERIOD28_1e As CheckBox = e.Item.FindControl("TPERIOD28_1e")
                Dim TPERIOD28_2e As CheckBox = e.Item.FindControl("TPERIOD28_2e")
                Dim TPERIOD28_3e As CheckBox = e.Item.FindControl("TPERIOD28_3e")
                Dim STrainDateTxt As TextBox = e.Item.FindControl("STrainDateTxt")
                Dim Img1 As HtmlImage = e.Item.FindControl("Img2")
                Dim Eddlh1 As DropDownList = e.Item.FindControl("Eddlh1")
                Dim Eddlm1 As DropDownList = e.Item.FindControl("Eddlm1")
                Dim Eddlh2 As DropDownList = e.Item.FindControl("Eddlh2")
                Dim Eddlm2 As DropDownList = e.Item.FindControl("Eddlm2")
                Dim PNameTxt As TextBox = e.Item.FindControl("PNameTxt")
                Dim PHourTxt As TextBox = e.Item.FindControl("PHourTxt")
                Dim PContEdit As TextBox = e.Item.FindControl("PContEdit")
                Dim drpClassEdit As DropDownList = e.Item.FindControl("drpClassEdit")
                Dim drpPTIDEdit1 As DropDownList = e.Item.FindControl("drpPTIDEdit1")
                Dim drpPTIDEdit2 As DropDownList = e.Item.FindControl("drpPTIDEdit2")
                Dim Tech1ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech1ValueEdit")
                Dim Tech1Edit As TextBox = e.Item.FindControl("Tech1Edit")
                Dim Tech2ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech2ValueEdit")
                Dim Tech2Edit As TextBox = e.Item.FindControl("Tech2Edit")
                Dim btn3 As Button = e.Item.FindControl("Button10") 'save
                Dim btn4 As Button = e.Item.FindControl("Button11") 'cancel

                TPERIOD28_1e.Checked = False
                TPERIOD28_2e.Checked = False
                TPERIOD28_3e.Checked = False
                If Convert.ToString(drv("TPERIOD28")) <> "" _
                    AndAlso Convert.ToString(drv("TPERIOD28")).Length >= 3 Then
                    If Convert.ToString(drv("TPERIOD28")).Substring(0, 1) = "Y" Then TPERIOD28_1e.Checked = True
                    If Convert.ToString(drv("TPERIOD28")).Substring(1, 1) = "Y" Then TPERIOD28_2e.Checked = True
                    If Convert.ToString(drv("TPERIOD28")).Substring(2, 1) = "Y" Then TPERIOD28_3e.Checked = True
                End If

                Call CreateTimesItem(Eddlh1, Eddlh2, Eddlm1, Eddlm2)

                Img1.Attributes("onclick") = "return chkTrainDate('" & STrainDateTxt.ClientID & "');"

                '任課教師
                Tech1Edit.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addx','" & Tech1Edit.ClientID & "','" & Tech1ValueEdit.ClientID & "');")
                Tech1Edit.Attributes("onchange") = "GetTeacherId(this.value,'" & Tech1ValueEdit.ClientID & "','" & Tech1Edit.ClientID & "');"
                Tech1Edit.Style.Item("CURSOR") = "hand"
                '助教
                Tech2Edit.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addy','" & Tech2Edit.ClientID & "','" & Tech2ValueEdit.ClientID & "');")
                Tech2Edit.Attributes("onchange") = "GetTeacherId(this.value,'" & Tech2ValueEdit.ClientID & "','" & Tech2Edit.ClientID & "');"
                Tech2Edit.Style.Item("CURSOR") = "hand"

                If drv("STrainDate").ToString <> "" Then
                    'STrainDateTxt.Text = Common.FormatDate(drv("STrainDate").ToString)
                    STrainDateTxt.Text = TIMS.Cdate3(drv("STrainDate"))
                End If

                PNameTxt.Text = drv("PName").ToString
                If PNameTxt.Text <> "" Then
                    Try
                        PNameTxt.Text = Replace(PNameTxt.Text, "：", ":")
                        PNameTxt.Text = Replace(PNameTxt.Text, "-", "~")
                        PNameTxt.Text = TIMS.ChangeIDNO(PNameTxt.Text)
                        Dim hm1hm2 As String() = Convert.ToString(PNameTxt.Text).Split("~")
                        Dim hm1 As String()
                        Dim hm2 As String()

                        If hm1hm2.Length > 1 Then
                            If hm1hm2(0).IndexOf(":") > -1 Then
                                hm1 = hm1hm2(0).Split(":")
                                hm2 = hm1hm2(1).Split(":")
                                If hm1.Length > 1 Then
                                    Common.SetListItem(Eddlh1, Convert.ToString(hm1(0)))
                                    Common.SetListItem(Eddlm1, Convert.ToString(hm1(1)))
                                End If
                                If hm2.Length > 1 Then
                                    Common.SetListItem(Eddlh2, Convert.ToString(hm2(0)))
                                    Common.SetListItem(Eddlm2, Convert.ToString(hm2(1)))
                                End If
                            Else
                                If Convert.ToString(hm1hm2(0)).Length = 4 AndAlso IsNumeric(hm1hm2(0)) Then
                                    Common.SetListItem(Eddlh1, Convert.ToString(hm1hm2(0).Substring(0, 2)))
                                    Common.SetListItem(Eddlm1, Convert.ToString(hm1hm2(0).Substring(2, 2)))
                                End If
                                If Convert.ToString(hm1hm2(0)).Length = 4 AndAlso IsNumeric(hm1hm2(1)) Then
                                    Common.SetListItem(Eddlh2, Convert.ToString(hm1hm2(1).Substring(0, 2)))
                                    Common.SetListItem(Eddlm2, Convert.ToString(hm1hm2(1).Substring(2, 2)))
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        Dim strErrmsg As String = ""
                        strErrmsg &= "/*  ex.ToString: */" & vbCrLf
                        strErrmsg += ex.ToString & vbCrLf
                        'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
                        strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                        Call TIMS.WriteTraceLog(strErrmsg)
                    End Try
                End If

                PHourTxt.Text = drv("PHour").ToString
                PContEdit.Text = Convert.ToString(drv("PCont"))
                PContEdit.Text = TIMS.ClearSQM(PContEdit.Text)

                drpClassEdit.Attributes.Add("onchange", "javascript:showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');Layer_change(5);")
                'If TIMS.ClearSQM(Request("ComIDNO") Is Nothing Then
                'drpPTIDEdit1 = TIMS.Get_SciPTID(drpPTIDEdit1, TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID.ToString))
                'drpPTIDEdit2 = TIMS.Get_TechPTID(drpPTIDEdit2, TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID.ToString))
                'Else
                'drpPTIDEdit1 = TIMS.Get_SciPTID(drpPTIDEdit1, TIMS.ClearSQM(Request("ComIDNO"))
                'drpPTIDEdit2 = TIMS.Get_TechPTID(drpPTIDEdit2, TIMS.ClearSQM(Request("ComIDNO"))
                'End If

                drpPTIDEdit1 = getPTID(drpPTIDEdit1, 1)
                drpPTIDEdit2 = getPTID(drpPTIDEdit2, 2)

                Common.SetListItem(drpClassEdit, drv("Classification1").ToString)
                Select Case drpClassEdit.SelectedValue
                    Case "1"
                        If drv("PTID").ToString <> "" Then
                            Common.SetListItem(drpPTIDEdit1, drv("PTID").ToString)
                        End If
                    Case "2"
                        If drv("PTID").ToString <> "" Then
                            Common.SetListItem(drpPTIDEdit2, drv("PTID").ToString)
                        End If
                End Select
                Page.RegisterStartupScript("Londing3", "<script>showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');</script>")

                Tech1ValueEdit.Value = ""
                Tech1Edit.Text = ""
                If Convert.ToString(drv("TechID")) <> "" Then
                    Tech1ValueEdit.Value = drv("TechID").ToString
                    Tech1Edit.Text = TIMS.Get_TeachCName(Tech1ValueEdit.Value, objconn) '
                End If
                Tech2ValueEdit.Value = ""
                Tech2Edit.Text = ""
                If Convert.ToString(drv("TechID2")) <> "" Then
                    Tech2ValueEdit.Value = drv("TechID2").ToString
                    Tech2Edit.Text = TIMS.Get_TeachCName(Tech2ValueEdit.Value, objconn) '
                End If

                btn3.CommandArgument = drv("PTDID").ToString
                'Page.RegisterStartupScript("Londing2", "<script>Layer_change(5);showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');</script>")

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    btn3.Visible = False
                '    btn4.Visible = False
                'End If
                'strAddsObjID = btn3.ClientID & "," & btn4.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End
        End Select
    End Sub

    Private Sub SciPlaceID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SciPlaceID.SelectedIndexChanged
        If center.Text = "" Then
            Common.RespWrite(Me, "<Script>alert('請先選擇【訓練機構】');</Script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
            Exit Sub
        End If
        Dim v_SciPlaceID As String = TIMS.GetListValue(SciPlaceID)
        Taddress2 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress2, v_SciPlaceID, 1, 1, objconn)
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    End Sub

    Private Sub TechPlaceID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TechPlaceID.SelectedIndexChanged
        If center.Text = "" Then
            Common.RespWrite(Me, "<Script>alert('請先選擇【訓練機構】');</Script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
            Exit Sub
        End If
        Dim v_TechPlaceID As String = TIMS.GetListValue(TechPlaceID)
        Taddress3 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress3, v_TechPlaceID, 2, 2, objconn)
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    End Sub

    Private Sub SciPlaceID2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SciPlaceID2.SelectedIndexChanged
        If center.Text = "" Then
            Common.RespWrite(Me, "<Script>alert('請先選擇【訓練機構】');</Script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
            Exit Sub
        End If
        Dim v_SciPlaceID2 As String = TIMS.GetListValue(SciPlaceID2)
        Taddress2 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress2, v_SciPlaceID2, 3, 1, objconn)
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    End Sub

    Private Sub TechPlaceID2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TechPlaceID2.SelectedIndexChanged
        If center.Text = "" Then
            Common.RespWrite(Me, "<Script>alert('請先選擇【訓練機構】');</Script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
            Exit Sub
        End If
        Dim v_TechPlaceID2 As String = TIMS.GetListValue(TechPlaceID2)
        Taddress3 = TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, Taddress3, v_TechPlaceID2, 4, 2, objconn)
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    End Sub

    Private Sub Classification1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Classification1.SelectedIndexChanged
        Select Case Classification1.SelectedValue
            Case "1"
                PTID1 = getPTID(PTID1, 1)
            Case "2"
                PTID2 = getPTID(PTID2, 2)
            Case Else
                PTID1.Items.Clear()
                PTID2.Items.Clear()
        End Select
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    End Sub

    Function getPTID(ByVal obj As ListControl, ByVal Type As Integer) As ListControl
        Dim tempdt As DataTable
        Dim dr() As DataRow = Nothing
        Dim i As Integer
        tempdt = ViewState("dtTaddress")
        obj.Items.Clear()
        If tempdt IsNot Nothing Then
            If Type = 1 Then     '學科
                If tempdt.Select("PID IN(1,3)").Length > 0 Then
                    dr = tempdt.Select("PID IN(1,3)")
                End If
            ElseIf Type = 2 Then '術科
                If tempdt.Select("PID IN(2,4)").Length > 0 Then
                    dr = tempdt.Select("PID IN(2,4)")
                End If
            End If
            If Not dr Is Nothing Then
                For i = 0 To dr.Length - 1
                    obj.Items.Insert(i, New ListItem(dr(i)("PlaceNAME"), dr(i)("PTID")))
                Next
            End If
        End If
        Return obj
    End Function

    Sub GetNewtable()
        Dim dtSpace As New DataTable
        'Dim sql As String
        Dim dr As DataRow
        Dim i As Integer

        ViewState("dtTaddress") = Nothing
        dtSpace.Columns.Add("PID")
        dtSpace.Columns.Add("PlaceID")
        dtSpace.Columns.Add("Name")
        dtSpace.Columns.Add("classification")
        dtSpace.Columns.Add("PTID")
        dtSpace.Columns.Add("PlaceNAME")
        For i = 0 To 4
            dr = dtSpace.NewRow()
            dr("PID") = i
            dr("PlaceID") = ""
            If i = 0 Then
                dr("Name") = "======請選擇======"
            Else
                dr("Name") = ""
            End If
            dr("classification") = ""
            dr("PTID") = ""
            dr("PlaceNAME") = ""
            dtSpace.Rows.Add(dr)
        Next
        ViewState("dtTaddress") = dtSpace

    End Sub

    Private Sub Taddress2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Taddress2.SelectedIndexChanged
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    End Sub

    Private Sub Taddress3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Taddress3.SelectedIndexChanged
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
    End Sub

    '(新增)Plan_BusPackage-計畫包班事業單位
    Private Sub btnAddBusPackage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddBusPackage.Click
        Const Cst_PKName As String = "BPID"
        'Plan_BusPackage
        ' Session(cst_Plan_BusPackage)
        '------'------'------'------錯誤檢查'------'------'------'------
        Dim Errmsg As String = ""
        txtUname.Text = TIMS.ClearSQM(txtUname.Text)
        txtIntaxno.Text = TIMS.ClearSQM(txtIntaxno.Text)
        If txtUname.Text = "" Then
            txtUname.Text = ""
            Errmsg &= "企業名稱，不可為空" & vbCrLf
        Else
            '錯誤檢查
            'txtUname.Text = Trim(txtUname.Text)
            If txtUname.Text.ToString.Length > 50 Then
                Errmsg &= "企業名稱，長度超過限制範圍50文字長度" & vbCrLf
            End If
        End If
        If txtIntaxno.Text <> "" Then
            'txtIntaxno.Text = Trim(txtIntaxno.Text)
            If Not TIMS.CheckIsECFA(TIMS.ChangeIDNO(txtIntaxno.Text), objconn) Then
                '未填寫 ECFA包班事業單位資料
                Errmsg &= "「" & Convert.ToString(txtUname.Text.Trim) & "」該企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf
            End If
        Else
            txtIntaxno.Text = ""
            Errmsg &= "服務單位統一編號，不可為空" & vbCrLf
        End If

        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        '------'------'------'------錯誤檢查'------'------'------'------
        'Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow

        If Session(cst_Plan_BusPackage) Is Nothing Then
            Call CreateBusPackage()
        End If
        dt = Session(cst_Plan_BusPackage)
        dt.Columns(Cst_PKName).AutoIncrement = True
        dt.Columns(Cst_PKName).AutoIncrementSeed = -1
        dt.Columns(Cst_PKName).AutoIncrementStep = -1

        dr = dt.NewRow
        dt.Rows.Add(dr)
        If TIMS.ClearSQM(Request("PlanID")) <> "" Then
            dr("PlanID") = TIMS.ClearSQM(Request("PlanID"))
            dr("ComIDNO") = TIMS.ClearSQM(Request("ComIDNO"))
            dr("SeqNo") = TIMS.ClearSQM(Request("SeqNo"))
        End If
        dr("Uname") = Convert.ToString(txtUname.Text)
        dr("Intaxno") = TIMS.ChangeIDNO(txtIntaxno.Text)
        dr("Ubno") = TIMS.ChangeIDNO(txtUbno.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        'Datagrid4Table.Visible = True
        'Datagrid4.DataSource = dt
        'Datagrid4.DataBind()

        Session(cst_Plan_BusPackage) = dt
        Call CreateBusPackage()

        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    Private Sub Datagrid4_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid4.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Const Cst_PKName As String = "BPID"
        Dim objTable As HtmlTable = CType(Datagrid4Table, HtmlTable)
        Select Case e.CommandName
            Case "xedit"
                source.EditItemIndex = e.Item.ItemIndex
            Case "xdel"
                Dim dt As DataTable = Session(cst_Plan_BusPackage)
                Dim DGobj As DataGrid = Datagrid4
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If
                'dt = Session(cst_Plan_BusPackage)
                If dt.Select(Cst_PKName & "='" & e.CommandArgument & "'").Length <> 0 Then
                    dt.Select(Cst_PKName & "='" & e.CommandArgument & "'")(0).Delete()
                End If
                Session(cst_Plan_BusPackage) = dt
                objTable.Visible = False
                If dt.Rows.Count > 0 Then
                    objTable.Visible = True
                    source.DataSource = dt
                End If
                source.EditItemIndex = -1
            Case "xsave"
                Dim okflag As Boolean = True
                Dim tUName As TextBox = e.Item.FindControl("ttxtUName")
                Dim tIntaxno As TextBox = e.Item.FindControl("ttxtIntaxno")
                Dim tUbno As TextBox = e.Item.FindControl("ttxtUbno")
                If Session(cst_Plan_BusPackage) Is Nothing Then okflag = False
                If tUName Is Nothing Then okflag = False
                If tIntaxno Is Nothing Then okflag = False
                If tUbno Is Nothing Then okflag = False
                If Not okflag Then Exit Sub '異常離開
                Dim dt As DataTable = Session(cst_Plan_BusPackage)
                If dt Is Nothing Then okflag = False
                If Not okflag Then Exit Sub '異常離開
                If dt.Select(Cst_PKName & "='" & e.CommandArgument & "'").Length <> 0 Then
                    Dim dr As DataRow = dt.Select(Cst_PKName & "='" & e.CommandArgument & "'")(0)
                    tUName.Text = TIMS.ClearSQM(tUName.Text)
                    tIntaxno.Text = TIMS.ClearSQM(tIntaxno.Text)
                    tUbno.Text = TIMS.ClearSQM(tUbno.Text)
                    dr("Uname") = Convert.ToString(tUName.Text)
                    dr("Intaxno") = TIMS.ChangeIDNO(tIntaxno.Text)
                    dr("Ubno") = TIMS.ChangeIDNO(tUbno.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                Session(cst_Plan_BusPackage) = dt
                source.EditItemIndex = -1
            Case "xcancel"
                source.EditItemIndex = -1
        End Select

        Call CreateBusPackage()
        Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")

    End Sub

    Private Sub Datagrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid4.ItemDataBound
        Const Cst_PKName As String = "BPID"
        Dim strSechObjID As String = "" '查詢按鈕ID
        Dim strAddsObjID As String = "" '維護按鈕ID
        Dim strPrntObjID As String = "" '列印按鈕ID

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim slsbUname As Label = e.Item.FindControl("slsbUname")
                Dim slabIntaxno As Label = e.Item.FindControl("slabIntaxno")
                Dim slabUbno As Label = e.Item.FindControl("slabUbno")
                Dim Button17 As Button = e.Item.FindControl("Button17") '修改
                Dim Button18 As Button = e.Item.FindControl("Button18") '刪除
                Button17.Enabled = btnAddBusPackage.Enabled
                Button18.Enabled = btnAddBusPackage.Enabled

                slsbUname.Text = drv("Uname").ToString
                slabIntaxno.Text = drv("Intaxno").ToString
                slabUbno.Text = drv("Ubno").ToString
                Button17.CommandArgument = drv(Cst_PKName)
                Button18.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                Button18.CommandArgument = drv(Cst_PKName)

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    Button17.Visible = False
                '    Button18.Visible = False
                'End If
                'strAddsObjID = Button17.ClientID & "," & Button18.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim ttxtUname As TextBox = e.Item.FindControl("ttxtUname")
                Dim ttxtIntaxno As TextBox = e.Item.FindControl("ttxtIntaxno")
                Dim ttxtUbno As TextBox = e.Item.FindControl("ttxtUbno")
                Dim Button19 As Button = e.Item.FindControl("Button19") '儲存
                Dim Button20 As Button = e.Item.FindControl("Button20") '取消
                ttxtUname.Text = Convert.ToString(drv("Uname"))
                ttxtIntaxno.Text = Convert.ToString(drv("Intaxno"))
                ttxtUbno.Text = Convert.ToString(drv("Ubno"))

                Button19.Enabled = btnAddBusPackage.Enabled
                Button20.Enabled = btnAddBusPackage.Enabled
                Button19.CommandArgument = drv(Cst_PKName)
                Button20.CommandArgument = drv(Cst_PKName)

                '2011 功能按鈕權限控管--Start
                'If Not au.blnCanAdds Then '維護
                '    Button19.Visible = False
                '    Button20.Visible = False
                'End If
                'strAddsObjID = Button19.ClientID & "," & Button20.ClientID
                'TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
                '2011 功能按鈕權限控管--End
        End Select
    End Sub

    '訓練費用項目
    Function InputCost5(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【訓練費用項目】"

        Dim i As Integer = 0
        Dim sql2 As String = ""
        Dim CI2 As DataTable
        sql2 = "SELECT * FROM KEY_COSTITEM2 ORDER BY SORT"
        CI2 = DbAccess.GetDataTable(sql2, objconn)
        If Session(cst_CostItemTable) IsNot Nothing Then
            Dim dt2 As DataTable
            'dr / dt2@PLAN_COSTITEM
            'dr2 / CI2@KEY_COSTITEM2
            dt2 = Session(cst_CostItemTable)
            If dt2.Rows.Count > 0 Then
                For Each dr As DataRow In dt2.Rows
                    If Not dr.RowState = DataRowState.Deleted Then
                        i += 1
                    End If
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                Else
                    rst = False
                End If

                For i = 0 To CI2.Rows.Count - 1
                    'Dim dr As DataRow = Nothing
                    Dim dr2 As DataRow = CI2.Rows(i)
                    Dim iItemCost As Integer = 0
                    iItemCost = 0
                    If dt2.Select("CostID='" & dr2("CostID") & "'", Nothing, DataViewRowState.CurrentRows).Length > 0 Then
                        'Dim dr As DataRow = dt2.Select("CostID='" & dr2("CostID") & "'", Nothing, DataViewRowState.CurrentRows)(0)
                        For Each dr As DataRow In dt2.Select("CostID='" & dr2("CostID") & "'", Nothing, DataViewRowState.CurrentRows)
                            '非刪除。
                            If Not dr.RowState = DataRowState.Deleted Then
                                Dim iItemAge As Integer = 0
                                Dim iOPrice As Integer = 0
                                If Convert.ToString(dr("ItemAge")) <> "" Then
                                    iItemAge = Val(dr("ItemAge"))
                                End If
                                If Convert.ToString(dr("OPrice")) <> "" Then
                                    iOPrice = Val(dr("OPrice"))
                                End If
                                If iItemAge <> 0 OrElse iOPrice <> 0 Then
                                    iItemCost = iOPrice * iItemAge
                                End If
                                Dim TmpStr As String = ""
                                TmpStr = "" & dr2("CostName") & "：單價" & CStr(iOPrice) & "元／共" & CStr(iItemAge) & CStr(dr2("ItemCostName")) & "=" & CStr(iItemCost) & "元 "
                                rNote += TmpStr & vbCrLf

                                dr3 = dt3.NewRow
                                dt3.Rows.Add(dr3)
                                dr3("str1") = TmpStr
                            End If
                        Next
                    End If
                Next
            End If
        End If

        Return rst
    End Function

    '匯出文字 一人份材料明細
    Function InputCost6(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【一人份材料明細】"
        If Not Session(Cst_PersonCostTable) Is Nothing Then
            Dim dt As DataTable
            dt = Session(Cst_PersonCostTable)
            Dim subtotal As Integer = 0
            subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then
                        i += 1
                    End If
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If

                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    'PIXOT原子筆（0.5mm藍）：單價10元╳1支╳30人＝300元
                    Dim iPerCount As Integer = Val(dr("PerCount"))
                    Dim iPrice As Integer = Val(dr("Price"))
                    Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                    dr("Total") = (iPerCount * iTNum) '顯示重算
                    dr("subtotal") = (iPrice * iPerCount * iTNum) '顯示重算 '小計
                    subtotal += Val(dr("subtotal"))
                    Dim tmpStr As String = ""
                    tmpStr = ""
                    tmpStr &= Convert.ToString(dr("CName"))
                    tmpStr &= "(" & Convert.ToString(dr("Standard")) & ")："
                    tmpStr &= "單價" & Convert.ToString(iPrice) & "元"
                    tmpStr &= "╳" & Convert.ToString(iPerCount) & " " & Convert.ToString(dr("Unit"))
                    tmpStr &= "╳" & Convert.ToString(iTNum) & "人"
                    tmpStr &= "＝" & Convert.ToString(dr("subtotal")) & "元"
                    rNote += tmpStr & vbCrLf

                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = tmpStr
                Next

            End If
        End If
        Return rst
    End Function

    '匯出文字 共同材料明細
    Function InputCost7(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        'Const cst_t1 As String = "【共同材料明細】(以下各項計算均四捨五入至整數位)"
        Const cst_t1 As String = "【共同材料明細】"
        If Not Session(Cst_CommonCostTable) Is Nothing Then
            Dim dt As DataTable
            dt = Session(Cst_CommonCostTable)
            Dim subtotal As Integer = 0
            subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then
                        i += 1
                    End If
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If

                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    If Not dr.RowState = DataRowState.Deleted Then
                        '南x 德國圓藝剪刀（21cm 鎢鋼）：單價855元╳2支÷30人＝57元
                        Dim iPrice As Integer = Val(dr("Price"))
                        Dim iAllCount As Integer = Val(dr("AllCount"))
                        Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                        dr("subtotal") = (iPrice * iAllCount)  '小計
                        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
                        subtotal += Val(dr("subtotal"))
                        Dim tmpStr As String = ""
                        tmpStr = ""
                        tmpStr &= Convert.ToString(dr("CName"))
                        tmpStr &= "(" & Convert.ToString(dr("Standard")) & ")："
                        tmpStr &= "單價" & Convert.ToString(iPrice) & "元"
                        tmpStr &= "╳" & Convert.ToString(iAllCount) & " " & Convert.ToString(dr("Unit"))
                        'tmpStr &= "÷" & Convert.ToString(iTNum) & "人"
                        tmpStr &= "＝" & Convert.ToString(dr("subtotal")) & "元"
                        rNote += tmpStr & vbCrLf

                        dr3 = dt3.NewRow
                        dt3.Rows.Add(dr3)
                        dr3("str1") = tmpStr
                    End If
                Next

            End If
        End If
        Return rst
    End Function

    '匯出文字 教材明細
    Function InputCost8(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【教材明細】"
        If Not Session(Cst_SheetCostTable) Is Nothing Then
            Dim dt As DataTable
            dt = Session(Cst_SheetCostTable)
            Dim subtotal As Integer = 0
            subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then
                        i += 1
                    End If
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If

                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    If Not dr.RowState = DataRowState.Deleted Then
                        '南x 德國圓藝剪刀（21cm 鎢鋼）：單價855元╳2支÷30人＝57元
                        Dim iPrice As Integer = Val(dr("Price"))
                        Dim iAllCount As Integer = Val(dr("AllCount"))
                        Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                        dr("subtotal") = (iPrice * iAllCount)  '小計
                        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
                        subtotal += Val(dr("subtotal"))
                        Dim tmpStr As String = ""
                        tmpStr = ""
                        tmpStr &= Convert.ToString(dr("CName"))
                        tmpStr &= "(" & Convert.ToString(dr("Standards")) & ")："
                        tmpStr &= "單價" & Convert.ToString(iPrice) & "元"
                        tmpStr &= "╳" & Convert.ToString(iAllCount) & " " & Convert.ToString(dr("Unit"))
                        'tmpStr &= "÷" & Convert.ToString(iTNum) & "人"
                        tmpStr &= "＝" & Convert.ToString(dr("subtotal")) & "元"
                        rNote += tmpStr & vbCrLf

                        dr3 = dt3.NewRow
                        dt3.Rows.Add(dr3)
                        dr3("str1") = tmpStr
                    End If
                Next

            End If
        End If
        Return rst
    End Function

    '匯出文字 其他費用明細
    Function InputCost9(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【其他費用明細】"
        If Not Session(Cst_OtherCostTable) Is Nothing Then
            Dim dt As DataTable
            dt = Session(Cst_OtherCostTable)
            Dim subtotal As Integer = 0
            subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then
                        i += 1
                    End If
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If

                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    If Not dr.RowState = DataRowState.Deleted Then
                        '南x 德國圓藝剪刀（21cm 鎢鋼）：單價855元╳2支÷30人＝57元
                        Dim iPrice As Integer = Val(dr("Price"))
                        Dim iAllCount As Integer = Val(dr("AllCount"))
                        Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                        dr("subtotal") = (iPrice * iAllCount)  '小計
                        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
                        subtotal += Val(dr("subtotal"))
                        Dim tmpStr As String = ""
                        tmpStr = ""
                        tmpStr &= Convert.ToString(dr("CName"))
                        tmpStr &= "(" & Convert.ToString(dr("Standards")) & ")："
                        tmpStr &= "單價" & Convert.ToString(iPrice) & "元"
                        tmpStr &= "╳" & Convert.ToString(iAllCount) & " " & Convert.ToString(dr("Unit"))
                        'tmpStr &= "÷" & Convert.ToString(iTNum) & "人"
                        tmpStr &= "＝" & Convert.ToString(dr("subtotal")) & "元"
                        rNote += tmpStr & vbCrLf

                        dr3 = dt3.NewRow
                        dt3.Rows.Add(dr3)
                        dr3("str1") = tmpStr
                    End If
                Next

            End If
        End If
        Return rst
    End Function

    '匯出文字 其他說明
    Function InputNote2(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        'dt3為資料主軸
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        '加入抬頭
        Const cst_t1 As String = "【其他說明】"
        If Trim(Me.tNote2.Text) <> "" Then
            rNote = cst_t1 & vbCrLf
            dr3 = dt3.NewRow
            dt3.Rows.Add(dr3)
            dr3("str1") = cst_t1

            Dim tmpStr As String = ""
            tmpStr = ""
            tmpStr &= Me.tNote2.Text
            rNote += tmpStr & vbCrLf

            dr3 = dt3.NewRow
            dt3.Rows.Add(dr3)
            dr3("str1") = tmpStr
        End If
        Return rst
    End Function

    '回傳 rNote  dt3資料連續加入
    Function InputNote2B(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        'dt3為資料主軸
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        '加入抬頭
        Const cst_t1 As String = "【其他說明】"
        If Trim(Me.tNote2b.Text) <> "" Then
            rNote = cst_t1 & vbCrLf
            dr3 = dt3.NewRow
            dt3.Rows.Add(dr3)
            dr3("str1") = cst_t1

            Dim tmpStr As String = ""
            tmpStr = ""
            tmpStr &= Me.tNote2b.Text
            rNote += tmpStr & vbCrLf

            dr3 = dt3.NewRow
            dt3.Rows.Add(dr3)
            dr3("str1") = tmpStr
        End If
        Return rst
    End Function

    '修正 Note中的文字 (匯出文字)
    Function ChangNoteText(ByRef tmpNoteDt As DataTable) As Boolean
        Dim rst As Boolean = True '正常/false:異常
        tmpNoteDt = Nothing
        tmpNoteDt = New DataTable
        tmpNoteDt.Columns.Add(New DataColumn("str1"))

        Dim tmpNote As String = ""
        Note.Text = ""
        Select Case Me.RadioButtonList1.SelectedValue
            Case cst_學分班
                Labmsg3.Text = Cst_msgother3b
                If rst Then
                    rst = InputNote2B(tmpNote, tmpNoteDt)
                    Note.Text &= tmpNote
                End If

            Case cst_非學分班
                Labmsg3.Text = Cst_msgother3
                If rst Then
                    rst = InputCost5(tmpNote, tmpNoteDt)
                    Note.Text &= tmpNote
                End If
                If rst Then
                    rst = InputCost6(tmpNote, tmpNoteDt)
                    Note.Text &= tmpNote
                End If
                If rst Then
                    rst = InputCost7(tmpNote, tmpNoteDt)
                    Note.Text &= tmpNote
                End If

                If rst Then
                    rst = InputCost8(tmpNote, tmpNoteDt)
                    Note.Text &= tmpNote
                End If
                If rst Then
                    rst = InputCost9(tmpNote, tmpNoteDt)
                    Note.Text &= tmpNote
                End If

                If rst Then
                    rst = InputNote2(tmpNote, tmpNoteDt)
                    Note.Text &= tmpNote
                End If
        End Select
        Return rst
    End Function

#Region "NO USE"
    ''帶入訓練費用 (匯出EXCEL)
    'Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21b.Click
    '    Dim rst As Boolean = True '正常/false:異常
    '    Dim dt As DataTable
    '    rst = ChangNoteText(dt)
    '    '匯出EXCEL
    '    Page.RegisterStartupScript("window_onload", "<script language=""javascript"">Layer_change(8);</script>")
    '    If rst Then
    '        Select Case Convert.ToString(CType(sender, Button).ID)
    '            Case "Button21b"
    '                Call ExpReport3(dt)
    '        End Select
    '    End If
    'End Sub
#End Region

    '匯出 訓練費用編列說明 Response
    Sub ExpReport3(ByRef dt As DataTable)

        Dim strTitle1 As String = "" '匯出種類(1:融合式訓練辦理情形 2:融合式訓練職類統計 3:(專班)辦理情形)
        strTitle1 = "訓練費用編列說明" & ".xls"

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strTitle1, System.Text.Encoding.UTF8))
        'Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        ''套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        ''mso-number-format:"0" 
        'Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '建立抬頭
        Dim ExportStr As String = ""
        For Each dr As DataRow In dt.Rows
            If Not dr.RowState = DataRowState.Deleted Then
                '建立資料面
                ExportStr = "<tr>" & vbCrLf
                ExportStr &= "<td>" & Convert.ToString(dr("str1")) & "</td>" & vbTab
                ExportStr += "</tr>" & vbCrLf
                Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
            End If
        Next

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")
        'Response.End()
        'Dim sScript1 As String = ""
        Call TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    '新增 '材料品名項目
    Private Sub btnAddMaterial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddMaterial.Click
        Dim dt As DataTable
        'Dim dr As DataRow
        Dim Errmsg As String = ""
        '材料品名項目
        PMcName.Text = TIMS.ClearSQM(PMcName.Text)
        Dim cNAMEValue As String = "" & PMcName.Text 'Trim(Me.PMcName.Text)
        If Session(Cst_MaterialTable) Is Nothing Then
            Call CreateMaterial()
        End If
        dt = Session(Cst_MaterialTable)
        dt.Columns(Cst_PMID).AutoIncrement = True
        dt.Columns(Cst_PMID).AutoIncrementSeed = -1
        dt.Columns(Cst_PMID).AutoIncrementStep = -1

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If dt.Rows(i).Item("cNAME").ToString = cNAMEValue Then
                        Errmsg &= "該材料品名項目已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If

        If cNAMEValue = "" Then
            Errmsg &= "請輸入材料品名" & vbCrLf
        End If

        Dim HaveCostID04 As Boolean = False '材料費
        HaveCostID04 = False '查詢是否有材料費項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID04 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID04 Then
            Errmsg &= "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
        End If

        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr("cNAME") = cNAMEValue '產業人才投資方案專用
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(Cst_MaterialTable) = dt
        Call CreateMaterial()

        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")

    End Sub

    Private Sub DataGrid5_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid5.ItemCommand
        Dim Errmsg As String = ""
        Dim dt As DataTable = Session(Cst_MaterialTable)
        Dim dr As DataRow
        Errmsg = ""
        Select Case e.CommandName
            Case "EDT5" '修改
                DataGrid5.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case "DEL5" '刪除
                Dim DGobj As DataGrid = DataGrid5
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                Dim sfilter As String = Cst_PMID & "='" & DataGrid5.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                   AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then
                            dr.Delete() '刪除
                        End If
                    Next
                End If

            Case "UPD5" '更新
                Dim txtPMcNAME As TextBox = e.Item.FindControl("txtPMcNAME")
                txtPMcNAME.Text = TIMS.ClearSQM(txtPMcNAME.Text)
                Dim cNAMEValue As String = "" & Trim(txtPMcNAME.Text)
                If TIMS.CheckInput(cNAMEValue) Then
                    Common.MessageBox(Me, cst_errmsg18)
                    Exit Sub
                End If

                If Convert.ToString(DataGrid5.DataKeys(e.Item.ItemIndex)) <> "" _
                                   AndAlso dt.Select(Cst_PMID & "='" & DataGrid5.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    dr = dt.Select(Cst_PMID & "='" & DataGrid5.DataKeys(e.Item.ItemIndex) & "'")(0)
                    If dt.Select(Cst_PMID & "<>'" & DataGrid5.DataKeys(e.Item.ItemIndex) & "' AND cName='" & cNAMEValue.Replace("'", "''") & "'").Length <> 0 Then
                        Errmsg &= "該材料品名項目 已在表格中" & vbCrLf
                        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
                        Common.MessageBox(Me, Errmsg)
                        Exit Sub
                    End If
                    dr("cNAME") = cNAMEValue
                End If
                DataGrid5.EditItemIndex = -1 '還原修改列數

            Case "CLS5" '取消
                DataGrid5.EditItemIndex = -1 '還原修改列數
        End Select
        Session(Cst_MaterialTable) = dt  '要新
        CreateMaterial() '建立
        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    Private Sub DataGrid5_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid5.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labPMcNAME As Label = e.Item.FindControl("labPMcNAME")
                Dim btnDEL5 As Button = e.Item.FindControl("btnDEL5") '刪除
                Dim btnEDT5 As Button = e.Item.FindControl("btnEDT5") '修改

                'labPMcNAME.Text = "" & Convert.ToString(drv("cNAME"))
                btnDEL5.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDEL5.Enabled = btnAddMaterial.Enabled
                btnEDT5.Enabled = btnAddMaterial.Enabled

            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim txtPMcNAME As TextBox = e.Item.FindControl("txtPMcNAME")
                Dim btnUPD5 As Button = e.Item.FindControl("btnUPD5") '更新
                Dim btnCLS5 As Button = e.Item.FindControl("btnCLS5") '取消

                'txtPMcNAME.Text = "" & Convert.ToString(drv("cNAME"))
                btnUPD5.Enabled = btnAddMaterial.Enabled
                btnCLS5.Enabled = True

        End Select
    End Sub

    '新增 一人份材料明細
    Sub AddPersonCost()
        Dim dt As DataTable
        Dim Errmsg As String = ""
        Dim iItemNo As Integer = Val(Me.tItemNo6.Text)
        Dim sCName As String = "" & TIMS.ClearSQM(Me.tCName6.Text) 'Trim(Me.tCName6.Text)
        Dim sStandard As String = "" & TIMS.ClearSQM(Me.tStandard6.Text) 'Trim(Me.tStandard6.Text)
        Dim sUnit As String = "" & TIMS.ClearSQM(Me.tUnit6.Text) 'Trim(Me.tUnit6.Text)
        Dim iPrice As Integer = Val(Me.tPrice6.Text)
        Dim iPerCount As Integer = Val(Me.tPerCount6.Text)
        Dim iTNum As Integer = Val(Me.TNum.Text) '取得外部資料
        Dim sPurpose As String = "" & TIMS.ClearSQM(Me.tPurpose6.Text) 'Trim(Me.tPurpose6.Text)
        If Session(Cst_PersonCostTable) Is Nothing Then
            dt = CreatePersonCost()
        Else
            '有資料
            dt = Session(Cst_PersonCostTable)
            dt.Columns(Cst_PersonCostpkName).AutoIncrement = True
            dt.Columns(Cst_PersonCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_PersonCostpkName).AutoIncrementStep = -1
        End If

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        If iTNum = 0 Then
            'CHECK:1
            Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf
        End If
        Dim HaveCostID04 As Boolean = False '材料費
        HaveCostID04 = False '查詢是否有材料費項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID04 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID04 Then
            Errmsg &= "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
        End If
        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入品名" & vbCrLf
            End If
            If sStandard = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iPerCount = 0 Then
                Errmsg &= "請輸入每人數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standard") = sStandard '300
        dr("Unit") = sUnit '30
        dr("Price") = iPrice 'int
        dr("PerCount") = iPerCount 'int
        dr("TNum") = iTNum 'int
        dr("Total") = (iPerCount * iTNum) '顯示重算
        dr("subtotal") = (iPrice * iPerCount * iTNum) '顯示重算
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(Cst_PersonCostTable) = dt
        Call CreatePersonCost()

        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")

    End Sub

    '新增 共同材料明細
    Sub AddCommonCost()
        Dim dt As DataTable
        Dim Errmsg As String = ""
        Dim iItemNo As Integer = Val(Me.tItemNo7.Text)
        Dim sCName As String = "" & Trim(Me.tCName7.Text)
        Dim sStandard As String = "" & Trim(Me.tStandard7.Text)
        Dim sUnit As String = "" & Trim(Me.tUnit7.Text)
        Dim iPrice As Integer = Val(Me.tPrice7.Text)
        Dim iAllCount As Integer = Val(Me.tAllCount7.Text)
        Dim iTNum As Integer = Val(Me.TNum.Text)   '取得外部資料
        Dim sPurpose As String = "" & Trim(Me.tPurPose7.Text)
        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
        Dim ieachCost As Integer = 0
        If iTNum > 0 Then
            ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum)) '每人分攤費用
        Else
            ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / 1)) '每人分攤費用
        End If

        If Session(Cst_CommonCostTable) Is Nothing Then
            dt = CreateCommonCost()
        Else
            '有資料
            dt = Session(Cst_CommonCostTable)
            dt.Columns(Cst_CommonCostpkName).AutoIncrement = True
            dt.Columns(Cst_CommonCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_CommonCostpkName).AutoIncrementStep = -1
        End If

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        If iTNum = 0 Then
            'CHECK:1
            Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf
        End If
        Dim HaveCostID04 As Boolean = False '材料費
        HaveCostID04 = False '查詢是否有材料費項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID04 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID04 Then
            Errmsg &= "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
        End If
        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入品名" & vbCrLf
            End If
            If sStandard = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iAllCount = 0 Then
                Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
            If ieachCost < 1 Then
                Errmsg &= "計算後每人分攤費用，不可小於1" & vbCrLf 'int
            End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standard") = sStandard '300
        dr("Unit") = sUnit '30
        dr("Price") = iPrice 'int
        dr("AllCount") = iAllCount 'int
        dr("TNum") = iTNum 'int
        dr("subtotal") = (iPrice * iAllCount) '顯示重算
        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(Cst_CommonCostTable) = dt
        Call CreateCommonCost()

        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")

    End Sub

    '修改 一人份材料明細
    Function chkdg6(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo6")
        Dim eCName As TextBox = e.Item.FindControl("eCName6")
        Dim eStandard As TextBox = e.Item.FindControl("eStandard6")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit6")
        Dim ePrice As TextBox = e.Item.FindControl("ePrice6")
        Dim ePerCount As TextBox = e.Item.FindControl("ePerCount6")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum6") '訓練人數
        Dim eTotal As TextBox = e.Item.FindControl("eTotal6") '總數量 = val(ePerCount6.text)* val(eTNum6.text)
        Dim esubtotal As TextBox = e.Item.FindControl("esubtotal6") '小計 = val(ePrice6.text)*val(ePerCount6.text)* val(eTNum6.text)
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose6")

        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & Trim(eCName.Text)
        Dim sStandard As String = "" & Trim(eStandard.Text)
        Dim sUnit As String = "" & Trim(eUnit.Text)
        Dim iPrice As Integer = Val(ePrice.Text)
        Dim iPerCount As Integer = Val(ePerCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)
        Dim sPurpose As String = "" & Trim(ePurPose.Text)
        '取得外部資料
        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

        Dim Errmsg As String = ""
        Errmsg = ""

        Dim HaveCostID04 As Boolean = False '材料費
        HaveCostID04 = False '查詢是否有材料費項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID04 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID04 Then
            Errmsg &= "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
        End If

        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入品名" & vbCrLf
            End If
            If sStandard = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iPerCount = 0 Then
                Errmsg &= "請輸入每人數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
            If Not TIMS.IsNumeric2(eItemNo.Text) Then
                Errmsg &= "項次格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(ePrice.Text) Then
                Errmsg &= "單價格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(ePerCount.Text) Then
                Errmsg &= "每人數量格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid6
                Dim dt As DataTable = Session(Cst_PersonCostTable)
                Dim sfilter As String = ""
                sfilter = ""
                sfilter &= "" & Cst_PersonCostpkName & "<>'" & Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) & "'"
                sfilter &= " AND ItemNo='" & eItemNo.Text & "'"
                'dt update
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                    Errmsg &= "[" & eItemNo.Text & "]該項次 已在表格中" & vbCrLf
                End If
                dt = Nothing
            End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            rst = False            'Exit Function
        End If
        Return rst
    End Function

    '修改 共同材料明細
    Function chkdg7(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo7")
        Dim eCName As TextBox = e.Item.FindControl("eCName7")
        Dim eStandard As TextBox = e.Item.FindControl("eStandard7")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit7")
        Dim ePrice As TextBox = e.Item.FindControl("ePrice7")
        Dim eAllCount As TextBox = e.Item.FindControl("eAllCount7")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum7") '訓練人數
        Dim esubtotal As TextBox = e.Item.FindControl("esubtotal7") '小計
        'Dim eeachCost As TextBox = e.Item.FindControl("eachCost7") '每人分攤費用　
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose7")

        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & Trim(eCName.Text)
        Dim sStandard As String = "" & Trim(eStandard.Text)
        Dim sUnit As String = "" & Trim(eUnit.Text)
        Dim iPrice As Integer = Val(ePrice.Text)
        Dim iAllCount As Integer = Val(eAllCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
        Dim ieachCost As Integer = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
        Dim sPurpose As String = "" & Trim(ePurPose.Text)
        '取得外部資料
        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

        Dim Errmsg As String = ""
        Errmsg = ""

        Dim HaveCostID04 As Boolean = False '材料費
        HaveCostID04 = False '查詢是否有材料費項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID04 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID04 Then
            Errmsg &= "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
        End If

        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入品名" & vbCrLf
            End If
            If sStandard = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iAllCount = 0 Then
                Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
            If ieachCost < 1 Then
                Errmsg &= "計算後每人分攤費用，不可小於1" & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(eItemNo.Text) Then
                Errmsg &= "項次格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(ePrice.Text) Then
                Errmsg &= "單價格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(eAllCount.Text) Then
                Errmsg &= "使用數量格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid7
                Dim dt As DataTable = Session(Cst_CommonCostTable)
                Dim sfilter As String = ""
                sfilter = ""
                sfilter &= "" & Cst_CommonCostpkName & "<>'" & Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) & "'"
                sfilter &= " AND ItemNo='" & eItemNo.Text & "'"
                'dt update
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                    Errmsg &= "[" & eItemNo.Text & "]該項次已在表格中" & vbCrLf
                End If
                dt = Nothing
            End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            rst = False            'Exit Function
        End If
        Return rst
    End Function

    Private Sub DataGrid6_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid6.ItemCommand
        Dim Errmsg As String = ""
        If Session(Cst_PersonCostTable) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid6
        Dim dt As DataTable = Session(Cst_PersonCostTable)
        Dim dr As DataRow = Nothing
        Errmsg = ""
        Select Case e.CommandName
            Case "EDT6" '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case "DEL6" '刪除
                'Dim dt As DataTable = Session(Cst_PersonCostTable)
                'Dim DGobj As DataGrid = DataGrid6
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                Dim sfilter As String = "" & Cst_PersonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                   AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then
                            dr.Delete() '刪除
                        End If
                    Next
                End If
            Case "UPD6" '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo6")
                Dim eCName As TextBox = e.Item.FindControl("eCName6")
                Dim eStandard As TextBox = e.Item.FindControl("eStandard6")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit6")
                Dim ePrice As TextBox = e.Item.FindControl("ePrice6")
                Dim ePerCount As TextBox = e.Item.FindControl("ePerCount6")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum6") '訓練人數
                Dim eTotal As TextBox = e.Item.FindControl("eTotal6") '總數量 = val(ePerCount6.text)* val(eTNum6.text)
                Dim esubtotal As TextBox = e.Item.FindControl("esubtotal6") '小計 = val(ePrice6.text)*val(ePerCount6.text)* val(eTNum6.text)
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose6")

                If chkdg6(e) Then
                    Dim sfilter As String = "" & Cst_PersonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    'dt update
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                                    AndAlso dt.Select(sfilter).Length > 0 Then
                        Dim iPrice As Integer = Val(ePrice.Text)
                        Dim iPerCount As Integer = Val(ePerCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)
                        '取得外部資料
                        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

                        dr = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standard") = eStandard.Text
                        dr("Unit") = eUnit.Text
                        dr("Price") = iPrice
                        dr("PerCount") = iPerCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("Total") = iPerCount * iTNum '顯示重算
                        dr("subtotal") = iPrice * iPerCount * iTNum '顯示重算
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If
            Case "CLS6" '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select

        Session(Cst_PersonCostTable) = dt  '要新  
        CreatePersonCost() '建立  
        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    Private Sub DataGrid6_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid6.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost6.Enabled
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                Dim lItemNo6 As Label = e.Item.FindControl("lItemNo6")
                Dim lCName6 As Label = e.Item.FindControl("lCName6")
                Dim lStandard6 As Label = e.Item.FindControl("lStandard6")
                Dim lUnit6 As Label = e.Item.FindControl("lUnit6")
                Dim lPrice6 As Label = e.Item.FindControl("lPrice6")
                Dim lPerCount6 As Label = e.Item.FindControl("lPerCount6")

                Dim lTNum6 As Label = e.Item.FindControl("lTNum6")
                Dim lTotal6 As Label = e.Item.FindControl("lTotal6")
                Dim lsubtotal6 As Label = e.Item.FindControl("lsubtotal6")

                Dim lPurPose6 As Label = e.Item.FindControl("lPurPose6")
                Dim btnDEL6 As Button = e.Item.FindControl("btnDEL6") '刪除
                Dim btnEDT6 As Button = e.Item.FindControl("btnEDT6") '修改

                lItemNo6.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName6.Text = "" & Convert.ToString(drv("CName"))
                lStandard6.Text = "" & Convert.ToString(drv("Standard"))
                lUnit6.Text = "" & Convert.ToString(drv("Unit"))
                lPrice6.Text = "" & Convert.ToString(drv("Price"))
                lPerCount6.Text = "" & Convert.ToString(drv("PerCount"))
                'dr("TNum") = Val(eTNum6.Text)
                'dr("Total") = Val(ePerCount6.Text) * Val(eTNum6.Text)
                'dr("subtotal") = Val(ePrice6.Text) * Val(ePerCount6.Text) * Val(eTNum6.Text)
                lTNum6.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                lTotal6.Text = "" & (Val(drv("PerCount")) * Val(drv("TNum"))) '顯示重算
                lsubtotal6.Text = "" & (Val(drv("Price")) * Val(drv("PerCount")) * Val(drv("TNum"))) '顯示重算
                lPurPose6.Text = "" & Convert.ToString(drv("PurPose"))

                btnDEL6.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDEL6.Enabled = Flag_AddEnabled
                btnEDT6.Enabled = Flag_AddEnabled

            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim tlItemNo6 As TextBox = e.Item.FindControl("eItemNo6")
                Dim tlCName6 As TextBox = e.Item.FindControl("eCName6")
                Dim tlStandard6 As TextBox = e.Item.FindControl("eStandard6")
                Dim tlUnit6 As TextBox = e.Item.FindControl("eUnit6")
                Dim tlPrice6 As TextBox = e.Item.FindControl("ePrice6")
                Dim tlPerCount6 As TextBox = e.Item.FindControl("ePerCount6")
                Dim tlTNum6 As TextBox = e.Item.FindControl("eTNum6")
                Dim tlTotal6 As TextBox = e.Item.FindControl("eTotal6")
                Dim tlsubtotal6 As TextBox = e.Item.FindControl("esubtotal6")
                Dim tlPurPose6 As TextBox = e.Item.FindControl("ePurPose6")

                Dim btnUPD6 As Button = e.Item.FindControl("btnUPD6") '更新
                Dim btnCLS6 As Button = e.Item.FindControl("btnCLS6") '取消

                tlItemNo6.Text = "" & Convert.ToString(drv("ItemNo"))
                tlCName6.Text = "" & Convert.ToString(drv("CName"))
                tlStandard6.Text = "" & Convert.ToString(drv("Standard"))
                tlUnit6.Text = "" & Convert.ToString(drv("Unit"))
                tlPrice6.Text = "" & Convert.ToString(drv("Price"))
                tlPerCount6.Text = "" & Convert.ToString(drv("PerCount"))
                'dr("TNum") = Val(eTNum6.Text)
                'dr("Total") = Val(ePerCount6.Text) * Val(eTNum6.Text)
                'dr("subtotal") = Val(ePrice6.Text) * Val(ePerCount6.Text) * Val(eTNum6.Text)
                tlTNum6.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                tlTotal6.Text = "" & (Val(drv("PerCount")) * Val(drv("TNum"))) '顯示重算
                tlsubtotal6.Text = "" & (Val(drv("Price")) * Val(drv("PerCount")) * Val(drv("TNum"))) '顯示重算
                tlPurPose6.Text = "" & Convert.ToString(drv("PurPose"))
                tlTNum6.ReadOnly = True
                tlTotal6.ReadOnly = True
                tlsubtotal6.ReadOnly = True
                'tlTNum6.Style.Item("background-color") = "#FFECEC"
                'tlTotal6.Style.Item("background-color") = "#FFECEC"
                'tlsubtotal6.Style.Item("background-color") = "#FFECEC"
                tlTNum6.Style.Item("background-color") = "#BDBDBD"
                tlTotal6.Style.Item("background-color") = "#BDBDBD"
                tlsubtotal6.Style.Item("background-color") = "#BDBDBD"

                btnUPD6.Enabled = Flag_AddEnabled
                btnCLS6.Enabled = True

        End Select
    End Sub

    '新增 'Plan_PersonCost–一人份材料明細 
    Private Sub btnAddCost6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCost6.Click
        Call AddPersonCost()
    End Sub

    '匯入明細 'Plan_PersonCost–一人份材料明細
    Protected Sub BtnImport1_Click(sender As Object, e As EventArgs) Handles BtnImport1.Click
        Dim Errmsg As String = ""
        If File1_test(Errmsg) Then
            '顯示 內容
            Call CreatePersonCost()
        Else
            Common.MessageBox(Me, Errmsg)
        End If
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    '新增 'Plan_CommonCost–共同材料明細
    Private Sub btnAddCost7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCost7.Click
        Call AddCommonCost()
    End Sub

    '匯入明細  'Plan_CommonCost–共同材料明細
    Protected Sub BtnImport2_Click(sender As Object, e As EventArgs) Handles BtnImport2.Click
        Dim Errmsg As String = ""
        If File2_test(Errmsg) Then
            '顯示 內容
            Call CreateCommonCost()
        Else
            Common.MessageBox(Me, Errmsg)
        End If
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    Private Sub DataGrid7_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid7.ItemCommand
        Dim Errmsg As String = ""
        If Session(Cst_CommonCostTable) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid7
        Dim dt As DataTable = Session(Cst_CommonCostTable)
        Dim dr As DataRow = Nothing
        Errmsg = ""
        Select Case e.CommandName
            Case "EDT7" '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case "DEL7" '刪除
                'Dim dt As DataTable = Session(Cst_CommonCostTable)
                'Dim DGobj As DataGrid = DataGrid7
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                Dim sfilter As String = "" & Cst_CommonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                   AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then
                            dr.Delete() '刪除
                        End If
                    Next
                End If
            Case "UPD7" '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo7")
                Dim eCName As TextBox = e.Item.FindControl("eCName7")
                Dim eStandard As TextBox = e.Item.FindControl("eStandard7")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit7")
                Dim ePrice As TextBox = e.Item.FindControl("ePrice7")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount7")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum7") '訓練人數
                Dim esubtotal As TextBox = e.Item.FindControl("esubtotal7") '小計 = val(ePrice6.text)*val(ePerCount6.text)* val(eTNum6.text)
                'Dim eeachCost As TextBox = e.Item.FindControl("eachCost7") '總數量 = val(ePerCount6.text)* val(eTNum6.text)
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose7")

                If chkdg7(e) Then
                    Dim sfilter As String = "" & Cst_CommonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    'dt update
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                                    AndAlso dt.Select(sfilter).Length > 0 Then
                        Dim iPrice As Integer = Val(ePrice.Text)
                        Dim iAllCount As Integer = Val(eAllCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
                        '取得外部資料
                        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)
                        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
                        Dim ieachCost As Integer = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用

                        dr = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standard") = eStandard.Text
                        dr("Unit") = eUnit.Text
                        dr("Price") = iPrice
                        dr("AllCount") = iAllCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("subtotal") = isubtotal '小計 '顯示重算
                        dr("eachCost") = ieachCost  '每人分攤費用 '顯示重算
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If

            Case "CLS7" '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select

        Session(Cst_CommonCostpkName) = dt  '要新  
        CreateCommonCost() '建立  
        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")

    End Sub

    Private Sub DataGrid7_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid7.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost7.Enabled
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                Dim lItemNo As Label = e.Item.FindControl("lItemNo7")
                Dim lCName As Label = e.Item.FindControl("lCName7")
                Dim lStandard As Label = e.Item.FindControl("lStandard7")
                Dim lUnit As Label = e.Item.FindControl("lUnit7")
                Dim lPrice As Label = e.Item.FindControl("lPrice7")
                Dim lAllCount As Label = e.Item.FindControl("lAllCount7")

                Dim lTNum As Label = e.Item.FindControl("lTNum7")
                Dim lsubtotal As Label = e.Item.FindControl("lsubtotal7")
                Dim leachCost As Label = e.Item.FindControl("leachCost7")

                Dim lPurPose As Label = e.Item.FindControl("lPurPose7")
                Dim btnDEL As Button = e.Item.FindControl("btnDEL7") '刪除
                Dim btnEDT As Button = e.Item.FindControl("btnEDT7") '修改

                lItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName.Text = "" & Convert.ToString(drv("CName"))
                lStandard.Text = "" & Convert.ToString(drv("Standard"))
                lUnit.Text = "" & Convert.ToString(drv("Unit"))
                'lPrice.Text = "" & Convert.ToString(drv("Price"))
                'lAllCount.Text = "" & Convert.ToString(drv("AllCount"))
                'lTNum.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                'lsubtotal.Text = "" & (Val(drv("Price")) * Val(drv("AllCount")))  '顯示重算
                'leachCost.Text = "" & ((Val(drv("Price")) * Val(drv("AllCount"))) / Val(drv("TNum")))  '顯示重算
                Dim iPrice As Integer = Val(drv("Price"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                '取得外部資料
                If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

                lPrice.Text = iPrice
                lAllCount.Text = iAllCount
                lTNum.Text = iTNum '顯示原資料
                lsubtotal.Text = (iPrice * iAllCount)  '小計  '顯示重算
                leachCost.Text = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算

                lPurPose.Text = "" & Convert.ToString(drv("PurPose"))

                btnDEL.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDEL.Enabled = Flag_AddEnabled
                btnEDT.Enabled = Flag_AddEnabled

            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim tlItemNo As TextBox = e.Item.FindControl("eItemNo7")
                Dim tlCName As TextBox = e.Item.FindControl("eCName7")
                Dim tlStandard As TextBox = e.Item.FindControl("eStandard7")
                Dim tlUnit As TextBox = e.Item.FindControl("eUnit7")
                Dim tlPrice As TextBox = e.Item.FindControl("ePrice7")
                Dim tlAllCount As TextBox = e.Item.FindControl("eAllCount7")
                Dim tlTNum As TextBox = e.Item.FindControl("eTNum7")
                Dim tlsubtotal As TextBox = e.Item.FindControl("esubtotal7")
                Dim tleachCost As TextBox = e.Item.FindControl("eeachCost7")
                Dim tlPurPose As TextBox = e.Item.FindControl("ePurPose7")

                Dim btnUPD As Button = e.Item.FindControl("btnUPD7") '更新
                Dim btnCLS As Button = e.Item.FindControl("btnCLS7") '取消

                tlItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                tlCName.Text = "" & Convert.ToString(drv("CName"))
                tlStandard.Text = "" & Convert.ToString(drv("Standard"))
                tlUnit.Text = "" & Convert.ToString(drv("Unit"))
                'dr("TNum") = Val(eTNum6.Text)
                'dr("Total") = Val(ePerCount6.Text) * Val(eTNum6.Text)
                'dr("subtotal") = Val(ePrice6.Text) * Val(ePerCount6.Text) * Val(eTNum6.Text)
                'tlTNum.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                'tlsubtotal.Text = "" & (Val(drv("Price")) * Val(drv("AllCount")))  '顯示重算
                'tleachCost.Text = "" & ((Val(drv("Price")) * Val(drv("AllCount"))) / Val(drv("TNum")))  '顯示重算
                Dim iPrice As Integer = Val(drv("Price"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                '取得外部資料
                If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

                tlPrice.Text = iPrice
                tlAllCount.Text = iAllCount
                tlTNum.Text = iTNum '顯示原資料
                tlsubtotal.Text = (iPrice * iAllCount)  '小計  '顯示重算
                tleachCost.Text = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算

                tlPurPose.Text = "" & Convert.ToString(drv("PurPose"))

                tlTNum.ReadOnly = True
                tlsubtotal.ReadOnly = True
                tleachCost.ReadOnly = True
                'tlTNum6.Style.Item("background-color") = "#FFECEC"
                'tlTotal6.Style.Item("background-color") = "#FFECEC"
                'tlsubtotal6.Style.Item("background-color") = "#FFECEC"
                tlTNum.Style.Item("background-color") = "#BDBDBD"
                tlsubtotal.Style.Item("background-color") = "#BDBDBD"
                tleachCost.Style.Item("background-color") = "#BDBDBD"

                btnUPD.Enabled = Flag_AddEnabled
                btnCLS.Enabled = True
        End Select
    End Sub

    Function sUtl_UpdateNote(ByRef tConn As SqlConnection) As Boolean
        Dim rst As Boolean = True '正常/false:異常
        Call TIMS.OpenDbConn(tConn)
        If Me.Note.Text <> "" Then
            Dim sql As String = ""
            Dim da As SqlDataAdapter = Nothing
            Dim dt As DataTable = Nothing
            Try
                Dim dr As DataRow = Nothing
                If Me.upt_PlanX.Value <> "" Then '有儲存資料過了
                    '有儲存資料過了
                    '準備儲存資料
                    tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                    PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                    ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                    SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                    sql = "SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                    dt = DbAccess.GetDataTable(sql, da, tConn)
                    dr = dt.Rows(0)
                Else
                    If (Convert.ToString(Request("PlanID")) = "" OrElse Convert.ToString(Request(cst_ccopy)) = "1") Then
                        '新增資料 、copy=1 、草稿新增 而來
                        'Call TIMS.CloseDbConn(conn)
                        Common.MessageBox(Me, cst_errmsg19)
                        rst = False
                        Return False
                        'Exit Function
                    Else
                        '修改
                        PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                        ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                        SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                        sql = "select * from Plan_PlanInfo where PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
                        dt = DbAccess.GetDataTable(sql, da, tConn)
                        dr = dt.Rows(0)
                    End If

                    tmpPCS = ""
                    TIMS.SetMyValue(tmpPCS, "PlanID", PlanID_value)
                    TIMS.SetMyValue(tmpPCS, "ComIDNO", ComIDNO_value)
                    TIMS.SetMyValue(tmpPCS, "SeqNO", SeqNO_value)
                    Me.upt_PlanX.Value = tmpPCS
                End If
                dr("Note") = Me.Note.Text
                'DbAccess.UpdateDataTable(dt, da, Trans)
                DbAccess.UpdateDataTable(dt, da)
                'DbAccess.CommitTrans(Trans)
            Catch ex As Exception
                Me.upt_PlanX.Value = ""
                'DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(tConn)
                Common.MessageBox(Me, cst_errmsg6)
                Common.MessageBox(Me, ex.ToString)
                rst = False
                Return False
                'Exit Function
            End Try

            'Call TIMS.CloseDbConn(conn)
        End If
        Return rst
    End Function

    '匯出EXCEL
    Private Sub Button21b_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21b.Click
        Dim rst As Boolean = True '正常/false:異常
        '將資料寫入Note
        rst = ChangNoteText(tmpNoteDt)
        '匯出EXCEL
        Page.RegisterStartupScript("window_onload", "<script language=""javascript"">Layer_change(8);</script>")

        If rst Then
            '只修改 Note儲存格
            rst = sUtl_UpdateNote(objconn) '<==內部有顯示錯誤訊息
            If Not rst Then Exit Sub '異常離開
        End If
        If Not rst Then
            Common.MessageBox(Me, cst_errmsg20)
            Exit Sub
        End If

        rst = False '異常 True:正常
        Dim dt As DataTable
        Dim dt1 As DataTable
        Dim dt2 As DataTable
        Dim sql As String = ""
        PlanID_value = TIMS.ClearSQM(Request("PlanID"))
        ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
        SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
        sql = "SELECT * FROM Plan_PlanInfo WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
        dt = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT * FROM Plan_PersonCost WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
        dt1 = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT * FROM Plan_CommonCost WHERE PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
        dt2 = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 1 Then
            If dt1.Rows.Count > 0 OrElse dt2.Rows.Count > 0 Then
                rst = True
            End If
        End If
        If Not rst Then
            Common.MessageBox(Me, Cst_msgother3)
            Exit Sub
        End If

        PlanID_value = TIMS.ClearSQM(Request("PlanID"))
        ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
        SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
        Dim MyValue As String = ""
        MyValue = ""
        MyValue = "YEARS=" & Convert.ToString(sm.UserInfo.Years - 1911)
        MyValue += "&PLANID=" & PlanID_value
        MyValue += "&ComIDNO=" & ComIDNO_value
        MyValue += "&SEQNO=" & SeqNO_value
        MyValue += "&PCSValue=" & PlanID_value & "x" & ComIDNO_value & "x" & SeqNO_value
        '材料明細表
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "BussinessTrain", "SD_14_020", MyValue)
    End Sub

    Private Sub btnUptNote2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUptNote2.Click
        Call ChangNoteText(tmpNoteDt)
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    Private Sub btnUptNote2b_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUptNote2b.Click
        Call ChangNoteText(tmpNoteDt)
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
    End Sub

    '匯入 一人份材料明細
    Function File1_test(ByRef rErrmsg As String) As Boolean
        '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim rst As Boolean = False

        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Dim Reason As String = ""               '儲存錯誤的原因
        'Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        'Dim drWrong As DataRow
        ''建立錯誤資料格式Table----------------Start
        'dtWrong.Columns.Add(New DataColumn("Index"))
        'dtWrong.Columns.Add(New DataColumn("PName"))
        'dtWrong.Columns.Add(New DataColumn("IDNO"))
        'dtWrong.Columns.Add(New DataColumn("Reason"))
        ''建立錯誤資料格式Table----------------End        

        'Dim MyFile As System.IO.File
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        Dim flag As String = ","

        Dim dt As DataTable
        If Session(Cst_PersonCostTable) Is Nothing Then
            dt = CreatePersonCost()
        Else
            '有資料
            dt = Session(Cst_PersonCostTable)
            dt.Columns(Cst_PersonCostpkName).AutoIncrement = True
            dt.Columns(Cst_PersonCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_PersonCostpkName).AutoIncrementStep = -1
        End If

        If File1.Value <> "" Then
            '檢查檔案格式與大小----------   Start
            If File1.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Return rst 'Exit Function
            Else
                '取出檔案名稱
                MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    Return rst 'Exit Function
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If Not LCase(MyFileType) = "csv" Then
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        Return rst 'Exit Function
                    End If
                End If
            End If
            '檢查檔案格式與大小----------   End

            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            '上傳檔案
            File1.PostedFile.SaveAs(Server.MapPath(Upload_Path & MyFileName))

            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(Server.MapPath(Upload_Path & MyFileName))
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0 '讀取行累計數
            Dim OneRow As String        'srr.ReadLine 一行一行的資料
            'Dim col As String           '欄位
            Dim colArray As Array

            Do While srr.Peek >= 0
                If Reason <> "" Then Exit Do
                OneRow = srr.ReadLine
                If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈

                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, flag)
                    If colArray.Length < 7 Then
                        rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                        rst = False
                        Exit Do
                    End If
                    Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                    Dim sCName As String = "" & Trim(colArray(1).ToString) '品名
                    Dim sStandard As String = "" & Trim(colArray(2).ToString) '規格
                    Dim sUnit As String = "" & Trim(colArray(3).ToString) '單位
                    Dim iPrice As Integer = Val(colArray(4).ToString) '單價
                    Dim iPerCount As Integer = Val(colArray(5).ToString)  '每人數量
                    Dim iTNum As Integer = Val(Me.TNum.Text) '取得外部資料(不可為0)
                    Dim sPurpose As String = "" & Trim(colArray(6).ToString) '用途說明

                    'If Reason = "" Then Reason += CheckImportData(colArray) '檢查資料正確性
                    '檢查資料正確性
                    If Reason = "" Then

                        If dt.Rows.Count > 0 Then
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                        Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If iTNum = 0 Then
                            'CHECK:1
                            Reason += "請先輸入訓練人數，不可為0" & vbCrLf
                        End If

                        Dim HaveCostID04 As Boolean = False '材料費
                        HaveCostID04 = False '查詢是否有材料費項目
                        If Not Session(cst_CostItemTable) Is Nothing Then
                            Dim dt2 As DataTable
                            dt2 = Session(cst_CostItemTable)
                            For i As Int16 = 0 To dt2.Rows.Count - 1
                                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                                    AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                                    HaveCostID04 = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not HaveCostID04 Then
                            Reason += "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
                        End If
                        If Reason = "" Then
                            If iItemNo = 0 Then
                                Reason += "請輸入項次，不可為0" & vbCrLf 'int
                            End If
                            If sCName = "" Then
                                Reason += "請輸入品名" & vbCrLf
                            End If
                            If sStandard = "" Then
                                Reason += "請輸入規格" & vbCrLf
                            End If
                            If sUnit = "" Then
                                Reason += "請輸入單位" & vbCrLf
                            End If
                            If iPrice = 0 Then
                                Reason += "請輸入單價，不可為0" & vbCrLf 'int
                            End If
                            If iPerCount = 0 Then
                                Reason += "請輸入每人數量，不可為0" & vbCrLf 'int
                            End If
                            If sPurpose = "" Then
                                Reason += "請輸入用途說明" & vbCrLf
                            End If
                        End If
                    End If

                    If Reason = "" Then
                        '產業人才投資方案專用
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)

                        dr("ItemNo") = iItemNo 'int
                        dr("CName") = sCName '30
                        dr("Standard") = sStandard '300
                        dr("Unit") = sUnit '30
                        dr("Price") = iPrice 'int
                        dr("PerCount") = iPerCount 'int
                        dr("TNum") = iTNum 'int
                        dr("Total") = (iPerCount * iTNum) '顯示重算
                        dr("subtotal") = (iPrice * iPerCount * iTNum) '顯示重算
                        dr("PurPose") = sPurpose '300
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now

                        Session(Cst_PersonCostTable) = dt
                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            sr.Close()
            srr.Close()
            IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        Else
            Reason += "請選擇匯入的檔案" & vbCrLf
        End If

        If Reason <> "" Then
            rErrmsg = Reason
            rst = False
        Else
            rst = True
        End If
        Return rst
    End Function

    '匯入 共同材料明細
    Function File2_test(ByRef rErrmsg As String) As Boolean
        '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim rst As Boolean = False

        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Dim Reason As String = ""               '儲存錯誤的原因
        'Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        'Dim drWrong As DataRow
        ''建立錯誤資料格式Table----------------Start
        'dtWrong.Columns.Add(New DataColumn("Index"))
        'dtWrong.Columns.Add(New DataColumn("PName"))
        'dtWrong.Columns.Add(New DataColumn("IDNO"))
        'dtWrong.Columns.Add(New DataColumn("Reason"))
        ''建立錯誤資料格式Table----------------End        

        'Dim MyFile As System.IO.File
        Dim MyFileName, MyFileType As String
        Dim flag As String = ","

        Dim dt As DataTable
        If Session(Cst_CommonCostTable) Is Nothing Then
            dt = CreateCommonCost()
        Else
            '有資料
            dt = Session(Cst_CommonCostTable)
            dt.Columns(Cst_CommonCostpkName).AutoIncrement = True
            dt.Columns(Cst_CommonCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_CommonCostpkName).AutoIncrementStep = -1
        End If

        Dim oFile As HtmlInputFile = File2
        If oFile.Value <> "" Then
            '檢查檔案格式與大小----------   Start
            If File2.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Return rst
                'Exit Function
            Else
                '取出檔案名稱
                MyFileName = Split(oFile.PostedFile.FileName, "\")((Split(oFile.PostedFile.FileName, "\")).Length - 1)
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    Return rst
                    'Exit Function
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If Not LCase(MyFileType) = "csv" Then
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        Return rst
                        'Exit Function
                    End If
                End If
            End If
            '檢查檔案格式與大小----------   End

            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            '上傳檔案
            oFile.PostedFile.SaveAs(Server.MapPath(Upload_Path & MyFileName))

            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(Server.MapPath(Upload_Path & MyFileName))
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0 '讀取行累計數
            Dim OneRow As String        'srr.ReadLine 一行一行的資料
            'Dim col As String           '欄位
            Dim colArray As Array

            Do While srr.Peek >= 0
                If Reason <> "" Then Exit Do
                OneRow = srr.ReadLine
                If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈

                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, flag)
                    If colArray.Length < 7 Then
                        rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                        rst = False
                        Exit Do
                    End If
                    Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                    Dim sCName As String = "" & Trim(colArray(1).ToString) '品名
                    Dim sStandard As String = "" & Trim(colArray(2).ToString) '規格
                    Dim sUnit As String = "" & Trim(colArray(3).ToString) '單位
                    Dim iPrice As Integer = Val(colArray(4).ToString) '單價
                    Dim iAllCount As Integer = Val(colArray(5).ToString) '使用數量
                    Dim iTNum As Integer = Val(Me.TNum.Text)   '取得外部資料(不可為0)
                    Dim sPurpose As String = "" & Trim(colArray(6).ToString) '用途說明

                    Dim isubtotal As Integer = (iPrice * iAllCount) '小計
                    Dim ieachCost As Integer = 0
                    If iTNum > 0 Then
                        ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum)) '每人分攤費用
                    Else
                        ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / 1)) '每人分攤費用
                    End If

                    'If Reason = "" Then Reason += CheckImportData(colArray) '檢查資料正確性
                    '檢查資料正確性
                    If Reason = "" Then

                        If dt.Rows.Count > 0 Then
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                        Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If iTNum = 0 Then
                            'CHECK:1
                            Reason += "請先輸入訓練人數，不可為0" & vbCrLf
                        End If

                        Dim HaveCostID04 As Boolean = False '材料費
                        HaveCostID04 = False '查詢是否有材料費項目
                        If Not Session(cst_CostItemTable) Is Nothing Then
                            Dim dt2 As DataTable
                            dt2 = Session(cst_CostItemTable)
                            For i As Int16 = 0 To dt2.Rows.Count - 1
                                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                                    AndAlso dt2.Select("CostID='04'").Length > 0 Then '已刪除者不可做更動
                                    HaveCostID04 = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not HaveCostID04 Then
                            Reason += "訓練費用中不含 材料費項目，不可新增 材料品名項目表" & vbCrLf
                        End If

                        If Reason = "" Then
                            If iItemNo = 0 Then
                                Reason += "請輸入項次，不可為0" & vbCrLf 'int
                            End If
                            If sCName = "" Then
                                Reason += "請輸入品名" & vbCrLf
                            End If
                            If sStandard = "" Then
                                Reason += "請輸入規格" & vbCrLf
                            End If
                            If sUnit = "" Then
                                Reason += "請輸入單位" & vbCrLf
                            End If
                            If iPrice = 0 Then
                                Reason += "請輸入單價，不可為0" & vbCrLf 'int
                            End If
                            If iAllCount = 0 Then
                                Reason += "請輸入使用數量，不可為0" & vbCrLf 'int
                            End If
                            If sPurpose = "" Then
                                Reason += "請輸入用途說明" & vbCrLf
                            End If
                            If ieachCost < 1 Then
                                Reason += "計算後每人分攤費用，不可小於1" & vbCrLf 'int
                            End If
                        End If
                    End If

                    If Reason = "" Then
                        '產業人才投資方案專用
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("ItemNo") = iItemNo 'int
                        dr("CName") = sCName '30
                        dr("Standard") = sStandard '300
                        dr("Unit") = sUnit '30
                        dr("Price") = iPrice 'int
                        dr("AllCount") = iAllCount 'int
                        dr("TNum") = iTNum 'int
                        dr("subtotal") = (iPrice * iAllCount) '顯示重算
                        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
                        dr("PurPose") = sPurpose '300
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now

                        Session(Cst_CommonCostTable) = dt

                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            sr.Close()
            srr.Close()
            IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        Else
            Reason += "請選擇匯入的檔案" & vbCrLf
        End If

        If Reason <> "" Then
            rErrmsg = Reason
            rst = False
        Else
            rst = True
        End If
        Return rst
    End Function

    '計算 一人份材料明細 + 共同材料明細 =材料費用總計
    Sub ChglabTotal67()
        Me.labTotal67.Text = Val(Me.labTotal6.Text) + Val(Me.labTotal7.Text)
        'Me.labTotal67.Text = Val(Me.labTotal6.Text) + Val(Me.labTotal7.Text) + Val(Me.labTotal8.Text) + Val(Me.labTotal9.Text)
    End Sub

    '修改 教材費用
    Function chkdg8(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Const cst_title1 As String = "教材費用"
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo8")
        Dim eCName As TextBox = e.Item.FindControl("eCName8")
        Dim eStandards As TextBox = e.Item.FindControl("eStandards8")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit8")
        Dim ePrice As TextBox = e.Item.FindControl("ePrice8")

        Dim eAllCount As TextBox = e.Item.FindControl("eAllCount8")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum8") '訓練人數
        'Dim eTotal As TextBox = e.Item.FindControl("eTotal8")
        Dim esubtotal As TextBox = e.Item.FindControl("esubtotal8") '小計 = val(ePrice6.text)*val(ePerCount6.text)* val(eTNum6.text)
        'Dim eeachCost As TextBox = e.Item.FindControl("eeachCost8")'每人分攤費用　
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose8")

        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & Trim(eCName.Text)
        Dim sStandards As String = "" & Trim(eStandards.Text)
        Dim sUnit As String = "" & Trim(eUnit.Text)
        Dim iPrice As Integer = Val(ePrice.Text)
        Dim iAllCount As Integer = Val(eAllCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
        Dim ieachCost As Integer = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
        Dim sPurpose As String = "" & Trim(ePurPose.Text)
        '取得外部資料
        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

        Dim Errmsg As String = ""
        Errmsg = ""

        Dim HaveCostID03 As Boolean = False '教材費
        HaveCostID03 = False '查詢是否有 教材費 項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='03'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID03 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID03 Then
            Errmsg &= "訓練費用中不含 教材費項目，不可新增 " & cst_title1 & vbCrLf
        End If

        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入品名" & vbCrLf
            End If
            If sStandards = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iAllCount = 0 Then
                Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
            'If ieachCost < 1 Then
            '    Errmsg &= "計算後每人分攤費用，不可小於1" & vbCrLf 'int
            'End If
            If Not TIMS.IsNumeric2(eItemNo.Text) Then
                Errmsg &= "項次格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(ePrice.Text) Then
                Errmsg &= "單價格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(eAllCount.Text) Then
                Errmsg &= "使用數量格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid8
                Dim dt As DataTable = Session(Cst_SheetCostTable)
                Dim sfilter As String = ""
                sfilter = ""
                sfilter &= "" & Cst_SheetCostpkName & "<>'" & Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) & "'"
                sfilter &= " AND ItemNo='" & eItemNo.Text & "'"
                'dt update
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                    Errmsg &= "[" & eItemNo.Text & "]該項次已在表格中" & vbCrLf
                End If
                dt = Nothing
            End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            rst = False            'Exit Function
        End If
        Return rst
    End Function

    '修改 其他費用
    Function chkdg9(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Const cst_title1 As String = "其他費用"
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo9")
        Dim eCName As TextBox = e.Item.FindControl("eCName9")
        Dim eStandards As TextBox = e.Item.FindControl("eStandards9")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit9")
        Dim ePrice As TextBox = e.Item.FindControl("ePrice9")
        Dim eAllCount As TextBox = e.Item.FindControl("eAllCount9")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum9") '訓練人數
        'Dim eTotal As TextBox = e.Item.FindControl("eTotal9")
        Dim esubtotal As TextBox = e.Item.FindControl("esubtotal9") '小計 = val(ePrice6.text)*val(ePerCount6.text)* val(eTNum6.text)
        'Dim eeachCost As TextBox = e.Item.FindControl("eeachCost9")'每人分攤費用　
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose9")

        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & Trim(eCName.Text)
        Dim sStandards As String = "" & Trim(eStandards.Text)
        Dim sUnit As String = "" & Trim(eUnit.Text)
        Dim iPrice As Integer = Val(ePrice.Text)
        Dim iAllCount As Integer = Val(eAllCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
        Dim ieachCost As Integer = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
        Dim sPurpose As String = "" & Trim(ePurPose.Text)
        '取得外部資料
        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

        Dim Errmsg As String = ""
        Errmsg = ""

        Dim HaveCostID11 As Boolean = False '其他費用
        HaveCostID11 = False '查詢是否有 其他費用 項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='11'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID11 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID11 Then
            Errmsg &= "訓練費用中不含 其他費用項目，不可新增 " & cst_title1 & vbCrLf
        End If

        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入項目" & vbCrLf
            End If
            If sStandards = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iAllCount = 0 Then
                Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
            'If ieachCost < 1 Then
            '    Errmsg &= "計算後每人分攤費用，不可小於1" & vbCrLf 'int
            'End If
            If Not TIMS.IsNumeric2(eItemNo.Text) Then
                Errmsg &= "項次格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(ePrice.Text) Then
                Errmsg &= "單價格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Not TIMS.IsNumeric2(eAllCount.Text) Then
                Errmsg &= "使用數量格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid9
                Dim dt As DataTable = Session(Cst_OtherCostTable)
                Dim sfilter As String = ""
                sfilter = ""
                sfilter &= "" & Cst_OtherCostpkName & "<>'" & Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) & "'"
                sfilter &= " AND ItemNo='" & eItemNo.Text & "'"
                'dt update
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                    Errmsg &= "[" & eItemNo.Text & "]該項次 已在表格中" & vbCrLf
                End If
                dt = Nothing
            End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            rst = False            'Exit Function
        End If
        Return rst
    End Function

    '匯入 教材明細
    Function File3_test(ByRef rErrmsg As String) As Boolean
        '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim rst As Boolean = False

        Const cst_title1 As String = "教材費用"
        Dim oFile As HtmlInputFile = File3
        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Dim Reason As String = "" '儲存錯誤的原因
        Const cst_flag As String = ","
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        Dim dt As DataTable
        If Session(Cst_SheetCostTable) Is Nothing Then
            dt = CreateSheetCost()
        Else
            '有資料
            dt = Session(Cst_SheetCostTable)
            dt.Columns(Cst_SheetCostpkName).AutoIncrement = True
            dt.Columns(Cst_SheetCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_SheetCostpkName).AutoIncrementStep = -1
        End If

        If oFile.Value <> "" Then
            '檢查檔案格式與大小----------   Start
            If oFile.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Return rst
                'Exit Function
            Else
                '取出檔案名稱
                MyFileName = Split(oFile.PostedFile.FileName, "\")((Split(oFile.PostedFile.FileName, "\")).Length - 1)
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    Return rst 'Exit Function
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If Not LCase(MyFileType) = "csv" Then
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        Return rst 'Exit Function
                    End If
                End If
            End If
            '檢查檔案格式與大小----------   End

            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            '上傳檔案
            oFile.PostedFile.SaveAs(Server.MapPath(Upload_Path & MyFileName))

            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(Server.MapPath(Upload_Path & MyFileName))
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0 '讀取行累計數
            Dim OneRow As String        'srr.ReadLine 一行一行的資料
            Dim colArray As Array

            Do While srr.Peek >= 0
                If Reason <> "" Then Exit Do
                OneRow = srr.ReadLine
                If Replace(OneRow, cst_flag, "") = "" Then Exit Do '若資料為空白行，則離開回圈

                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, cst_flag)
                    If colArray.Length < 7 Then
                        rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                        rst = False
                        Exit Do
                    End If
                    Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                    Dim sCName As String = "" & Trim(colArray(1).ToString) '項目
                    Dim sStandards As String = "" & Trim(colArray(2).ToString) '規格
                    Dim sUnit As String = "" & Trim(colArray(3).ToString) '單位
                    Dim iPrice As Integer = Val(colArray(4).ToString) '單價
                    Dim iAllCount As Integer = Val(colArray(5).ToString) '使用數量
                    Dim iTNum As Integer = Val(Me.TNum.Text)   '取得外部資料(不可為0)
                    Dim sPurpose As String = "" & Trim(colArray(6).ToString) '用途說明
                    Dim isubtotal As Integer = (iPrice * iAllCount) '小計
                    Dim ieachCost As Integer = 0
                    If iTNum > 0 Then
                        ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum)) '每人分攤費用
                    Else
                        ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / 1)) '每人分攤費用
                    End If

                    'If Reason = "" Then Reason += CheckImportData(colArray) '檢查資料正確性
                    '檢查資料正確性
                    If Reason = "" Then
                        If dt.Rows.Count > 0 Then
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                        Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If iTNum = 0 Then
                            'CHECK:1
                            Reason += "請先輸入訓練人數，不可為0" & vbCrLf
                        End If

                        Dim HaveCostID03 As Boolean = False '教材費
                        HaveCostID03 = False '查詢是否有 教材費 項目
                        If Not Session(cst_CostItemTable) Is Nothing Then
                            Dim dt2 As DataTable
                            dt2 = Session(cst_CostItemTable)
                            For i As Int16 = 0 To dt2.Rows.Count - 1
                                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                                    AndAlso dt2.Select("CostID='03'").Length > 0 Then '已刪除者不可做更動
                                    HaveCostID03 = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not HaveCostID03 Then
                            Reason += "訓練費用中不含 教材費項目，不可新增 " & cst_title1 & vbCrLf
                        End If

                        If Reason = "" Then
                            If iItemNo = 0 Then
                                Reason += "請輸入項次，不可為0" & vbCrLf 'int
                            End If
                            If sCName = "" Then
                                Reason += "請輸入品名" & vbCrLf
                            End If
                            If sStandards = "" Then
                                Reason += "請輸入規格" & vbCrLf
                            End If
                            If sUnit = "" Then
                                Reason += "請輸入單位" & vbCrLf
                            End If
                            If iPrice = 0 Then
                                Reason += "請輸入單價，不可為0" & vbCrLf 'int
                            End If
                            If iAllCount = 0 Then
                                Reason += "請輸入使用數量，不可為0" & vbCrLf 'int
                            End If
                            If sPurpose = "" Then
                                Reason += "請輸入用途說明" & vbCrLf
                            End If
                            'If ieachCost < 1 Then
                            '    Reason += "計算後每人分攤費用，不可小於1" & vbCrLf 'int
                            'End If
                        End If
                    End If

                    If Reason = "" Then
                        '產業人才投資方案專用
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("ItemNo") = iItemNo 'int
                        dr("CName") = sCName '30
                        dr("Standards") = sStandards '300
                        dr("Unit") = sUnit '30
                        dr("Price") = iPrice 'int
                        dr("AllCount") = iAllCount 'int
                        dr("TNum") = iTNum 'int
                        dr("subtotal") = (iPrice * iAllCount) '顯示重算
                        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
                        dr("PurPose") = sPurpose '300
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        Session(Cst_SheetCostTable) = dt
                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            sr.Close()
            srr.Close()
            IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        Else
            Reason += "請選擇匯入的檔案" & vbCrLf
        End If

        If Reason <> "" Then
            rErrmsg = Reason
            rst = False
        Else
            rst = True
        End If
        Return rst
    End Function

    '匯入 其他費用明細
    Function File4_test(ByRef rErrmsg As String) As Boolean
        '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim rst As Boolean = False

        Const cst_title1 As String = "其他費用明細"
        Dim oFile As HtmlInputFile = File4
        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Dim Reason As String = "" '儲存錯誤的原因
        Const cst_flag As String = ","
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        Dim dt As DataTable
        If Session(Cst_OtherCostTable) Is Nothing Then
            dt = CreateOtherCost()
        Else
            '有資料
            dt = Session(Cst_OtherCostTable)
            dt.Columns(Cst_OtherCostpkName).AutoIncrement = True
            dt.Columns(Cst_OtherCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_OtherCostpkName).AutoIncrementStep = -1
        End If

        If oFile.Value <> "" Then
            '檢查檔案格式與大小----------   Start
            If oFile.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Return rst 'Exit Function
            Else
                '取出檔案名稱
                MyFileName = Split(oFile.PostedFile.FileName, "\")((Split(oFile.PostedFile.FileName, "\")).Length - 1)
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    Return rst 'Exit Function
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If Not LCase(MyFileType) = "csv" Then
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        Return rst 'Exit Function
                    End If
                End If
            End If
            '檢查檔案格式與大小----------   End

            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            '上傳檔案
            oFile.PostedFile.SaveAs(Server.MapPath(Upload_Path & MyFileName))

            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(Server.MapPath(Upload_Path & MyFileName))
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0 '讀取行累計數
            Dim OneRow As String        'srr.ReadLine 一行一行的資料
            Dim colArray As Array

            Do While srr.Peek >= 0
                If Reason <> "" Then Exit Do
                OneRow = srr.ReadLine
                If Replace(OneRow, cst_flag, "") = "" Then Exit Do '若資料為空白行，則離開回圈

                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, cst_flag)
                    If colArray.Length < 7 Then
                        rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                        rst = False
                        Exit Do
                    End If
                    Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                    Dim sCName As String = "" & Trim(colArray(1).ToString) '品名
                    Dim sStandards As String = "" & Trim(colArray(2).ToString) '規格
                    Dim sUnit As String = "" & Trim(colArray(3).ToString) '單位
                    Dim iPrice As Integer = Val(colArray(4).ToString) '單價
                    Dim iAllCount As Integer = Val(colArray(5).ToString) '使用數量
                    Dim iTNum As Integer = Val(Me.TNum.Text)   '取得外部資料(不可為0)
                    Dim sPurpose As String = "" & Trim(colArray(6).ToString) '用途說明
                    Dim isubtotal As Integer = (iPrice * iAllCount) '小計
                    Dim ieachCost As Integer = 0
                    If iTNum > 0 Then
                        ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum)) '每人分攤費用
                    Else
                        ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / 1)) '每人分攤費用
                    End If

                    'If Reason = "" Then Reason += CheckImportData(colArray) '檢查資料正確性
                    '檢查資料正確性
                    If Reason = "" Then
                        If dt.Rows.Count > 0 Then
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                        Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If iTNum = 0 Then
                            'CHECK:1
                            Reason += "請先輸入訓練人數，不可為0" & vbCrLf
                        End If

                        Dim HaveCostID11 As Boolean = False '其他費用
                        HaveCostID11 = False '查詢是否有 其他費用 項目
                        If Not Session(cst_CostItemTable) Is Nothing Then
                            Dim dt2 As DataTable
                            dt2 = Session(cst_CostItemTable)
                            For i As Int16 = 0 To dt2.Rows.Count - 1
                                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                                    AndAlso dt2.Select("CostID='11'").Length > 0 Then '已刪除者不可做更動
                                    HaveCostID11 = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not HaveCostID11 Then
                            Reason += "訓練費用中不含 其他費用項目，不可新增 " & cst_title1 & vbCrLf
                        End If

                        If Reason = "" Then
                            If iItemNo = 0 Then
                                Reason += "請輸入項次，不可為0" & vbCrLf 'int
                            End If
                            If sCName = "" Then
                                Reason += "請輸入項目" & vbCrLf
                            End If
                            If sStandards = "" Then
                                Reason += "請輸入規格" & vbCrLf
                            End If
                            If sUnit = "" Then
                                Reason += "請輸入單位" & vbCrLf
                            End If
                            If iPrice = 0 Then
                                Reason += "請輸入單價，不可為0" & vbCrLf 'int
                            End If
                            If iAllCount = 0 Then
                                Reason += "請輸入使用數量，不可為0" & vbCrLf 'int
                            End If
                            If sPurpose = "" Then
                                Reason += "請輸入用途說明" & vbCrLf
                            End If
                            'If ieachCost < 1 Then
                            '    Reason += "計算後每人分攤費用，不可小於1" & vbCrLf 'int
                            'End If
                        End If
                    End If

                    If Reason = "" Then
                        '產業人才投資方案專用
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)

                        dr("ItemNo") = iItemNo 'int
                        dr("CName") = sCName '30
                        dr("Standards") = sStandards '300
                        dr("Unit") = sUnit '30
                        dr("Price") = iPrice 'int
                        dr("AllCount") = iAllCount 'int
                        dr("TNum") = iTNum 'int
                        dr("subtotal") = (iPrice * iAllCount) '顯示重算
                        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
                        dr("PurPose") = sPurpose '300
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        Session(Cst_OtherCostTable) = dt
                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            sr.Close()
            srr.Close()
            IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        Else
            Reason += "請選擇匯入的檔案" & vbCrLf
        End If

        If Reason <> "" Then
            rErrmsg = Reason
            rst = False
        Else
            rst = True
        End If
        Return rst
    End Function

    '建立 教材費用,教材明細
    Function CreateSheetCost() As DataTable
        Dim dt As DataTable
        Dim DGobj As DataGrid = Me.DataGrid8
        Const cst_sSupFd As String = ",0 subtotal, 0 eachCost" '補充欄位
        Dim sql As String = ""
        If Session(Cst_SheetCostTable) Is Nothing Then
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")

                sql = "SELECT Plan_SheetCost.*" & cst_sSupFd & " FROM Plan_SheetCost where PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                If Request(cst_ccopy) = "1" Then
                    'Copy機制
                    sql = "SELECT Plan_SheetCost.*" & cst_sSupFd & " FROM Plan_SheetCost where 1<>1"
                Else
                    '修改資料取得
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = ""
                    sql &= " SELECT Plan_SheetCost.*" & cst_sSupFd & " FROM Plan_SheetCost WHERE 1=1" & vbCrLf
                    sql &= " AND PlanID='" & PlanID_value & "'" & vbCrLf
                    sql &= " AND ComIDNO='" & ComIDNO_value & "'" & vbCrLf
                    sql &= " AND SeqNO='" & SeqNO_value & "'" & vbCrLf
                    sql &= " ORDER BY ItemNo" & vbCrLf
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            '有資料
            dt = Session(Cst_SheetCostTable)
        End If
        dt.Columns(Cst_SheetCostpkName).AutoIncrement = True
        dt.Columns(Cst_SheetCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_SheetCostpkName).AutoIncrementStep = -1
        Session(Cst_SheetCostTable) = dt
        With DGobj
            .Style.Item("display") = "none"
            If dt.Rows.Count > 0 Then
                .Style.Item("display") = ""
                .DataSource = dt
                .DataKeyField = Cst_SheetCostpkName
                .DataBind()
            End If
        End With
        Dim subtotal As Integer = 0
        subtotal = 0
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If Not dr.RowState = DataRowState.Deleted Then
                    Dim iPrice As Integer = Val(dr("Price"))
                    Dim iAllCount As Integer = Val(dr("AllCount"))
                    Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                    dr("subtotal") = (iPrice * iAllCount)  '小計
                    dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
                    subtotal += Val(dr("subtotal"))
                End If
            Next
        End If

        Me.labTotal8.Text = subtotal
        'Me.labTotal67.Text = Val(Me.labTotal6.Text) + Val(Me.labTotal7.Text)
        trlabTotal8.Visible = False
        If subtotal > 0 Then trlabTotal8.Visible = True

        Call ChglabTotal67()
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function

    '建立 其他費用明細
    Function CreateOtherCost() As DataTable
        Dim dt As DataTable
        Dim DGobj As DataGrid = Me.DataGrid9
        'Const cst_sSupFd As String = ",0 Total,0 subtotal" '補充欄位
        Const cst_sSupFd As String = ",0 subtotal, 0 eachCost" '補充欄位
        Dim sql As String = ""
        If Session(Cst_OtherCostTable) Is Nothing Then
            If Me.upt_PlanX.Value <> "" Then
                tmpPCS = Me.upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")

                'Copy機制
                sql = "SELECT Plan_OtherCost.*" & cst_sSupFd & " FROM Plan_OtherCost where PlanID='" & PlanID_value & "' and ComIDNO='" & ComIDNO_value & "' and SeqNO='" & SeqNO_value & "'"
            Else
                If Request(cst_ccopy) = "1" Then
                    'Copy機制
                    sql = "SELECT Plan_OtherCost.*" & cst_sSupFd & " FROM Plan_OtherCost where 1<>1"
                Else
                    '修改資料取得
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    sql = ""
                    sql &= " SELECT Plan_OtherCost.*" & cst_sSupFd & " FROM Plan_OtherCost WHERE 1=1" & vbCrLf
                    sql &= " AND PlanID='" & PlanID_value & "'" & vbCrLf
                    sql &= " AND ComIDNO='" & ComIDNO_value & "'" & vbCrLf
                    sql &= " AND SeqNO='" & SeqNO_value & "'" & vbCrLf
                    sql &= " ORDER BY ItemNo" & vbCrLf
                    If g_flagNG Then sql = " SELECT Plan_OtherCost.* " & cst_sSupFd & " FROM Plan_OtherCost WHERE 1<>1 "
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            '有資料
            dt = Session(Cst_OtherCostTable)
        End If
        dt.Columns(Cst_OtherCostpkName).AutoIncrement = True
        dt.Columns(Cst_OtherCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_OtherCostpkName).AutoIncrementStep = -1
        Session(Cst_OtherCostTable) = dt

        With DGobj
            .Style.Item("display") = "none"
            If dt.Rows.Count > 0 Then
                .Style.Item("display") = ""

                .DataSource = dt
                .DataKeyField = Cst_OtherCostpkName
                .DataBind()
            End If
        End With
        Dim subtotal As Integer = 0
        subtotal = 0
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If Not dr.RowState = DataRowState.Deleted Then
                    Dim iPrice As Integer = Val(dr("Price"))
                    Dim iAllCount As Integer = Val(dr("AllCount"))
                    Dim iTNum As Integer = Val(dr("TNum"))  '取得外部資料
                    dr("subtotal") = (iPrice * iAllCount)  '小計
                    dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用
                    subtotal += Val(dr("subtotal"))
                End If
            Next
        End If
        Me.labTotal9.Text = subtotal
        'Me.labTotal67.Text = Val(Me.labTotal6.Text) + Val(Me.labTotal7.Text)
        trlabTotal9.Visible = False
        If subtotal > 0 Then trlabTotal9.Visible = True

        Call ChglabTotal67()
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function

    '新增 教材費用
    Sub AddSheetCost()
        Dim dt As DataTable
        Dim Errmsg As String = ""
        Const cst_title1 As String = "教材費用"
        Dim iItemNo As Integer = Val(Me.tItemNo8.Text)
        Dim sCName As String = "" & Trim(Me.tCName8.Text)
        Dim sStandards As String = "" & Trim(Me.tStandards8.Text)
        Dim sUnit As String = "" & Trim(Me.tUnit8.Text)
        Dim iPrice As Integer = Val(Me.tPrice8.Text)
        Dim iAllCount As Integer = Val(Me.tAllCount8.Text)
        Dim iTNum As Integer = Val(Me.TNum.Text)   '取得外部資料
        Dim sPurpose As String = "" & Trim(Me.tPurPose8.Text)
        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
        Dim ieachCost As Integer = 0
        If iTNum > 0 Then
            ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum)) '每人分攤費用
        Else
            ieachCost = TIMS.ROUND(Val((iPrice * iAllCount) / 1)) '每人分攤費用
        End If

        If Session(Cst_SheetCostTable) Is Nothing Then
            dt = CreateSheetCost()
        Else
            '有資料
            dt = Session(Cst_SheetCostTable)
            dt.Columns(Cst_SheetCostpkName).AutoIncrement = True
            dt.Columns(Cst_SheetCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_SheetCostpkName).AutoIncrementStep = -1
        End If

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        If iTNum = 0 Then
            'CHECK:1
            Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf
        End If

        Dim HaveCostID03 As Boolean = False '教材費
        HaveCostID03 = False '查詢是否有 教材費 項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='03'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID03 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID03 Then
            Errmsg &= "訓練費用中不含 教材費項目，不可新增 " & cst_title1 & vbCrLf
        End If

        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入品名" & vbCrLf
            End If
            If sStandards = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iAllCount = 0 Then
                Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
            'If ieachCost < 1 Then
            '    Errmsg &= "計算後每人分攤費用，不可小於1" & vbCrLf 'int
            'End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standards") = sStandards '300
        dr("Unit") = sUnit '30
        dr("Price") = iPrice 'int
        dr("AllCount") = iAllCount 'int
        dr("TNum") = iTNum 'int
        dr("subtotal") = (iPrice * iAllCount) '顯示重算
        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(Cst_SheetCostTable) = dt
        Call CreateSheetCost()

        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    '新增 其他費用明細
    Sub AddOtherCost()
        Dim dt As DataTable
        Dim Errmsg As String = ""
        Const cst_title1 As String = "其他費用明細"
        Dim iItemNo As Integer = Val(Me.tItemNo9.Text)
        Dim sCName As String = "" & Trim(Me.tCName9.Text)
        Dim sStandards As String = "" & Trim(Me.tStandards9.Text)
        Dim sUnit As String = "" & Trim(Me.tUnit9.Text)
        Dim iPrice As Integer = Val(Me.tPrice9.Text)
        Dim iAllCount As Integer = Val(Me.tAllCount9.Text)
        Dim iTNum As Integer = Val(Me.TNum.Text)   '取得外部資料
        Dim sPurpose As String = "" & Trim(Me.tPurpose9.Text)
        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
        Dim ieachCost As Integer = 0

        If Session(Cst_OtherCostTable) Is Nothing Then
            dt = CreateOtherCost()
        Else
            '有資料
            dt = Session(Cst_OtherCostTable)
            dt.Columns(Cst_OtherCostpkName).AutoIncrement = True
            dt.Columns(Cst_OtherCostpkName).AutoIncrementSeed = -1
            dt.Columns(Cst_OtherCostpkName).AutoIncrementStep = -1
        End If

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        If iTNum = 0 Then
            'CHECK:1
            Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf
        End If

        Dim HaveCostID11 As Boolean = False '其他費用
        HaveCostID11 = False '查詢是否有 其他費用 項目
        If Not Session(cst_CostItemTable) Is Nothing Then
            Dim dt2 As DataTable
            dt2 = Session(cst_CostItemTable)
            For i As Int16 = 0 To dt2.Rows.Count - 1
                If Not dt2.Rows(i).RowState = DataRowState.Deleted _
                    AndAlso dt2.Select("CostID='11'").Length > 0 Then '已刪除者不可做更動
                    HaveCostID11 = True
                    Exit For
                End If
            Next
        End If
        If Not HaveCostID11 Then
            Errmsg &= "訓練費用中不含 其他費用項目，不可新增 " & cst_title1 & vbCrLf
        End If

        If Errmsg = "" Then
            If iItemNo = 0 Then
                Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            End If
            If sCName = "" Then
                Errmsg &= "請輸入項目" & vbCrLf
            End If
            If sStandards = "" Then
                Errmsg &= "請輸入規格" & vbCrLf
            End If
            If sUnit = "" Then
                Errmsg &= "請輸入單位" & vbCrLf
            End If
            If iPrice = 0 Then
                Errmsg &= "請輸入單價，不可為0" & vbCrLf 'int
            End If
            If iAllCount = 0 Then
                Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            End If
            If sPurpose = "" Then
                Errmsg &= "請輸入用途說明" & vbCrLf
            End If
            'If ieachCost < 1 Then
            '    Errmsg &= "計算後每人分攤費用，不可小於1" & vbCrLf 'int
            'End If
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standards") = sStandards '300
        dr("Unit") = sUnit '30
        dr("Price") = iPrice 'int
        dr("AllCount") = iAllCount 'int
        dr("TNum") = iTNum 'int
        dr("subtotal") = (iPrice * iAllCount) '顯示重算
        dr("eachCost") = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(Cst_OtherCostTable) = dt
        Call CreateOtherCost()

        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    '匯入'PLAN_SHEETCOST–教材費用
    Protected Sub BtnImport8_Click(sender As Object, e As EventArgs) Handles BtnImport8.Click
        Dim Errmsg As String = ""
        If File3_test(Errmsg) Then
            '顯示 內容
            Call CreateSheetCost()
        Else
            Common.MessageBox(Me, Errmsg)
        End If
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    '新增 'PLAN_SHEETCOST–教材費用
    Protected Sub btnAddCost8_Click(sender As Object, e As EventArgs) Handles btnAddCost8.Click
        Call AddSheetCost()
    End Sub

    Private Sub DataGrid8_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid8.ItemCommand
        Dim Errmsg As String = ""
        If Session(Cst_SheetCostTable) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid8
        Const cst_eCmdEDT As String = "EDT8"
        Const cst_eCmdDEL As String = "DEL8"
        Const cst_eCmdUPD As String = "UPD8"
        Const cst_eCmdCLS As String = "CLS8"

        Dim dt As DataTable = Session(Cst_SheetCostTable)
        Dim dr As DataRow = Nothing
        Errmsg = ""
        Select Case e.CommandName
            Case cst_eCmdEDT '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case cst_eCmdDEL '刪除
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If
                Dim sfilter As String = "" & Cst_SheetCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                   AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then
                            dr.Delete() '刪除
                        End If
                    Next
                End If
            Case cst_eCmdUPD '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo8")
                Dim eCName As TextBox = e.Item.FindControl("eCName8")
                Dim eStandards As TextBox = e.Item.FindControl("eStandards8")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit8")
                Dim ePrice As TextBox = e.Item.FindControl("ePrice8")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount8")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum8") '訓練人數
                'Dim eTotal As TextBox = e.Item.FindControl("eTotal8")
                Dim esubtotal As TextBox = e.Item.FindControl("esubtotal8") '小計 = val(ePrice6.text)*val(ePerCount6.text)* val(eTNum6.text)
                'Dim eeachCost As TextBox = e.Item.FindControl("eeachCost8")'每人分攤費用　
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose8")

                If chkdg8(e) Then
                    Dim sfilter As String = "" & Cst_SheetCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    'dt update
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                        AndAlso dt.Select(sfilter).Length > 0 Then

                        Dim iPrice As Integer = Val(ePrice.Text)
                        Dim iAllCount As Integer = Val(eAllCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
                        '取得外部資料
                        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)
                        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
                        Dim ieachCost As Integer = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用

                        dr = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standards") = eStandards.Text
                        dr("Unit") = eUnit.Text
                        dr("Price") = iPrice
                        dr("AllCount") = iAllCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("subtotal") = isubtotal '小計 '顯示重算
                        dr("eachCost") = ieachCost  '每人分攤費用 '顯示重算
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If

            Case cst_eCmdCLS '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select

        Session(Cst_SheetCostpkName) = dt  '要新  
        CreateSheetCost() '建立  
        'Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")

    End Sub

    Private Sub DataGrid8_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid8.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost8.Enabled

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                Dim lItemNo As Label = e.Item.FindControl("lItemNo8")
                Dim lCName As Label = e.Item.FindControl("lCName8")
                Dim lStandards As Label = e.Item.FindControl("lStandards8")
                Dim lUnit As Label = e.Item.FindControl("lUnit8")
                Dim lPrice As Label = e.Item.FindControl("lPrice8")
                Dim lAllCount As Label = e.Item.FindControl("lAllCount8")
                Dim lTNum As Label = e.Item.FindControl("lTNum8")
                'Dim lTotal As Label = e.Item.FindControl("lTotal8")
                Dim lsubtotal As Label = e.Item.FindControl("lsubtotal8")
                Dim leachCost As Label = e.Item.FindControl("leachCost8")
                Dim lPurPose As Label = e.Item.FindControl("lPurPose8")

                Dim btnDel8 As Button = e.Item.FindControl("btnDel8") '刪除
                Dim btnEdt8 As Button = e.Item.FindControl("btnEdt8") '修改
                'Dim btnUpd8 As Button = e.Item.FindControl("btnUpd8")
                'Dim btnCls8 As Button = e.Item.FindControl("btnCls8")

                lItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName.Text = "" & Convert.ToString(drv("CName"))
                lStandards.Text = "" & Convert.ToString(drv("Standards"))
                lUnit.Text = "" & Convert.ToString(drv("Unit"))
                'lPrice.Text = "" & Convert.ToString(drv("Price"))
                'lAllCount.Text = "" & Convert.ToString(drv("AllCount"))
                'lTNum.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                'lsubtotal.Text = "" & (Val(drv("Price")) * Val(drv("AllCount")))  '顯示重算
                'leachCost.Text = "" & ((Val(drv("Price")) * Val(drv("AllCount"))) / Val(drv("TNum")))  '顯示重算
                Dim iPrice As Integer = Val(drv("Price"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                '取得外部資料
                If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

                lPrice.Text = iPrice
                lAllCount.Text = iAllCount
                lTNum.Text = iTNum '顯示原資料
                lsubtotal.Text = (iPrice * iAllCount)  '小計  '顯示重算
                leachCost.Text = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算

                lPurPose.Text = "" & Convert.ToString(drv("PurPose"))

                btnDel8.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDel8.Enabled = Flag_AddEnabled
                btnEdt8.Enabled = Flag_AddEnabled

            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim tlItemNo As TextBox = e.Item.FindControl("eItemNo8")
                Dim tlCName As TextBox = e.Item.FindControl("eCName8")
                Dim tlStandards As TextBox = e.Item.FindControl("eStandards8")
                Dim tlUnit As TextBox = e.Item.FindControl("eUnit8")
                Dim tlPrice As TextBox = e.Item.FindControl("ePrice8")
                Dim tlAllCount As TextBox = e.Item.FindControl("eAllCount8")
                Dim tlTNum As TextBox = e.Item.FindControl("eTNum8")
                'Dim tTotal As TextBox = e.Item.FindControl("eTotal8")
                Dim tlsubtotal As TextBox = e.Item.FindControl("esubtotal8")
                Dim tleachCost As TextBox = e.Item.FindControl("eeachCost8")
                Dim tlPurPose As TextBox = e.Item.FindControl("ePurPose8")

                Dim btnUpd8 As Button = e.Item.FindControl("btnUpd8") '更新
                Dim btnCls8 As Button = e.Item.FindControl("btnCls8") '取消

                tlItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                tlCName.Text = "" & Convert.ToString(drv("CName"))
                tlStandards.Text = "" & Convert.ToString(drv("Standards"))
                tlUnit.Text = "" & Convert.ToString(drv("Unit"))
                'dr("TNum") = Val(eTNum6.Text)
                'dr("Total") = Val(ePerCount6.Text) * Val(eTNum6.Text)
                'dr("subtotal") = Val(ePrice6.Text) * Val(ePerCount6.Text) * Val(eTNum6.Text)
                'tlTNum.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                'tlsubtotal.Text = "" & (Val(drv("Price")) * Val(drv("AllCount")))  '顯示重算
                'tleachCost.Text = "" & ((Val(drv("Price")) * Val(drv("AllCount"))) / Val(drv("TNum")))  '顯示重算
                Dim iPrice As Integer = Val(drv("Price"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                '取得外部資料
                If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

                tlPrice.Text = iPrice
                tlAllCount.Text = iAllCount
                tlTNum.Text = iTNum '顯示原資料
                tlsubtotal.Text = (iPrice * iAllCount)  '小計  '顯示重算
                tleachCost.Text = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
                tlPurPose.Text = "" & Convert.ToString(drv("PurPose"))

                tlTNum.ReadOnly = True
                tlsubtotal.ReadOnly = True
                tleachCost.ReadOnly = True
                'tlTNum6.Style.Item("background-color") = "#FFECEC"
                'tlTotal6.Style.Item("background-color") = "#FFECEC"
                'tlsubtotal6.Style.Item("background-color") = "#FFECEC"
                tlTNum.Style.Item("background-color") = "#BDBDBD"
                tlsubtotal.Style.Item("background-color") = "#BDBDBD"
                tleachCost.Style.Item("background-color") = "#BDBDBD"

                btnUpd8.Enabled = Flag_AddEnabled
                btnCls8.Enabled = True
        End Select
    End Sub

    '匯入'PLAN_OtherCOST–其他費用
    Protected Sub BtnImport9_Click(sender As Object, e As EventArgs) Handles BtnImport9.Click
        Dim Errmsg As String = ""
        If File4_test(Errmsg) Then
            '顯示 內容
            Call CreateOtherCost()
        Else
            Common.MessageBox(Me, Errmsg)
        End If
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")
    End Sub

    '新增 'PLAN_OtherCOST–其他費用
    Protected Sub btnAddCost9_Click(sender As Object, e As EventArgs) Handles btnAddCost9.Click
        Call AddOtherCost()
    End Sub

    Private Sub DataGrid9_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid9.ItemCommand
        Dim Errmsg As String = ""
        If Session(Cst_OtherCostTable) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid9
        Const cst_eCmdEDT As String = "EDT9"
        Const cst_eCmdDEL As String = "DEL9"
        Const cst_eCmdUPD As String = "UPD9"
        Const cst_eCmdCLS As String = "CLS9"

        Dim dt As DataTable = Session(Cst_OtherCostTable)
        Dim dr As DataRow = Nothing
        Errmsg = ""
        Select Case e.CommandName
            Case cst_eCmdEDT '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case cst_eCmdDEL '刪除
                If DGobj Is Nothing _
                    OrElse dt Is Nothing Then
                    Common.MessageBox(Me, cst_errmsg16)
                    Exit Sub
                End If

                Dim sfilter As String = "" & Cst_OtherCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                   AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then
                            dr.Delete() '刪除
                        End If
                    Next
                End If
            Case cst_eCmdUPD '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo9")
                Dim eCName As TextBox = e.Item.FindControl("eCName9")
                Dim eStandards As TextBox = e.Item.FindControl("eStandards9")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit9")
                Dim ePrice As TextBox = e.Item.FindControl("ePrice9")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount9")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum9") '訓練人數
                'Dim eTotal As TextBox = e.Item.FindControl("eTotal9")
                Dim esubtotal As TextBox = e.Item.FindControl("esubtotal9") '小計 = val(ePrice6.text)*val(ePerCount6.text)* val(eTNum6.text)
                'Dim eeachCost As TextBox = e.Item.FindControl("eeachCost9")'每人分攤費用　
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose9")

                If chkdg9(e) Then
                    Dim sfilter As String = "" & Cst_OtherCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    'dt update
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" _
                                    AndAlso dt.Select(sfilter).Length > 0 Then
                        Dim iPrice As Integer = Val(ePrice.Text)
                        Dim iAllCount As Integer = Val(eAllCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
                        '取得外部資料
                        If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)
                        Dim isubtotal As Integer = (iPrice * iAllCount) '小計
                        Dim ieachCost As Integer = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用

                        dr = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standards") = eStandards.Text
                        dr("Unit") = eUnit.Text
                        dr("Price") = iPrice
                        dr("AllCount") = iAllCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("subtotal") = isubtotal '小計 '顯示重算
                        dr("eachCost") = ieachCost  '每人分攤費用 '顯示重算
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If
            Case cst_eCmdCLS '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select

        Session(Cst_OtherCostTable) = dt  '要新  
        CreateOtherCost() '建立  
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);ShowItemCostName('CostID2','ItemCostName','Itemage');</script>")

    End Sub

    Private Sub DataGrid9_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid9.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost6.Enabled
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                Dim lItemNo As Label = e.Item.FindControl("lItemNo9")
                Dim lCName As Label = e.Item.FindControl("lCName9")
                Dim lStandards As Label = e.Item.FindControl("lStandards9")
                Dim lUnit As Label = e.Item.FindControl("lUnit9")
                Dim lPrice As Label = e.Item.FindControl("lPrice9")
                Dim lAllCount As Label = e.Item.FindControl("lAllCount9")
                Dim lTNum As Label = e.Item.FindControl("lTNum9")
                'Dim lTotal As Label = e.Item.FindControl("lTotal9")
                Dim lsubtotal As Label = e.Item.FindControl("lsubtotal9")
                Dim leachCost As Label = e.Item.FindControl("leachCost9")
                Dim lPurPose As Label = e.Item.FindControl("lPurPose9")

                Dim btnDel9 As Button = e.Item.FindControl("btnDel9") '刪除
                Dim btnEdt9 As Button = e.Item.FindControl("btnEdt9") '修改

                lItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName.Text = "" & Convert.ToString(drv("CName"))
                lStandards.Text = "" & Convert.ToString(drv("Standards"))
                lUnit.Text = "" & Convert.ToString(drv("Unit"))

                Dim iPrice As Integer = Val(drv("Price"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                '取得外部資料
                If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

                lPrice.Text = iPrice
                lAllCount.Text = iAllCount
                lTNum.Text = iTNum '顯示原資料
                lsubtotal.Text = (iPrice * iAllCount)  '小計  '顯示重算
                leachCost.Text = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算

                lPurPose.Text = "" & Convert.ToString(drv("PurPose"))

                btnDel9.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDel9.Enabled = Flag_AddEnabled
                btnDel9.Enabled = Flag_AddEnabled

            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo9")
                Dim eCName As TextBox = e.Item.FindControl("eCName9")
                Dim eStandards As TextBox = e.Item.FindControl("eStandards9")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit9")
                Dim ePrice As TextBox = e.Item.FindControl("ePrice9")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount9")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum9")
                'Dim eotal As TextBox = e.Item.FindControl("eTotal9")
                Dim esubtotal As TextBox = e.Item.FindControl("esubtotal9")
                Dim eeachCost As TextBox = e.Item.FindControl("eeachCost9")
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose9")

                Dim btnUpd9 As Button = e.Item.FindControl("btnUpd9") '更新
                Dim btnCls9 As Button = e.Item.FindControl("btnCls9") '取消

                eItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                eCName.Text = "" & Convert.ToString(drv("CName"))
                eStandards.Text = "" & Convert.ToString(drv("Standards"))
                eUnit.Text = "" & Convert.ToString(drv("Unit"))
                'ePrice.Text = "" & Convert.ToString(drv("Price"))

                Dim iPrice As Integer = Val(drv("Price"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                '取得外部資料
                If iTNum <> Val(Me.TNum.Text) Then iTNum = Val(Me.TNum.Text)

                ePrice.Text = iPrice
                eAllCount.Text = iAllCount
                eTNum.Text = iTNum '顯示原資料
                esubtotal.Text = (iPrice * iAllCount)  '小計  '顯示重算
                eeachCost.Text = TIMS.ROUND(Val((iPrice * iAllCount) / iTNum))  '每人分攤費用  '顯示重算
                ePurPose.Text = "" & Convert.ToString(drv("PurPose"))

                eTNum.ReadOnly = True
                esubtotal.ReadOnly = True
                eeachCost.ReadOnly = True
                'tlTNum6.Style.Item("background-color") = "#FFECEC"
                'tlTotal6.Style.Item("background-color") = "#FFECEC"
                'tlsubtotal6.Style.Item("background-color") = "#FFECEC"
                eTNum.Style.Item("background-color") = "#BDBDBD"
                esubtotal.Style.Item("background-color") = "#BDBDBD"
                eeachCost.Style.Item("background-color") = "#BDBDBD"

                btnUpd9.Enabled = Flag_AddEnabled
                btnCls9.Enabled = True
        End Select
    End Sub

    Sub sUtl_PageInit1()
        'Dim iMaxLength As Integer = 0
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("PLAN_PLANINFO,PLAN_ONCLASS", objconn) ' DbAccess.GetDataTable(sql)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "ENTERPRISENAME", EnterpriseName) '企業包班名稱
        Call TIMS.sUtl_SetMaxLen(dt, "PLANEMAIL", EMail) 'EMAIL
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER1", Other1) '其他一
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER2", Other2) '其他二
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER3", Other3) '其他三
        Call TIMS.sUtl_SetMaxLen(dt, "TMSCIENCE", TMScience) '訓練方式
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSNAME", ClassName, -15) '班別名稱
        Call TIMS.sUtl_SetMaxLen(dt, "CYCLTYPE", CyclType) '期別
        Call TIMS.sUtl_SetMaxLen(dt, "ROOMNAME", RoomName) '上課教室名稱
        Call TIMS.sUtl_SetMaxLen(dt, "FACTMODEOTHER", FactModeOther) '場地類型其他說明
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTNAME", ContactName) '聯絡人
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTPHONE", ContactPhone) '電話
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTEMAIL", ContactEmail) '電子郵件
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTFAX", ContactFax) '傳真
        Call TIMS.sUtl_SetMaxLen(dt, "TIMES", Times) '時間
    End Sub

    Protected Sub DataGrid2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid2.SelectedIndexChanged

    End Sub
End Class